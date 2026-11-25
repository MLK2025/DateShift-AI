import JSZip from "jszip";
import type { PDFDocumentProxy, PDFPageProxy } from "pdfjs-dist";
import { openPdf } from "./document-text";
import { hasTransactionColumns } from "./ledger-docx";
import { imageViewport, positionPdfPage, renderPdfCanvas, type PositionedPage, type PositionedText } from "./pdf-word-positioned";

export type WordConversionMode = "editable" | "reflow" | "layout";
export type ConvertedWord = {
  blob: Blob;
  fileName: string;
  pageCount: number;
  imagePages: number;
};

const wordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const officeRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const packageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";
const wordMime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const xmlHeader = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

type TextCell = { left: number; right: number; text: string };
type TextLine = { y: number; size: number; cells: TextCell[] };
type PageSize = { width: number; height: number; margin: number };

async function collectLines(page: PDFPageProxy): Promise<TextLine[]> {
  const viewport = page.getViewport({ scale: 1 });
  const content = await page.getTextContent();
  const runs = content.items.flatMap((item) => {
    if (!("str" in item) || !item.str.trim()) return [];
    const [left, y] = viewport.convertToViewportPoint(item.transform[4], item.transform[5]);
    const size = Math.max(item.height || Math.hypot(item.transform[2], item.transform[3]), 1);
    return [{ left, right: left + item.width, y, size, text: item.str }];
  }).sort((a, b) => a.y - b.y || a.left - b.left);

  const lines: TextLine[] = [];
  for (const run of runs) {
    const line = lines.at(-1);
    if (line && Math.abs(run.y - line.y) <= Math.max(2, Math.min(line.size, run.size) * .25)) {
      line.cells.push(run);
      line.size = Math.max(line.size, run.size);
    } else {
      lines.push({ y: run.y, size: run.size, cells: [run] });
    }
  }
  for (const line of lines) {
    const cells: TextCell[] = [];
    for (const run of line.cells.sort((a, b) => a.left - b.left)) {
      const previous = cells.at(-1);
      const gap = run.left - (previous?.right ?? 0);
      if (previous && gap < Math.max(16, line.size * 1.65)) {
        const addSpace = /\s$/.test(previous.text) || /^\s/.test(run.text) || gap < line.size * .12 ? "" : " ";
        previous.text += addSpace + run.text;
        previous.right = Math.max(previous.right, run.right);
      } else {
        cells.push({ left: run.left, right: run.right, text: run.text });
      }
    }
    line.cells = cells;
  }
  return lines;
}

function alignedRows(lines: TextLine[], start: number) {
  const header = lines[start];
  const columns = header.cells;
  if (columns.length < 2) return [];
  const rows = [columns.map((cell) => cell.text.trim())];
  const ledger = /^\s*(?:(?:transaction|posting|booking) )?date\b/i.test(columns[0].text)
    && hasTransactionColumns(rows[0].join(" "));
  const boundaries = columns.slice(1).map((cell, index) => (columns[index].right + cell.left) / 2);

  for (let i = start + 1; i < lines.length; i += 1) {
    const line = lines[i];
    if (line.y - lines[i - 1].y > Math.max(24, header.size * 2.6) || line.cells.length < 2) break;
    const values = Array<string>(columns.length).fill("");
    if (ledger) {
      for (const cell of line.cells) {
        const next = boundaries.findIndex((boundary) => (cell.left + cell.right) / 2 < boundary);
        const col = next < 0 ? columns.length - 1 : next;
        values[col] = [values[col], cell.text.trim()].filter(Boolean).join(" ");
      }
      if (!/\d/.test(values[0]) || !/\d/.test(values.at(-1) ?? "")) break;
    } else {
      if (line.cells.length !== columns.length || line.cells.some((cell, index) =>
        Math.abs(cell.left - columns[index].left) > 14 && Math.abs(cell.right - columns[index].right) > 14,
      )) break;
      line.cells.forEach((cell, index) => { values[index] = cell.text.trim(); });
    }
    rows.push(values);
  }
  return rows.length >= (ledger ? 2 : 3) ? rows : [];
}

function escapeWordText(text: string) {
  return text.replace(/\p{Cc}|[\uD800-\uDFFF\uFFFE\uFFFF]/gu, "")
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function textParagraph(text: string, size: number, after = 0, indent = 0) {
  const halfPoints = Math.max(12, Math.min(192, Math.round(size * 2)));
  return `<w:p><w:pPr><w:spacing w:after="${after}"/><w:ind w:left="${indent}"/></w:pPr><w:r><w:rPr><w:sz w:val="${halfPoints}"/><w:szCs w:val="${halfPoints}"/></w:rPr><w:t xml:space="preserve">${escapeWordText(text)}</w:t></w:r></w:p>`;
}

function editablePage(lines: TextLine[], page: PageSize) {
  const availableWidth = page.width - 2 * page.margin;
  const leftEdge = Math.min(...lines.map((line) => line.cells[0].left));
  const body: string[] = [];
  for (let i = 0; i < lines.length;) {
    const line = lines[i];
    const rows = alignedRows(lines, i);
    if (rows.length) {
      const columns = line.cells;
      const left = columns[0].left;
      const right = Math.max(...lines.slice(i, i + rows.length).map((row) => row.cells.at(-1)!.right), columns.at(-1)!.left + line.size * 4) + 12;
      const width = Math.min(availableWidth, Math.round((right - left) * 20));
      const widths = columns.map((cell, column) => {
        const next = columns[column + 1]?.left ?? right;
        return Math.round(width * (next - cell.left) / (right - left));
      });
      const grid = widths.map((w) => `<w:gridCol w:w="${w}"/>`).join("");
      const cells = rows.map((row) => `<w:tr>${row.map((value, col) => `<w:tc><w:tcPr><w:tcW w:w="${widths[col]}" w:type="dxa"/></w:tcPr>${textParagraph(value, line.size)}</w:tc>`).join("")}</w:tr>`).join("");
      body.push(`<w:tbl><w:tblPr><w:tblW w:w="${width}" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid>${grid}</w:tblGrid>${cells}</w:tbl>`);
      i += rows.length;
    } else {
      const nextY = lines[i + 1]?.y ?? line.y;
      const after = Math.round(Math.max(0, Math.min(24, nextY - line.y - line.size * 1.18)) * 20);
      const indent = Math.min(Math.round((line.cells[0].left - leftEdge) * 20), Math.round(availableWidth * .65));
      body.push(textParagraph(line.cells.map((cell) => cell.text.trim()).join("  "), line.size, after, indent));
      i += 1;
    }
  }
  return body.join("");
}

async function renderPageImage(page: PDFPageProxy, signal?: AbortSignal) {
  const { canvas } = await renderPdfCanvas(page, imageViewport(page), signal);
  try {
    const blob = await new Promise<Blob>((resolve, reject) => canvas.toBlob((image) => {
      if (image) resolve(image);
      else reject(new Error("A PDF page could not be saved as an image."));
    }, "image/jpeg", .92));
    return blob.arrayBuffer();
  } finally {
    canvas.width = 0;
    canvas.height = 0;
  }
}

function pictureGraphic(pageNumber: number, width: number, height: number) {
  return `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${pageNumber}" name="PDF page ${pageNumber}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="page${pageNumber}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${width}" cy="${height}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>`;
}

function pageDrawing(pageNumber: number, size: PageSize, originalWidth: number, originalHeight: number) {
  const unitsPerTwip = 635;
  // Reserve a few points for Word's paragraph baseline and section break.
  const scale = Math.min(size.width / originalWidth, (size.height - 100) / originalHeight);
  const width = Math.round(originalWidth * scale * unitsPerTwip);
  const height = Math.round(originalHeight * scale * unitsPerTwip);
  const description = `PDF page ${pageNumber}`;
  return `<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="0"/><w:rPr><w:sz w:val="2"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="2"/></w:rPr><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${width}" cy="${height}"/><wp:docPr id="${pageNumber}" name="${description}" descr="${description}"/>${pictureGraphic(pageNumber, width, height)}</wp:inline></w:drawing></w:r></w:p>`;
}

const wordShapeNamespace = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
const emu = (points: number) => Math.round(points * 12700);

function anchoredDrawing(id: number, name: string, box: { x: number; y: number; width: number; height: number }, behindText: boolean, graphic: string) {
  return `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${id}" behindDoc="${behindText ? 1 : 0}" locked="0" layoutInCell="0" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>${emu(box.x)}</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>${emu(box.y)}</wp:posOffset></wp:positionV><wp:extent cx="${emu(box.width)}" cy="${emu(box.height)}"/><wp:wrapNone/><wp:docPr id="${id}" name="${name}"/>${graphic}</wp:anchor></w:drawing></w:r>`;
}

function textBoxGraphic(run: PositionedText, box: { width: number; height: number }) {
  const direction = run.rotation < 0 ? "vert270" : run.rotation > 0 ? "vert" : "horz";
  const halfPoints = Math.round(run.size * 2);
  const width = emu(box.width);
  const height = emu(box.height);
  return `<a:graphic><a:graphicData uri="${wordShapeNamespace}"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${width}" cy="${height}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr><wps:txbx><w:txbxContent><w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="${Math.round(run.height * 20)}" w:lineRule="exact"/><w:jc w:val="${run.rtl ? "right" : "left"}"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="${run.font}" w:hAnsi="${run.font}" w:cs="${run.font}"/>${run.bold ? "<w:b/><w:bCs/>" : ""}${run.italic ? "<w:i/><w:iCs/>" : ""}<w:color w:val="${run.color}"/><w:w w:val="${run.stretch}"/><w:sz w:val="${halfPoints}"/><w:szCs w:val="${halfPoints}"/>${run.rtl ? "<w:rtl/>" : ""}</w:rPr><w:t xml:space="preserve">${escapeWordText(run.text)}</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr lIns="0" rIns="0" tIns="0" bIns="0" wrap="none" anchor="t" vert="${direction}"><a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic>`;
}

function positionedPageDrawing(page: PositionedPage, pageNumber: number, firstTextId: number) {
  const picture = anchoredDrawing(pageNumber, `PDF page ${pageNumber} artwork`, { x: 0, y: 0, width: page.width, height: page.height }, true,
    pictureGraphic(pageNumber, emu(page.width), emu(page.height)));
  const text = page.text.map((run, index) => {
    const offset = (run.width - run.height) / 2;
    const box = run.rotation === 0 ? run : { x: run.x + offset, y: run.y - offset, width: run.height, height: run.width };
    return anchoredDrawing(firstTextId + index, `Editable text ${index + 1}`, box, false, textBoxGraphic(run, box));
  }).join("");
  return `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/></w:pPr>${picture}${text}</w:p>`;
}

function pageSection(size: PageSize, last: boolean) {
  const { width, height, margin } = size;
  const properties = `<w:sectPr><w:type w:val="nextPage"/><w:pgSz w:w="${width}" w:h="${height}"${width > height ? ' w:orient="landscape"' : ""}/><w:pgMar w:top="${margin}" w:right="${margin}" w:bottom="${margin}" w:left="${margin}" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr>`;
  return last ? properties : `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/>${properties}</w:pPr></w:p>`;
}

export async function convertPdfDocument(
  pdf: PDFDocumentProxy,
  fileName: string,
  mode: WordConversionMode,
  onProgress?: (completed: number, total: number) => void,
  signal?: AbortSignal,
): Promise<ConvertedWord> {
  const zip = new JSZip();
  const body: string[] = [];
  const images: string[] = [];
  let nextTextId = pdf.numPages + 1;
  let imagePages = 0;
  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    signal?.throwIfAborted();
    const page = await pdf.getPage(pageNumber);
    try {
      const viewport = page.getViewport({ scale: 1 });
      const lines = mode === "reflow" ? await collectLines(page) : [];
      signal?.throwIfAborted();
      const size = {
        width: Math.min(31680, Math.max(1440, Math.round(viewport.width * 20))),
        height: Math.min(31680, Math.max(1440, Math.round(viewport.height * 20))),
        margin: 0,
      };
      if (mode === "editable") {
        const positioned = await positionPdfPage(page, size.width / 20, size.height / 20, signal);
        const path = `media/pdf-page-${pageNumber}.jpeg`;
        zip.file(`word/${path}`, positioned.image.slice(positioned.image.indexOf(",") + 1), { base64: true, compression: "STORE" });
        images.push(`<Relationship Id="page${pageNumber}" Type="${officeRelationships}/image" Target="${path}"/>`);
        body.push(positionedPageDrawing(positioned, pageNumber, nextTextId));
        nextTextId += positioned.text.length;
        if (!positioned.text.length) imagePages += 1;
      } else if (lines.length) {
        size.margin = Math.min(840, Math.floor(size.width / 10), Math.floor(size.height / 10));
        body.push(editablePage(lines, size));
      } else {
        const path = `media/pdf-page-${pageNumber}.jpeg`;
        zip.file(`word/${path}`, await renderPageImage(page, signal), { compression: "STORE" });
        images.push(`<Relationship Id="page${pageNumber}" Type="${officeRelationships}/image" Target="${path}"/>`);
        body.push(pageDrawing(pageNumber, size, viewport.width, viewport.height));
        imagePages += 1;
      }
      body.push(pageSection(size, pageNumber === pdf.numPages));
      onProgress?.(pageNumber, pdf.numPages);
    } finally {
      page.cleanup();
    }
  }

  signal?.throwIfAborted();
  zip.file("[Content_Types].xml", `${xmlHeader}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="jpeg" ContentType="image/jpeg"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`);
  zip.file("_rels/.rels", `${xmlHeader}<Relationships xmlns="${packageRelationships}"><Relationship Id="rId1" Type="${officeRelationships}/officeDocument" Target="word/document.xml"/></Relationships>`);
  if (images.length) zip.file("word/_rels/document.xml.rels", `${xmlHeader}<Relationships xmlns="${packageRelationships}">${images.join("")}</Relationships>`);
  zip.file("word/document.xml", `${xmlHeader}<w:document xmlns:w="${wordNamespace}" xmlns:r="${officeRelationships}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:wps="${wordShapeNamespace}"><w:body>${body.join("")}</w:body></w:document>`);
  const blob = await zip.generateAsync({ type: "blob", mimeType: wordMime, compression: "DEFLATE" });
  return { blob, fileName: fileName.replace(/\.pdf$/i, "") + ".docx", pageCount: pdf.numPages, imagePages };
}

export async function convertPdfToWord(
  file: File,
  mode: WordConversionMode,
  onProgress?: (completed: number, total: number) => void,
  signal?: AbortSignal,
) {
  const task = await openPdf(file, { fontExtraProperties: mode === "editable" });
  try {
    signal?.throwIfAborted();
    return await convertPdfDocument(await task.promise, file.name, mode, onProgress, signal);
  } finally {
    await task.destroy();
  }
}
