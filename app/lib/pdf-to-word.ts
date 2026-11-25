import JSZip from "jszip";
import type { PDFDocumentProxy, PDFPageProxy } from "pdfjs-dist";
import { openPdf } from "./document-text";
import { hasTransactionColumns } from "./ledger-docx";

export type WordConversionMode = "editable" | "layout";
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
  const size = page.getViewport({ scale: 1 });
  const scale = Math.min(1.65, 1600 / Math.max(size.width, size.height));
  const viewport = page.getViewport({ scale });
  const canvas = document.createElement("canvas");
  canvas.width = Math.ceil(viewport.width);
  canvas.height = Math.ceil(viewport.height);
  try {
    const render = page.render({ canvas, viewport, background: "#ffffff" });
    const cancel = () => render.cancel();
    signal?.addEventListener("abort", cancel, { once: true });
    try {
      await render.promise;
    } finally {
      signal?.removeEventListener("abort", cancel);
    }
    signal?.throwIfAborted();
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

function pageDrawing(pageNumber: number, size: PageSize, originalWidth: number, originalHeight: number) {
  const unitsPerTwip = 635;
  // Reserve a few points for Word's paragraph baseline and section break.
  const scale = Math.min(size.width / originalWidth, (size.height - 100) / originalHeight);
  const width = Math.round(originalWidth * scale * unitsPerTwip);
  const height = Math.round(originalHeight * scale * unitsPerTwip);
  const description = `PDF page ${pageNumber}`;
  return `<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="0"/><w:rPr><w:sz w:val="2"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="2"/></w:rPr><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${width}" cy="${height}"/><wp:docPr id="${pageNumber}" name="${description}" descr="${description}"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${pageNumber}" name="${description}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="page${pageNumber}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${width}" cy="${height}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
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
  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    signal?.throwIfAborted();
    const page = await pdf.getPage(pageNumber);
    try {
      const viewport = page.getViewport({ scale: 1 });
      const lines = mode === "editable" ? await collectLines(page) : [];
      signal?.throwIfAborted();
      const size = {
        width: Math.min(31680, Math.max(1440, Math.round(viewport.width * 20))),
        height: Math.min(31680, Math.max(1440, Math.round(viewport.height * 20))),
        margin: 0,
      };
      if (lines.length) {
        size.margin = Math.min(840, Math.floor(size.width / 10), Math.floor(size.height / 10));
        body.push(editablePage(lines, size));
      } else {
        const path = `media/pdf-page-${pageNumber}.jpeg`;
        zip.file(`word/${path}`, await renderPageImage(page, signal), { compression: "STORE" });
        images.push(`<Relationship Id="page${pageNumber}" Type="${officeRelationships}/image" Target="${path}"/>`);
        body.push(pageDrawing(pageNumber, size, viewport.width, viewport.height));
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
  zip.file("word/document.xml", `${xmlHeader}<w:document xmlns:w="${wordNamespace}" xmlns:r="${officeRelationships}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body>${body.join("")}</w:body></w:document>`);
  const blob = await zip.generateAsync({ type: "blob", mimeType: wordMime, compression: "DEFLATE" });
  return { blob, fileName: fileName.replace(/\.pdf$/i, "") + ".docx", pageCount: pdf.numPages, imagePages: images.length };
}

export async function convertPdfToWord(
  file: File,
  mode: WordConversionMode,
  onProgress?: (completed: number, total: number) => void,
  signal?: AbortSignal,
) {
  const task = await openPdf(file);
  try {
    signal?.throwIfAborted();
    return await convertPdfDocument(await task.promise, file.name, mode, onProgress, signal);
  } finally {
    await task.destroy();
  }
}
