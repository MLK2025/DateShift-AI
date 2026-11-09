import JSZip from "jszip";
import { degrees, PDFDocument, rgb, StandardFonts } from "pdf-lib";
import type { DocumentExtraction } from "./document-text";

export type DocumentChange = {
  original: string;
  updated: string;
  sourceIndex: number;
  selected: boolean;
};

export type UpdatedDocument = {
  blob: Blob;
  fileName: string;
  updatedCount: number;
  warning?: string;
};

function updatedName(fileName: string, extension: "docx" | "pdf") {
  return `${fileName.replace(/\.[^.]+$/, "")}-updated.${extension}`;
}

function setTextPreservingSpaces(element: Element, value: string) {
  element.textContent = value;
  if (/^\s|\s$/.test(value)) element.setAttribute("xml:space", "preserve");
}

function applyReplacementAcrossNodes(nodes: Element[], starts: number[], start: number, end: number, replacement: string) {
  const affected: Array<{
    node: Element;
    value: string;
    localStart: number;
    localEnd: number;
  }> = [];

  for (let index = 0; index < nodes.length; index += 1) {
    const value = nodes[index].textContent ?? "";
    const nodeStart = starts[index];
    const nodeEnd = nodeStart + value.length;
    if (nodeEnd <= start || nodeStart >= end) continue;

    const localStart = Math.max(0, start - nodeStart);
    const localEnd = Math.min(value.length, end - nodeStart);
    affected.push({ node: nodes[index], value, localStart, localEnd });
  }

  if (!affected.length) return null;

  let replacementOffset = 0;
  for (let index = 0; index < affected.length; index += 1) {
    const { node, value, localStart, localEnd } = affected[index];
    const isFirst = index === 0;
    const isLast = index === affected.length - 1;
    const prefix = isFirst ? value.slice(0, localStart) : "";
    const suffix = isLast ? value.slice(localEnd) : "";
    const remainingLength = replacement.length - replacementOffset;
    const originalShare = localEnd - localStart;
    const replacementShare = isLast ? remainingLength : Math.min(originalShare, remainingLength);
    const segment = replacement.slice(replacementOffset, replacementOffset + replacementShare);
    replacementOffset += replacementShare;
    setTextPreservingSpaces(node, `${prefix}${segment}${suffix}`);
  }

  return affected[0].node;
}

function ancestorByLocalName(node: Element, localName: string) {
  let current: Element | null = node;
  while (current) {
    if (current.localName === localName || current.nodeName.endsWith(`:${localName}`)) return current;
    const parent = current.parentElement ?? current.parentNode;
    current = parent?.nodeType === 1 ? parent as Element : null;
  }
  return null;
}

function isoDateForControl(value: string) {
  const iso = value.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (iso) return `${iso[1]}-${iso[2]}-${iso[3]}T00:00:00Z`;

  const monthNames = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
  const timestamp = value.match(/^(?:Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday|Mon|Tue|Tues|Wed|Thu|Thur|Thurs|Fri|Sat|Sun),?\s+(January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)\.?\s+(\d{1,2})(?:st|nd|rd|th)?[,]?\s+(\d{2}:\d{2}(?::\d{2})?)\s+((?:GMT|UTC)(?:[+-]\d{4})?)\s+(\d{4})$/i);
  if (timestamp) {
    const month = monthNames.indexOf(timestamp[1].slice(0, 3).toLowerCase()) + 1;
    const time = timestamp[3].length === 5 ? `${timestamp[3]}:00` : timestamp[3];
    const offset = timestamp[4].match(/[+-](\d{2})(\d{2})$/);
    const zone = offset ? `${timestamp[4].includes("-") ? "-" : "+"}${offset[1]}:${offset[2]}` : "Z";
    return `${timestamp[5]}-${String(month).padStart(2, "0")}-${String(Number(timestamp[2])).padStart(2, "0")}T${time}${zone}`;
  }

  const named = value.match(/^(?:(?:Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday|Mon|Tue|Tues|Wed|Thu|Thur|Thurs|Fri|Sat|Sun),?\s+)?(January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)\.?\s+(\d{1,2})(?:st|nd|rd|th)?[,]?\s+(\d{4})$/i);
  if (named) {
    const month = monthNames.indexOf(named[1].slice(0, 3).toLowerCase()) + 1;
    return `${named[3]}-${String(month).padStart(2, "0")}-${String(Number(named[2])).padStart(2, "0")}T00:00:00Z`;
  }

  const numeric = value.match(/^(\d{1,2})([./-])(\d{1,2})\2(\d{4})$/);
  if (!numeric) return null;
  const dayFirst = numeric[2] === "-" || Number(numeric[1]) > 12;
  const day = String(dayFirst ? numeric[1] : numeric[3]).padStart(2, "0");
  const month = String(dayFirst ? numeric[3] : numeric[1]).padStart(2, "0");
  return `${numeric[4]}-${month}-${day}T00:00:00Z`;
}

function syncDateContentControl(insertionNode: Element, updatedValue: string) {
  const contentControl = ancestorByLocalName(insertionNode, "sdt");
  const isoDate = isoDateForControl(updatedValue);
  if (!contentControl || !isoDate) return;

  const dateProperties = Array.from(contentControl.getElementsByTagNameNS("*", "date"));
  if (!dateProperties.length) {
    dateProperties.push(...Array.from(contentControl.getElementsByTagName("*")).filter((element) => element.localName === "date" || element.nodeName.endsWith(":date")));
  }
  for (const dateProperty of dateProperties) {
    const currentAttribute = Array.from(dateProperty.attributes).find((attribute) => attribute.localName === "fullDate");
    if (currentAttribute) dateProperty.setAttributeNS(currentAttribute.namespaceURI, currentAttribute.name, isoDate);
  }
}

export function updateDocumentXml(xml: string, changes: DocumentChange[]) {
  const parser = new DOMParser();
  const document = parser.parseFromString(xml, "application/xml");
  if (document.getElementsByTagName("parsererror").length) throw new Error("This Word document contains XML that could not be edited safely.");

  const queues = new Map<string, DocumentChange[]>();
  for (const change of [...changes].sort((a, b) => a.sourceIndex - b.sourceIndex)) {
    const queue = queues.get(change.original) ?? [];
    queue.push(change);
    queues.set(change.original, queue);
  }
  const consumed = new Map<string, number>();
  const paragraphs = Array.from(document.getElementsByTagNameNS("*", "p"));
  let appliedCount = 0;

  for (const paragraph of paragraphs) {
    const nodes = Array.from(paragraph.getElementsByTagNameNS("*", "t"));
    if (!nodes.length) continue;
    const starts: number[] = [];
    let paragraphText = "";
    for (const node of nodes) {
      starts.push(paragraphText.length);
      paragraphText += node.textContent ?? "";
    }

    const replacements: Array<{ start: number; end: number; value: string }> = [];
    for (const [original, queue] of queues) {
      let searchFrom = 0;
      let occurrence = paragraphText.indexOf(original, searchFrom);
      while (occurrence >= 0) {
        const queueIndex = consumed.get(original) ?? 0;
        const change = queue[queueIndex];
        if (change) {
          consumed.set(original, queueIndex + 1);
          if (change.selected) replacements.push({ start: occurrence, end: occurrence + original.length, value: change.updated });
        }
        searchFrom = occurrence + original.length;
        occurrence = paragraphText.indexOf(original, searchFrom);
      }
    }

    for (const replacement of replacements.sort((a, b) => b.start - a.start)) {
      const insertionNode = applyReplacementAcrossNodes(nodes, starts, replacement.start, replacement.end, replacement.value);
      if (insertionNode) {
        syncDateContentControl(insertionNode, replacement.value);
        appliedCount += 1;
      }
    }
  }

  return { xml: new XMLSerializer().serializeToString(document), appliedCount };
}

async function updateDocx(file: File, changes: DocumentChange[]): Promise<UpdatedDocument> {
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const documentEntry = zip.file("word/document.xml");
  if (!documentEntry) throw new Error("This file does not contain a readable Word document body.");
  const selectedCount = changes.filter((change) => change.selected).length;
  const update = updateDocumentXml(await documentEntry.async("string"), changes);
  if (!update.appliedCount) throw new Error("The dates were found, but this Word file stores them in a format that could not be safely updated. No unchanged copy was downloaded.");
  zip.file("word/document.xml", update.xml);
  const blob = await zip.generateAsync({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });
  return {
    blob,
    fileName: updatedName(file.name, "docx"),
    updatedCount: update.appliedCount,
    warning: update.appliedCount < selectedCount
      ? `${selectedCount - update.appliedCount} selected date${selectedCount - update.appliedCount === 1 ? " was" : "s were"} not stored as editable Word text and remained unchanged.`
      : undefined,
  };
}

function escapeXml(value: string) {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

async function createSampleDocx(text: string, changes: DocumentChange[]): Promise<UpdatedDocument> {
  let updatedText = text;
  for (const change of changes.filter((item) => item.selected).sort((a, b) => b.sourceIndex - a.sourceIndex)) {
    updatedText = `${updatedText.slice(0, change.sourceIndex)}${change.updated}${updatedText.slice(change.sourceIndex + change.original.length)}`;
  }
  const paragraphs = updatedText.split(/\n+/).filter(Boolean).map((paragraph) => `<w:p><w:r><w:t xml:space="preserve">${escapeXml(paragraph)}</w:t></w:r></w:p>`).join("");
  const zip = new JSZip();
  zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`);
  zip.folder("_rels")?.file(".rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`);
  zip.folder("word")?.file("document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${paragraphs}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`);
  const blob = await zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
  return { blob, fileName: "Service_Agreement_2026-updated.docx", updatedCount: changes.filter((change) => change.selected).length };
}

async function updatePdf(file: File, extraction: DocumentExtraction, changes: DocumentChange[]): Promise<UpdatedDocument> {
  const pdf = await PDFDocument.load(await file.arrayBuffer());
  const font = await pdf.embedFont(StandardFonts.Helvetica);
  const pages = pdf.getPages();
  let updatedCount = 0;
  let unmatched = 0;

  for (const change of changes.filter((item) => item.selected)) {
    const start = change.sourceIndex;
    const end = start + change.original.length;
    const matchingSpans = extraction.pdfSpans.filter((span) => span.end > start && span.start < end);
    if (!matchingSpans.length) {
      unmatched += 1;
      continue;
    }
    const pageNumber = matchingSpans[0].pageNumber;
    const samePage = matchingSpans.filter((span) => span.pageNumber === pageNumber);
    const page = pages[pageNumber - 1];
    if (!page) {
      unmatched += 1;
      continue;
    }

    const fragments = samePage.map((span) => {
      const characterCount = Math.max(span.text.length, 1);
      const localStart = Math.max(0, start - span.start);
      const localEnd = Math.min(characterCount, end - span.start);
      const measuredText = font.widthOfTextAtSize(span.text, 100);
      const measuredPrefix = font.widthOfTextAtSize(span.text.slice(0, localStart), 100);
      const measuredMatch = font.widthOfTextAtSize(span.text.slice(localStart, localEnd), 100);
      const startRatio = measuredText > 0 ? measuredPrefix / measuredText : localStart / characterCount;
      const widthRatio = measuredText > 0 ? measuredMatch / measuredText : Math.max(localEnd - localStart, 1) / characterCount;
      const radians = span.rotation * Math.PI / 180;
      const advance = span.width * startRatio;
      return {
        x: span.x + Math.cos(radians) * advance,
        y: span.y + Math.sin(radians) * advance,
        width: Math.max(span.width * widthRatio, 4),
        height: span.height,
        rotation: span.rotation,
        backgroundColor: span.backgroundColor,
        textColor: span.textColor,
      };
    });
    const first = fragments[0];
    const availableWidth = Math.max(fragments.reduce((total, fragment) => total + fragment.width, 0), 8);
    const height = Math.max(...fragments.map((fragment) => fragment.height));
    let fontSize = Math.max(5, Math.min(height * 0.92, 18));
    const renderedWidth = font.widthOfTextAtSize(change.updated, fontSize);
    if (renderedWidth > availableWidth) fontSize = Math.max(4.5, fontSize * availableWidth / renderedWidth);

    for (const fragment of fragments) {
      const background = fragment.backgroundColor ?? { red: 1, green: 1, blue: 1 };
      page.drawRectangle({
        x: fragment.x - 1.5,
        y: fragment.y - 2,
        width: fragment.width + 3,
        height: fragment.height + 4,
        color: rgb(background.red, background.green, background.blue),
        rotate: degrees(fragment.rotation),
      });
    }
    const textColor = first.textColor ?? { red: 0, green: 0, blue: 0 };
    page.drawText(change.updated, { x: first.x, y: first.y, size: fontSize, font, color: rgb(textColor.red, textColor.green, textColor.blue), rotate: degrees(first.rotation), maxWidth: availableWidth });
    updatedCount += 1;
  }

  if (!updatedCount) throw new Error("The dates were found, but their positions could not be mapped onto this PDF. No unchanged copy was downloaded.");
  const bytes = await pdf.save();
  return {
    blob: new Blob([bytes.slice().buffer as ArrayBuffer], { type: "application/pdf" }),
    fileName: updatedName(file.name, "pdf"),
    updatedCount,
    warning: unmatched ? `${unmatched} selected date${unmatched === 1 ? " could" : "s could"} not be positioned on the PDF page and were left unchanged.` : undefined,
  };
}

export async function createUpdatedDocument(file: File | null, extraction: DocumentExtraction, changes: DocumentChange[]) {
  if (!file) return createSampleDocx(extraction.text, changes);
  if (extraction.kind === "docx") return updateDocx(file, changes);
  return updatePdf(file, extraction, changes);
}

export function downloadUpdatedDocument(result: UpdatedDocument) {
  const anchor = document.createElement("a");
  anchor.href = URL.createObjectURL(result.blob);
  anchor.download = result.fileName;
  anchor.style.display = "none";
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  window.setTimeout(() => URL.revokeObjectURL(anchor.href), 60_000);
}
