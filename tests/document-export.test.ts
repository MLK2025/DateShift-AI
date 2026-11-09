import assert from "node:assert/strict";
import { mkdir, writeFile } from "node:fs/promises";
import { createRequire } from "node:module";
import { dirname } from "node:path";
import test from "node:test";
import JSZip from "jszip";
import { PDFDocument, StandardFonts, rgb } from "pdf-lib";
import { createUpdatedDocument, type DocumentChange } from "../app/lib/document-export";

const require = createRequire(import.meta.url);
const mammothRoot = dirname(require.resolve("mammoth/package.json"));
const xmldom = require(require.resolve("@xmldom/xmldom", { paths: [mammothRoot] }));
Object.assign(globalThis, { DOMParser: xmldom.DOMParser, XMLSerializer: xmldom.XMLSerializer });

const outputDir = "/tmp/dateshift-export-qa";

function changes(): DocumentChange[] {
  return [{ original: "20-12-2026", updated: "20-03-2027", sourceIndex: 0, selected: true }];
}

async function sampleDocx() {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`);
  zip.folder("_rels")?.file(".rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`);
  zip.folder("word")?.file("document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl><w:tblPr><w:tblW w:w="7200" w:type="dxa"/><w:tblBorders><w:top w:val="single" w:sz="8" w:color="24405F"/><w:left w:val="single" w:sz="8" w:color="24405F"/><w:bottom w:val="single" w:sz="8" w:color="24405F"/><w:right w:val="single" w:sz="8" w:color="24405F"/></w:tblBorders></w:tblPr><w:tr><w:tc><w:tcPr><w:shd w:fill="DCE6F1"/></w:tcPr><w:sdt><w:sdtPr><w:date w:fullDate="2026-12-20T00:00:00Z"><w:dateFormat w:val="dd-MM-yyyy"/></w:date></w:sdtPr><w:sdtContent><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:color w:val="C00000"/></w:rPr><w:t>20-</w:t></w:r><w:r><w:rPr><w:i/><w:color w:val="0070C0"/></w:rPr><w:t>12-</w:t></w:r><w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>2026</w:t></w:r></w:p></w:sdtContent></w:sdt></w:tc></w:tr></w:tbl><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`);
  zip.folder("word")?.file("styles.xml", "STYLES-MUST-STAY-UNCHANGED");
  zip.folder("word")?.file("header1.xml", "HEADER-MUST-STAY-UNCHANGED");
  return zip.generateAsync({ type: "uint8array" });
}

async function weekdaySampleDocx() {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`);
  zip.folder("_rels")?.file(".rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`);
  zip.folder("word")?.file("document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:sdt><w:sdtPr><w:date w:fullDate="2026-03-04T15:57:15Z"><w:dateFormat w:val="ddd MMM dd HH:mm:ss 'GMT' yyyy"/></w:date></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Date : Wed Mar </w:t></w:r><w:r><w:t>04 15:57:15 GMT 2026</w:t></w:r></w:p></w:sdtContent></w:sdt><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`);
  return zip.generateAsync({ type: "uint8array" });
}

test("writes a selected date into a DOCX table and syncs its date control", async () => {
  await mkdir(outputDir, { recursive: true });
  const original = await sampleDocx();
  const file = new File([original.slice().buffer as ArrayBuffer], "table-dates.docx", { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
  const result = await createUpdatedDocument(file, { kind: "docx", text: "20-12-2026", pdfSpans: [] }, changes());
  assert.equal(result.updatedCount, 1);

  const updatedBytes = new Uint8Array(await result.blob.arrayBuffer());
  const updatedZip = await JSZip.loadAsync(updatedBytes);
  const xml = await updatedZip.file("word/document.xml")?.async("string");
  const parsed = new DOMParser().parseFromString(xml ?? "", "application/xml");
  const visibleText = Array.from(parsed.getElementsByTagNameNS("*", "t")).map((node) => node.textContent ?? "").join("");
  assert.ok(visibleText.includes("20-03-2027"));
  assert.ok(!visibleText.includes("20-12-2026"));
  assert.ok(xml?.includes("2027-03-20T00:00:00Z"));
  assert.match(xml ?? "", /<w:r><w:rPr><w:b\/?><w:color w:val="C00000"\/?><\/w:rPr><w:t>20-<\/w:t><\/w:r>/);
  assert.match(xml ?? "", /<w:r><w:rPr><w:i\/?><w:color w:val="0070C0"\/?><\/w:rPr><w:t>03-<\/w:t><\/w:r>/);
  assert.match(xml ?? "", /<w:r><w:rPr><w:u w:val="single"\/?><\/w:rPr><w:t>2027<\/w:t><\/w:r>/);
  assert.equal(await updatedZip.file("word/styles.xml")?.async("string"), "STYLES-MUST-STAY-UNCHANGED");
  assert.equal(await updatedZip.file("word/header1.xml")?.async("string"), "HEADER-MUST-STAY-UNCHANGED");
  await writeFile(`${outputDir}/table-dates-updated.docx`, updatedBytes);
});

test("writes a timestamp date and syncs its Word date control", async () => {
  await mkdir(outputDir, { recursive: true });
  const original = await weekdaySampleDocx();
  const file = new File([original.slice().buffer as ArrayBuffer], "weekday-date.docx", { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
  const change = { original: "Wed Mar 04 15:57:15 GMT 2026", updated: "Thu Jun 04 15:57:15 GMT 2026", sourceIndex: 7, selected: true };
  const result = await createUpdatedDocument(file, { kind: "docx", text: `Date : ${change.original}`, pdfSpans: [] }, [change]);
  assert.equal(result.updatedCount, 1);

  const updatedBytes = new Uint8Array(await result.blob.arrayBuffer());
  const updatedZip = await JSZip.loadAsync(updatedBytes);
  const xml = await updatedZip.file("word/document.xml")?.async("string");
  const parsed = new DOMParser().parseFromString(xml ?? "", "application/xml");
  const visibleText = Array.from(parsed.getElementsByTagNameNS("*", "t")).map((node) => node.textContent ?? "").join("");
  assert.ok(visibleText.includes("Date : Thu Jun 04 15:57:15 GMT 2026"));
  assert.ok(!visibleText.includes("Wed Mar 04 15:57:15 GMT 2026"));
  assert.ok(xml?.includes("2026-06-04T15:57:15Z"));
  await writeFile(`${outputDir}/weekday-date-updated.docx`, updatedBytes);
});

test("writes a selected date over the correct substring in a PDF text span", async () => {
  await mkdir(outputDir, { recursive: true });
  const source = await PDFDocument.create();
  const page = source.addPage([612, 792]);
  const font = await source.embedFont(StandardFonts.Helvetica);
  const text = "Date : Wed Mar 04 15:57:15 GMT 2026";
  const size = 12;
  const x = 72;
  const y = 700;
  const backgroundColor = { red: 0.1, green: 0.28, blue: 0.58 };
  const textColor = { red: 1, green: 1, blue: 1 };
  page.drawRectangle({ x: 60, y: 680, width: 400, height: 44, color: rgb(backgroundColor.red, backgroundColor.green, backgroundColor.blue) });
  page.drawText(text, { x, y, size, font, color: rgb(textColor.red, textColor.green, textColor.blue) });
  const original = await source.save();
  const file = new File([original.slice().buffer as ArrayBuffer], "date-table.pdf", { type: "application/pdf" });
  const result = await createUpdatedDocument(file, {
    kind: "pdf",
    text,
    pdfSpans: [{
      pageNumber: 1,
      start: 0,
      end: text.length,
      text,
      x,
      y,
      width: font.widthOfTextAtSize(text, size),
      height: size,
      rotation: 0,
      backgroundColor,
      textColor,
    }],
  }, [{ original: "Wed Mar 04 15:57:15 GMT 2026", updated: "Thu Jun 04 15:57:15 GMT 2026", sourceIndex: 7, selected: true }]);
  assert.equal(result.updatedCount, 1);
  const updatedBytes = new Uint8Array(await result.blob.arrayBuffer());
  assert.ok(updatedBytes.length > original.length);
  await writeFile(`${outputDir}/date-table-updated.pdf`, updatedBytes);
});
