import assert from "node:assert/strict";
import { createRequire } from "node:module";
import { dirname, join, sep } from "node:path";
import test from "node:test";
import JSZip from "jszip";
import { PDFDocument, StandardFonts, rgb } from "pdf-lib";
import { readLedgerXml } from "../app/lib/ledger-docx";
import { convertPdfDocument } from "../app/lib/pdf-to-word";

const require = createRequire(import.meta.url);
const pdfRoot = dirname(require.resolve("pdfjs-dist/package.json"));
const mammothRoot = dirname(require.resolve("mammoth/package.json"));
const { DOMParser } = require(require.resolve("@xmldom/xmldom", { paths: [mammothRoot] }));

test("converts editable transaction rows and scanned pages into a valid Word package", async () => {
  const { createCanvas } = require(require.resolve("@napi-rs/canvas", { paths: [pdfRoot] }));
  const { getDocument } = await import("pdfjs-dist/legacy/build/pdf.mjs");
  const source = await PDFDocument.create();
  const page = source.addPage([612, 792]);
  const font = await source.embedFont(StandardFonts.Helvetica);
  page.drawRectangle({ x: 30, y: 725, width: 550, height: 36, color: rgb(.93, .95, 1) });
  page.drawText("Account activity & charges", { font, x: 42, y: 740, size: 18 });
  const rows = [
    ["Date", "Description", "Debit", "Credit", "Balance"],
    ["12-Oct-2026", "Deposit", "", "30.00", "130.00"],
    ["13-Oct-2026", "Supplies", "15.00", "", "115.00"],
  ];
  rows.forEach((row, index) => row.forEach((text, col) => {
    if (text) page.drawText(text, { font, x: [42, 135, 325, 405, 490][col], y: 690 - index * 19, size: 11 });
  }));

  const scan = source.addPage([792, 612]);
  const image = createCanvas(350, 200);
  const context = image.getContext("2d");
  context.fillStyle = "#3155d9";
  context.fillRect(0, 0, image.width, image.height);
  context.fillStyle = "#ffffff";
  context.font = "23px Arial";
  context.fillText("Image-only PDF page", 22, 60);
  scan.drawImage(await source.embedPng(image.toBuffer("image/png")), { x: 45, y: 350, width: 350, height: 200 });

  const load = getDocument({ data: await source.save(), standardFontDataUrl: join(pdfRoot, "standard_fonts") + sep, useWorkerFetch: false });
  const originalDocument = Object.getOwnPropertyDescriptor(globalThis, "document");
  const originalParser = Object.getOwnPropertyDescriptor(globalThis, "DOMParser");
  Object.defineProperty(globalThis, "document", { configurable: true, value: { createElement: () => createCanvas(1, 1) } });
  Object.defineProperty(globalThis, "DOMParser", { configurable: true, value: DOMParser });
  try {
    const pdf = await load.promise;
    const progress: [number, number][] = [];
    const edited = await convertPdfDocument(pdf, "monthly.PDF", "editable", (page, total) => progress.push([page, total]));
    assert.deepEqual(progress, [[1, 2], [2, 2]]);
    assert.equal(edited.fileName, "monthly.docx");
    assert.equal(edited.pageCount, 2);
    assert.equal(edited.imagePages, 1);
    const zip = await JSZip.loadAsync(await edited.blob.arrayBuffer());
    const doc = await zip.file("word/document.xml")?.async("string");
    assert.ok(doc);
    assert.match(doc, /Account activity &amp; charges/);
    assert.match(doc, /<w:tbl>/);
    assert.match(doc, /w:orient="landscape"/);
    assert.match(doc, /r:embed="page2"/);
    assert.ok(zip.file("word/media/pdf-page-2.jpeg"));
    assert.equal(zip.file("word/media/pdf-page-1.jpeg"), null);
    const ledger = readLedgerXml(doc);
    assert.deepEqual(ledger.transactions.map(({ date, debit, credit, balance }) => [date, debit, credit, balance]),
      [["12-Oct-2026", 0n, 3000n, 13000n], ["13-Oct-2026", 1500n, 0n, 11500n]]);

    const layout = await convertPdfDocument(pdf, "monthly.pdf", "layout");
    const imageZip = await JSZip.loadAsync(await layout.blob.arrayBuffer());
    const imageDoc = await imageZip.file("word/document.xml")?.async("string");
    assert.equal(layout.imagePages, 2);
    assert.ok(imageDoc);
    assert.doesNotMatch(imageDoc, /Account activity|<w:tbl>/);
    assert.match(imageDoc, /r:embed="page1"[\s\S]*r:embed="page2"/);
    for (const number of [1, 2]) {
      const jpg = await imageZip.file(`word/media/pdf-page-${number}.jpeg`)?.async("uint8array");
      assert.ok(jpg && jpg.length > 1000);
      assert.deepEqual(jpg.slice(0, 2), Uint8Array.of(0xff, 0xd8));
    }
  } finally {
    if (originalDocument) Object.defineProperty(globalThis, "document", originalDocument);
    else Reflect.deleteProperty(globalThis, "document");
    if (originalParser) Object.defineProperty(globalThis, "DOMParser", originalParser);
    else Reflect.deleteProperty(globalThis, "DOMParser");
    await load.destroy();
  }
});
