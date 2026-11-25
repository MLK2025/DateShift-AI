import assert from "node:assert/strict";
import { createRequire } from "node:module";
import { dirname, join, sep } from "node:path";
import test from "node:test";
import JSZip from "jszip";
import { PDFDocument, StandardFonts, degrees, rgb } from "pdf-lib";
import { updateDocumentXml } from "../app/lib/document-export";
import { readLedgerXml } from "../app/lib/ledger-docx";
import { convertPdfDocument } from "../app/lib/pdf-to-word";
import { readPositionedWordPreview } from "../app/lib/positioned-word-preview";
import { indexWordText } from "../app/lib/word-text";

const require = createRequire(import.meta.url);
const pdfRoot = dirname(require.resolve("pdfjs-dist/package.json"));
const mammothRoot = dirname(require.resolve("mammoth/package.json"));
const { DOMParser, XMLSerializer } = require(require.resolve("@xmldom/xmldom", { paths: [mammothRoot] }));

test("converts editable transaction rows and scanned pages into a valid Word package", async () => {
  const { createCanvas, loadImage } = require(require.resolve("@napi-rs/canvas", { paths: [pdfRoot] }));
  const { getDocument } = await import("pdfjs-dist/legacy/build/pdf.mjs");
  const source = await PDFDocument.create();
  const page = source.addPage([612, 792]);
  const font = await source.embedFont(StandardFonts.Helvetica);
  const bold = await source.embedFont(StandardFonts.HelveticaBold);
  page.drawRectangle({ x: 30, y: 725, width: 550, height: 36, color: rgb(.1, .2, .6) });
  page.drawText("Account activity & charges", { font: bold, x: 42, y: 740, size: 18, color: rgb(1, 1, 1) });
  page.drawText("Side note", { font, x: 22, y: 410, size: 10, rotate: degrees(90) });
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

  const load = getDocument({
    data: await source.save(), standardFontDataUrl: join(pdfRoot, "standard_fonts") + sep,
    useWorkerFetch: false, fontExtraProperties: true,
  });
  const originalDocument = Object.getOwnPropertyDescriptor(globalThis, "document");
  const originalParser = Object.getOwnPropertyDescriptor(globalThis, "DOMParser");
  const originalSerializer = Object.getOwnPropertyDescriptor(globalThis, "XMLSerializer");
  Object.defineProperty(globalThis, "document", { configurable: true, value: { createElement: () => createCanvas(1, 1) } });
  Object.defineProperty(globalThis, "DOMParser", { configurable: true, value: DOMParser });
  Object.defineProperty(globalThis, "XMLSerializer", { configurable: true, value: XMLSerializer });
  try {
    const pdf = await load.promise;
    const progress: [number, number][] = [];
    const edited = await convertPdfDocument(pdf, "monthly.PDF", "reflow", (page, total) => progress.push([page, total]));
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
      [["12-Oct-2026", BigInt(0), BigInt(3000), BigInt(13000)], ["13-Oct-2026", BigInt(1500), BigInt(0), BigInt(11500)]]);

    const positioned = await convertPdfDocument(pdf, "monthly.pdf", "editable");
    const positionZip = await JSZip.loadAsync(await positioned.blob.arrayBuffer());
    const positionDoc = await positionZip.file("word/document.xml")?.async("string");
    assert.ok(positionDoc);
    assert.equal(positioned.imagePages, 1);
    assert.match(positionDoc, /<wps:txbx><w:txbxContent>/);
    assert.match(positionDoc, /Account activity &amp; charges/);
    assert.match(positionDoc, /w:orient="landscape"/);
    assert.match(positionDoc, /behindDoc="1"/);
    assert.match(positionDoc, /vert="vert270"/);
    assert.ok(positionZip.file("word/media/pdf-page-1.jpeg"));
    assert.ok(positionZip.file("word/media/pdf-page-2.jpeg"));
    const preview = await readPositionedWordPreview(positioned.blob);
    assert.ok(preview);
    const title = preview[0].text.find(({ text }) => text === "Account activity & charges");
    assert.ok(title && Math.abs(title.x - 42) < 1 && Math.abs(title.y - 35) < 2);
    assert.equal(title.color, "FFFFFF");
    assert.equal(title.bold, true);
    assert.equal(preview[1].text.length, 0);
    const artwork = await loadImage(preview[0].image);
    const artworkCanvas = createCanvas(artwork.width, artwork.height);
    const artworkContext = artworkCanvas.getContext("2d");
    artworkContext.drawImage(artwork, 0, 0);
    const pixels = artworkContext.getImageData(82, 70, 460, 40).data;
    assert.ok(pixels.every((byte: number, index: number) => index % 4 >= 2 || byte < 100));
    const originalDate = "12-Oct-2026";
    const sourceIndex = indexWordText(positionDoc).text.indexOf(originalDate);
    assert.ok(sourceIndex >= 0);
    const changed = updateDocumentXml(positionDoc, [{ original: originalDate, updated: "12-Nov-2026", sourceIndex, selected: true }]);
    assert.equal(changed.appliedCount, 1);
    assert.match(changed.xml, /<w:t xml:space="preserve">12-Nov-2026<\/w:t>/);
    positionZip.file("word/document.xml", changed.xml);
    const changedPreview = await readPositionedWordPreview(await positionZip.generateAsync({ type: "blob" }));
    assert.ok(changedPreview?.[0].text.some(({ text }) => text === "12-Nov-2026"));

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
    if (originalSerializer) Object.defineProperty(globalThis, "XMLSerializer", originalSerializer);
    else Reflect.deleteProperty(globalThis, "XMLSerializer");
    await load.destroy();
  }
});
