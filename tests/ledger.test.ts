import assert from "node:assert/strict";
import { createRequire } from "node:module";
import { dirname } from "node:path";
import test from "node:test";
import JSZip from "jszip";
import { createLedgerReport, createUpdatedLedgerDocument, readLedgerXml } from "../app/lib/ledger-docx";
import { formatMoney, parseMoney, projectLedger, reconcileLedger } from "../app/lib/ledger";
import { extractDocument } from "../app/lib/document-text";
import { findDates, interpretInstruction } from "../app/page";

const require = createRequire(import.meta.url);
const mammothDirectory = dirname(require.resolve("mammoth/package.json"));
const { DOMParser: XmlParser, XMLSerializer: XmlSerializer } = require(require.resolve("@xmldom/xmldom", { paths: [mammothDirectory] }));
Object.assign(globalThis, { DOMParser: XmlParser, XMLSerializer: XmlSerializer });

const wordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

function table(rows: string[][]) {
  return `<w:tbl>${rows.map((row) => `<w:tr>${row.map((cell) => `<w:tc><w:p><w:r><w:t>${cell}</w:t></w:r></w:p></w:tc>`).join("")}</w:tr>`).join("")}</w:tbl>`;
}

const summary = table([["Currency", "GHS"], ["Opening Balance", "1,000.00"], ["Closing Balance", "2,010.00"]]);
const statement = `<w:document xmlns:w="${wordNamespace}"><w:body>${summary}${summary}${table([
  ["Date", "Description", "Ref", "Value Date", "Debit", "Credit", "Balance"],
  ["01-Jan-2026", "Initial credit", "123456789", "02-Jan-2026", "", "400.00", "1,400.00"],
])}${table([
  ["03-Jan-2026", "Withdrawal", "987654321", "03-Jan-2026", "200.00", "", "1,200.00"],
  ["04-Jan-2026", "Final debit", "987654320", "04-Jan-2026", "10.00", "", "1,180.00"],
  ["", "", "", "", "", "", ""],
])}</w:body></w:document>`;

const cents = (value: string) => {
  const result = parseMoney(value);
  assert.notEqual(result, null);
  return result!;
};

test("reads continuation tables and duplicate summaries while flagging source discrepancies", () => {
  const ledger = readLedgerXml(statement);
  assert.equal(ledger.transactions.length, 3);
  assert.equal(ledger.currency, "GHS");
  assert.equal(ledger.reportedClosing, cents("2010.00"));
  assert.deepEqual(reconcileLedger(ledger), { debit: cents("210"), credit: cents("400"), closing: cents("1190"), mismatches: 1 });
});

test("each method anchors both the opening and closing rows, including the sample report", async () => {
  const ledger = readLedgerXml(statement);
  const target = cents("1250.01");
  const opening = cents("1000.00");
  for (const [specified, method] of [[null, "spread"], [opening, "spread"], [opening, "last"]] as const) {
    const projection = projectLedger(ledger, target, specified, method);
    assert.deepEqual(
      [projection.transactions[0].debit, projection.transactions[0].credit, projection.transactions[0].balance],
      [cents("0"), cents("0"), projection.opening],
    );
    let running = projection.opening;
    for (const transaction of projection.transactions) {
      running += transaction.credit - transaction.debit;
      assert.equal(transaction.balance, running);
    }
    assert.equal(projection.opening + projection.creditTotal - projection.debitTotal, target);
    assert.equal(projection.transactions.at(-1)?.balance, target);
  }
  const projection = projectLedger(ledger, target, opening, "last");
  assert.equal(projection.transactions[2].debit, cents("0"));
  assert.equal(projection.transactions[2].credit, cents("450.01"));
  const report = await JSZip.loadAsync(await (await createLedgerReport(ledger, projection)).arrayBuffer());
  const xml = await report.file("word/document.xml")?.async("string");
  assert.ok(xml);
  assert.match(xml, /Transaction projection/);
  assert.ok(xml.includes("Closing balance (GHS)  1,250.01"));
  const parsed = new DOMParser().parseFromString(xml, "application/xml");
  const rows = Array.from(parsed.getElementsByTagNameNS(wordNamespace, "tr"));
  assert.equal(rows[1].getElementsByTagNameNS(wordNamespace, "t").item(5)?.textContent, "1,000.00");
  assert.equal(rows.at(-1)?.getElementsByTagNameNS(wordNamespace, "t").item(5)?.textContent, "1,250.01");
});

test("a dedicated opening row leaves the first regular transaction in the calculation", () => {
  const openingRow = table([["01-Jan-2026", "Opening Balance", "", "", "", "", "1,000.00"]]);
  const withOpening = statement.replace("</w:tr><w:tr><w:tc><w:p><w:r><w:t>01-Jan-2026", `</w:tr>${openingRow.slice("<w:tbl>".length, -"</w:tbl>".length)}<w:tr><w:tc><w:p><w:r><w:t>01-Jan-2026`);
  const ledger = readLedgerXml(withOpening);
  assert.equal(ledger.hasOpeningEntry, true);
  assert.equal(ledger.transactions.length, 3);
  const projection = projectLedger(ledger, cents("1190"), cents("1000"), "last");
  assert.equal(projection.transactions[0].credit, cents("400"));
  assert.equal(projection.transactions[0].balance, cents("1400"));
});

test("a single opening row can only use the same opening and closing balance", () => {
  const source = readLedgerXml(statement);
  const single = { ...source, transactions: source.transactions.slice(0, 1) };
  assert.throws(() => projectLedger(single, cents("1250"), cents("1000"), "last"), /opening and closing balances must match/);
  const projection = projectLedger(single, cents("1250"), null, "spread");
  assert.equal(projection.opening, projection.closing);
  assert.equal(projection.transactions[0].balance, projection.opening);
});

test("preserves cent precision with large amounts and accepts negative balances", () => {
  const ledger = readLedgerXml(statement);
  const large = cents("999,999,999,999,999,999.99");
  for (const target of [large, -large, BigInt(0)]) {
    const projection = projectLedger(ledger, target, ledger.opening, "spread");
    assert.equal(projection.transactions.at(-1)?.balance, target);
    assert.equal(projection.opening + projection.creditTotal - projection.debitTotal, target);
  }
  assert.equal(parseMoney("1,000.001"), null);
  assert.equal(parseMoney("USD 1,000.00", "GHS"), null);
  assert.equal(parseMoney("GHS (1,000.00)", "GHS"), -cents("1000"));
  assert.equal(parseMoney("1,000.00 DR"), -cents("1000"));
  assert.equal(formatMoney(-cents("1000")), "(1,000.00)");
});

test("an undated row and source minus signs do not break cash reconciliation", () => {
  const ledger = readLedgerXml(statement.replace("03-Jan-2026", "date unreadable").replace(">200.00<", ">-200.00<"));
  assert.equal(ledger.transactions[1].date, "date unreadable");
  assert.equal(ledger.transactions[1].debit, cents("200"));
  assert.equal(reconcileLedger(ledger).closing, cents("1190"));
});

test("exports projected values in the uploaded DOCX while retaining its other parts and transaction columns", async () => {
  const decorated = statement
    .replaceAll("<w:tbl>", '<w:tbl><w:tblPr><w:tblStyle w:val="OriginalTableStyle"/></w:tblPr>')
    .replace("<w:body>", '<w:body><w:p><w:r><w:t>Account</w:t><w:t xml:space="preserve"> </w:t><w:t>Summary</w:t></w:r></w:p><w:p><w:r><w:t>Total Dr: </w:t><w:t>210.00</w:t><w:t>Total Cr: </w:t><w:t>400.00</w:t></w:r></w:p>')
    .replace("<w:t>10.00</w:t>", '<w:rPr><w:b/></w:rPr><w:t>10.</w:t><w:t>00</w:t>')
    .replace("</w:body>", `${table([["Total debits", "210.00"], ["Total credits", "400.00"]])}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body>`);
  const source = new JSZip();
  source.file("word/document.xml", decorated);
  source.file("word/styles.xml", "original styles");
  source.file("word/header1.xml", "original header");
  source.file("word/media/logo.png", new Uint8Array([3, 9, 20]));
  const file = new File([await source.generateAsync({ type: "arraybuffer" })], "original.docx");
  const ledger = readLedgerXml(decorated);
  const projection = projectLedger(ledger, cents("1250.01"), cents("1100.00"), "last");
  const copy = await JSZip.loadAsync(await (await createUpdatedLedgerDocument(file, projection)).arrayBuffer());
  const xml = await copy.file("word/document.xml")?.async("string");
  assert.ok(xml);
  assert.equal(await copy.file("word/styles.xml")?.async("string"), "original styles");
  assert.equal(await copy.file("word/header1.xml")?.async("string"), "original header");
  assert.deepEqual(await copy.file("word/media/logo.png")?.async("uint8array"), new Uint8Array([3, 9, 20]));
  assert.match(xml, /OriginalTableStyle/);
  assert.match(xml, /w:w="11906" w:h="16838"/);
  assert.match(xml, /<w:t>Summary \(Projection\)<\/w:t>/);
  assert.match(xml, /<w:t>Total Dr: <\/w:t><w:t>200.00<\/w:t><w:t>Total Cr: <\/w:t><w:t>350.01<\/w:t>/);
  assert.match(xml, /<w:t>987654320<\/w:t>.*<w:t>04-Jan-2026<\/w:t>/);
  assert.equal(xml.split("<w:t>1,100.00</w:t>").length - 1, 3);
  assert.equal(xml.split("<w:t>1,250.01</w:t>").length - 1, 3);
  const updated = readLedgerXml(xml);
  assert.equal(updated.opening, projection.opening);
  assert.equal(updated.transactions[0].balance, updated.opening);
  assert.equal(updated.reportedClosing, projection.closing);
  assert.equal(reconcileLedger(updated).mismatches, 0);
  assert.deepEqual(updated.transactions, projection.transactions.map(({ date, description, debit, credit, balance }) => ({ date, description, debit, credit, balance })));
});

test("writes selected dates and reconciled balances into one copy of the uploaded ledger", async () => {
  const zip = new JSZip();
  zip.file("word/document.xml", statement);
  const file = new File([await zip.generateAsync({ type: "arraybuffer" })], "ledger.docx");
  const extraction = await extractDocument(file);
  const instruction = interpretInstruction("Move dates backward by two months.", 3, "forward");
  const { dates } = findDates(extraction.text, instruction);
  const selected = dates.map((date) => ({ ...date, selected: date.original !== "02-Jan-2026" }));
  assert.equal(selected.filter((date) => date.selected).length, 5);
  const projection = projectLedger(readLedgerXml(statement), cents("1200.01"), cents("1130.25"), "spread");
  const updatedFile = await createUpdatedLedgerDocument(file, projection, selected);
  const copy = await JSZip.loadAsync(await updatedFile.arrayBuffer());
  const xml = await copy.file("word/document.xml")?.async("string");
  assert.ok(xml);
  const updated = readLedgerXml(xml);

  assert.deepEqual(updated.transactions.map((row) => row.date), ["01-Nov-2025", "03-Nov-2025", "04-Nov-2025"]);
  assert.match(xml, /<w:t>123456789<\/w:t>.*<w:t>02-Jan-2026<\/w:t>/);
  assert.match(xml, /<w:t>987654320<\/w:t>.*<w:t>04-Nov-2025<\/w:t>/);
  assert.equal(updated.transactions[0].balance, projection.opening);
  assert.equal(updated.transactions.at(-1)?.balance, projection.closing);
  assert.equal(updated.reportedClosing, projection.closing);
  assert.equal(reconcileLedger(updated).mismatches, 0);
});
