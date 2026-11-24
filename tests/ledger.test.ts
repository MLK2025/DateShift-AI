import assert from "node:assert/strict";
import { createRequire } from "node:module";
import { dirname } from "node:path";
import test from "node:test";
import JSZip from "jszip";
import { createLedgerReport, readLedgerXml } from "../app/lib/ledger-docx";
import { parseMoney, projectLedger, reconcileLedger } from "../app/lib/ledger";

const require = createRequire(import.meta.url);
const mammothDirectory = dirname(require.resolve("mammoth/package.json"));
globalThis.DOMParser = require(require.resolve("@xmldom/xmldom", { paths: [mammothDirectory] })).DOMParser;

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

test("each method reconciles exactly and the report contains the projected final balance", async () => {
  const ledger = readLedgerXml(statement);
  const target = cents("1250.01");
  const opening = cents("1000.00");
  for (const [specified, method] of [[null, "spread"], [opening, "spread"], [opening, "last"]] as const) {
    const projection = projectLedger(ledger, target, specified, method);
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
  assert.equal(projection.transactions[2].credit, cents("50.01"));
  const report = await JSZip.loadAsync(await (await createLedgerReport(ledger, projection)).arrayBuffer());
  const xml = await report.file("word/document.xml")?.async("string");
  assert.ok(xml);
  assert.match(xml, /Transaction projection/);
  assert.ok(xml.includes("Closing balance (GHS)  1,250.01"));
  const parsed = new DOMParser().parseFromString(xml, "application/xml");
  const rows = Array.from(parsed.getElementsByTagNameNS(wordNamespace, "tr"));
  assert.equal(rows.at(-1)?.getElementsByTagNameNS(wordNamespace, "t").item(5)?.textContent, "1,250.01");
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
});
