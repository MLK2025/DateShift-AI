import assert from "node:assert/strict";
import { createRequire } from "node:module";
import { dirname } from "node:path";
import test from "node:test";
import JSZip from "jszip";
import { indexLedgerXml } from "../app/lib/ledger-docx";
import { reconcileLedger } from "../app/lib/ledger";
import { createSampleLedgerDocument } from "../app/lib/mock-ledger-docx";

const require = createRequire(import.meta.url);
const mammoth = dirname(require.resolve("mammoth/package.json"));
Object.assign(globalThis, require(require.resolve("@xmldom/xmldom", { paths: [mammoth] })));
const word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const office = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const relationships = "http://schemas.openxmlformats.org/package/2006/relationships";
const types = "http://schemas.openxmlformats.org/package/2006/content-types";

function table(rows: string[][]) {
  return `<w:tbl><w:tblPr><w:tblStyle w:val="SourceStyle"/></w:tblPr>${rows.map((row) => `<w:tr>${row.map((cell) => `<w:tc><w:tcPr><w:shd w:fill="E3E8EA"/></w:tcPr><w:p><w:r><w:t>${cell}</w:t></w:r></w:p></w:tc>`).join("")}</w:tr>`).join("")}</w:tbl>`;
}

async function template(withFooter: boolean, withReference = true) {
  const zip = new JSZip();
  const summaries = table([["Currency", "GHS"], ["Opening Balance", "100.00"], ["Closing Balance", "110.00"], ["Total Debits", "30.00"], ["Total Credits", "40.00"]]);
  const headers = withReference
    ? ["Date", "Description", "Ref / Chq No.", "Value Date", "Debit", "Credit", "Balance"]
    : ["Transaction Date", "Value Date", "Transaction Description", "Debits", "Credits", "Balance"];
  const movements = ["02-Jan-2028", "18-Jan-2028", "23-Jan-2028", "03-Feb-2028", "11-Feb-2028", "10-Mar-2028", "11-Mar-2028"];
  const transactions = movements.map((date, index) => {
    const original = ["OLD_REF", "ORIGINAL;DETAIL", "ORIGINAL;DETAIL;CODE", "ORIGINAL / DETAIL", "ORIGINAL : DETAIL", "ORIGINAL : DETAIL", "POS 26"][index];
    return withReference ? [date, original, `SOURCE${index}`, date, "0.00", "10.00", "110.00"]
      : [date, date, original, "0.00", "10.00", "110.00"];
  });
  const transTables = table([headers, ...transactions.slice(0, 4)]) + table(transactions.slice(4));
  const section = withFooter ? '<w:footerReference w:type="default" r:id="rId7"/>' : '<w:pgMar w:footer="0"/><w:titlePg/>';
  zip.file("word/document.xml", `<w:document xmlns:w="${word}" xmlns:r="${office}"><w:body><w:p><w:r><w:t>Statement Period</w:t></w:r></w:p><w:p><w:r><w:t>02-Jan-2028 - 11-Mar-2028</w:t></w:r></w:p>${summaries}${transTables}<w:sectPr>${section}</w:sectPr></w:body></w:document>`);
  zip.file("[Content_Types].xml", `<Types xmlns="${types}"/>`);
  zip.file("word/styles.xml", "preserve this part");
  if (withFooter) {
    zip.file("word/_rels/document.xml.rels", `<Relationships xmlns="${relationships}"><Relationship Id="rId7" Type="${office}/footer" Target="footer1.xml"/></Relationships>`);
    zip.file("word/footer1.xml", `<w:ftr xmlns:w="${word}"><w:p><w:r><w:t>Original footer text</w:t></w:r></w:p></w:ftr>`);
  } else {
    zip.file("word/settings.xml", `<w:settings xmlns:w="${word}"><w:evenAndOddHeaders/></w:settings>`);
  }
  return new File([await zip.generateAsync({ type: "arraybuffer" })], "template.docx");
}

test("writes synthetic month-end payroll and reconciled figures into the template with a visible footer", async () => {
  const file = await template(true);
  const result = await createSampleLedgerDocument(file, "SAMPLE SMP Example; Co", "salary", 1);
  assert.equal(result.fileName, "template-SAMPLE.docx");
  assert.equal(result.rowCount, 7);
  assert.equal(result.payrollCount, 3);
  const zip = await JSZip.loadAsync(await result.blob.arrayBuffer());
  const xml = await zip.file("word/document.xml")?.async("string");
  assert.ok(xml);
  assert.match(xml, /01-Jan-2028 - 31-Mar-2028/);
  assert.doesNotMatch(xml, /ORIGINAL|OLD_REF|SOURCE\d/);
  assert.equal(xml.split("<w:tbl>").length - 1, 3);
  assert.match(xml, /w:val="SourceStyle"/);
  const { ledger, summaries, cashRows } = indexLedgerXml(xml);
  const rows = ledger.transactions;
  const payroll = rows.filter((row) => /example co/i.test(row.description));
  assert.deepEqual(payroll.map((row) => row.date), ["31-Jan-2028", "29-Feb-2028", "31-Mar-2028"]);
  assert.deepEqual(payroll.map((row) => row.description),
    ["STAFF PAY Example Co;020003;", "STAFF PAY Example Co;020005;", "STAFF PAY Example Co;020007;"]);
  assert.ok(payroll.every((row) => row.credit > BigInt(0)));
  assert.equal(new Set(rows.map((row) => row.description)).size, 7);
  assert.ok(rows[0].description.includes("_"));
  assert.equal(rows[2].description.split(";").length, 3);
  assert.ok(rows[3].description.includes("/"));
  assert.ok(rows[5].description.includes(" : "));
  assert.doesNotMatch(xml, /\b(?:sample|smp)(?=\b|\d)/i);
  const refs = cashRows.map(({ reference }) => reference?.textContent);
  assert.equal(new Set(refs).size, rows.length);
  assert.ok(refs.every((ref) => /^[A-Z]{3}\d{6}$/.test(ref ?? "")));
  assert.equal(rows[0].balance, ledger.opening);
  assert.equal(ledger.reportedClosing, rows.at(-1)?.balance);
  const totals = reconcileLedger(ledger);
  assert.equal(totals.mismatches, 0);
  assert.equal(totals.closing, ledger.opening + totals.credit - totals.debit);
  assert.equal(summaries.debit.length, 1);
  assert.equal(await zip.file("word/styles.xml")?.async("string"), "preserve this part");
  const footer = await zip.file("word/footer1.xml")?.async("string");
  assert.match(footer ?? "", /Original footer text[\s\S]*<w:t>SAMPLE<\/w:t>/);

  const repeat = await createSampleLedgerDocument(new File([result.blob], result.fileName), "Different Co", "non-salary", 2);
  const repeatedZip = await JSZip.loadAsync(await repeat.blob.arrayBuffer());
  assert.equal(repeat.fileName, "template-SAMPLE.docx");
  assert.equal(repeat.payrollCount, 0);
  assert.equal((await repeatedZip.file("word/footer1.xml")?.async("string"))?.split("<w:t>SAMPLE</w:t>").length, 2);
});

test("updates Transaction Description tables with no reference column and adds a footer to first and even pages", async () => {
  const result = await createSampleLedgerDocument(await template(false, false), "", "non-salary", 3);
  assert.equal(result.payrollCount, 0);
  const zip = await JSZip.loadAsync(await result.blob.arrayBuffer());
  const xml = await zip.file("word/document.xml")?.async("string");
  assert.ok(xml);
  const { ledger, cashRows } = indexLedgerXml(xml);
  assert.equal(ledger.transactions.length, 7);
  assert.equal(reconcileLedger(ledger).mismatches, 0);
  assert.ok(cashRows.every((row) => !row.reference && row.valueDate));
  assert.doesNotMatch(xml, /\b(?:sample|smp)(?=\b|\d)/i);
  for (const variant of ["default", "first", "even"]) assert.match(xml, new RegExp(`<w:footerReference w:type="${variant}"`));
  assert.match(xml, /<w:pgMar w:footer="120"/);
  assert.match(await zip.file("word/dateshift-sample-footer1.xml")?.async("string") ?? "", /<w:t>SAMPLE<\/w:t>/);
  assert.match(await zip.file("[Content_Types].xml")?.async("string") ?? "", /wordprocessingml\.footer\+xml/);
});
