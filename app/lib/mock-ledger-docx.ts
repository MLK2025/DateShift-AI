import { cellText, directTextRuns, indexLedgerXml, setLedgerCellText, writeLedgerAmounts } from "./ledger-docx";
import { reconcileLedger } from "./ledger";
import { generateMockLedger, templateDate } from "./mock-ledger";
import type { MockPattern } from "./mock-transactions";
import { addSampleFooter } from "./sample-docx-footer";
import { openWordDocument, replaceTextInRuns } from "./word-text";

const word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const docxMime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

function updateDateCell(cell: Element, date: Date, text: string) {
  setLedgerCellText(cell, text);
  for (const control of Array.from(cell.getElementsByTagNameNS(word, "date"))) {
    if (control.hasAttributeNS(word, "fullDate")) control.setAttributeNS(word, "w:fullDate", date.toISOString().replace(".000Z", "Z"));
  }
}

function formatPeriod(value: string, start: Date, end: Date) {
  const range = /^(.*?(?:Statement Period\s*:?\s*)?)(\d[^\r\n]*?)(\s+(?:[-–—]|to)\s+)(\d[^\r\n]*)$/i.exec(value);
  if (!range) return null;
  const first = templateDate(range[2]);
  const last = templateDate(range[4]);
  return first && last ? range[1] + first.format(start) + range[3] + last.format(end) : null;
}

function writeStatementPeriod(document: Document, start: Date, end: Date) {
  const body = document.getElementsByTagNameNS(word, "body")[0];
  if (!body) return;
  const label = /^(?:Statement Period|Period)\s*:?\s*$/i;
  for (const row of Array.from(body.getElementsByTagNameNS(word, "tr"))) {
    const cells = Array.from(row.getElementsByTagNameNS(word, "tc")).filter((cell) => cell.parentNode === row);
    if (cells.length !== 2 || !label.test(cellText(cells[0]))) continue;
    const value = formatPeriod(cellText(cells[1]), start, end);
    if (value) setLedgerCellText(cells[1], value);
  }

  const paragraphs = Array.from(body.getElementsByTagNameNS(word, "p")).filter((p) => p.parentNode === body);
  for (const [index, paragraph] of paragraphs.entries()) {
    const text = directTextRuns(paragraph).map((node) => node.textContent ?? "").join("");
    if (!/^\s*Statement Period\b/i.test(text)) continue;
    const nearby = label.test(text.trim()) ? paragraphs.slice(index + 1, index + 4) : [paragraph];
    for (const candidate of nearby) {
      const runs = directTextRuns(candidate);
      let value = "";
      const starts = runs.map((run) => {
        const offset = value.length;
        value += run.textContent ?? "";
        return offset;
      });
      const replacement = formatPeriod(value, start, end);
      if (!replacement) continue;
      if (replacement !== value) replaceTextInRuns(runs, starts, 0, value.length, replacement);
      break;
    }
  }
}

export async function createSampleLedgerDocument(file: File, company: string, pattern: MockPattern, batch: number) {
  const { zip, xml } = await openWordDocument(file);
  const { document, ledger, cashRows, summaries } = indexLedgerXml(xml);
  if (cashRows.some(({ description }) => !description)) {
    throw new Error("The Word transaction table needs a Description or Narration column.");
  }
  const plan = generateMockLedger(ledger, pattern, company, batch);
  for (const [index, cells] of cashRows.entries()) {
    const row = plan.transactions[index];
    updateDateCell(cells.date, row.dateValue, row.date);
    setLedgerCellText(cells.description!, row.description, { singleRun: true });
    if (cells.reference) setLedgerCellText(cells.reference, row.reference);
    if (cells.valueDate) {
      const style = templateDate(cellText(cells.valueDate));
      updateDateCell(cells.valueDate, row.dateValue, style ? style.format(row.dateValue) : row.date);
    }
  }
  writeLedgerAmounts(document, cashRows, summaries, plan, true);
  writeStatementPeriod(document, plan.startDate, plan.endDate);
  await addSampleFooter(zip, document);

  const resultXml = new XMLSerializer().serializeToString(document);
  const updated = indexLedgerXml(resultXml);
  const sameRows = updated.ledger.transactions.length === plan.transactions.length && updated.ledger.transactions.every((row, index) => {
    const planned = plan.transactions[index];
    const reference = updated.cashRows[index].reference;
    return row.date === planned.date && row.description === planned.description
      && row.debit === planned.debit && row.credit === planned.credit && row.balance === planned.balance
      && (!reference || cellText(reference) === planned.reference);
  });
  const totals = reconcileLedger(updated.ledger);
  if (!sameRows || totals.mismatches !== 0 || updated.ledger.opening !== plan.opening || totals.closing !== plan.closing
    || (summaries.closing.length > 0 && updated.ledger.reportedClosing !== plan.closing)) {
    throw new Error("The exported table could not be verified. Use a Word template with editable transaction cells.");
  }
  zip.file("word/document.xml", resultXml);
  return {
    blob: await zip.generateAsync({ type: "blob", compression: "DEFLATE", mimeType: docxMime }),
    fileName: `${file.name.replace(/(-SAMPLE)?\.docx$/i, "")}-SAMPLE.docx`,
    rowCount: plan.transactions.length,
    payrollCount: plan.transactions.filter((row) => row.category === "Payroll").length,
  };
}
