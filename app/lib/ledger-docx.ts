import JSZip from "jszip";
import { formatMoney, parseMoney, type Ledger, type LedgerProjection, type LedgerTransaction } from "./ledger";

const wordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

function elements(parent: Element, name: string) {
  return Array.from(parent.getElementsByTagNameNS(wordNamespace, name));
}

function children(parent: Element, name: string) {
  return Array.from(parent.childNodes).filter((node): node is Element =>
    node.nodeType === 1 && (node as Element).namespaceURI === wordNamespace && (node as Element).localName === name,
  );
}

function cellText(cell: Element) {
  return elements(cell, "p").map((paragraph) => elements(paragraph, "t").map((run) => run.textContent ?? "").join("")).join(" ").replace(/\s+/g, " ").trim();
}

function headerColumns(row: string[]) {
  const labels = row.map((cell) => cell.toLowerCase().replace(/[.()]/g, " ").trim());
  const date = labels.findIndex((label) => /^(?:transaction |posting |booking )?date\b/.test(label));
  const debit = labels.findIndex((label) => /^(?:debits?|withdrawals?|money out|payments out)\b/.test(label));
  const credit = labels.findIndex((label) => /^(?:credits?|deposits?|money in|payments in)\b/.test(label));
  const balance = labels.findIndex((label) => /^(?:(?:running|available) )?balance\b/.test(label));
  if ([date, debit, credit, balance].some((index) => index < 0)) return null;
  const description = labels.findIndex((label) => /^(?:description|details|narration|particulars|transaction details)\b/.test(label));
  return { date, debit, credit, balance, description, width: row.length };
}

function summaryAmount(rows: string[][], label: RegExp, currency: string) {
  const matches = rows.filter((row) => row.length === 2 && label.test(row[0]));
  const amounts = matches.map((row) => parseMoney(row[1], currency));
  if (!amounts.length) return null;
  if (amounts[0] === null || amounts.some((amount) => amount !== amounts[0])) {
    throw new Error("The Word file contains unreadable or conflicting summary balances. Use a document for one account and statement period.");
  }
  return amounts[0];
}

const emptyAmount = /^(?:[-–—]|\s*)$/;
const transactionDate = /^(?:\d{1,4}\s*[-./– ]\s*[\p{L}\d]{1,9}\s*[-./–, ]\s*\d{1,4}|[\p{L}]{3,9}\s+\d{1,2},?\s+\d{4})$/u;

export function readLedgerXml(xml: string): Ledger {
  const document = new DOMParser().parseFromString(xml, "application/xml");
  if (!document.documentElement || document.getElementsByTagName("parsererror").length) throw new Error("The Word file contains unreadable XML.");
  const tables = elements(document.documentElement, "tbl").map((table) =>
    children(table, "tr").map((row) => children(row, "tc").map(cellText)),
  );
  const summary = tables.flat().filter((row) => row.length === 2);
  const currencies = new Set(summary.filter((row) => /^currency:?$/i.test(row[0])).map((row) => row[1].toUpperCase()).filter(Boolean));
  if (currencies.size > 1) throw new Error("Use a statement with one currency per projection.");
  const currency = [...currencies][0] ?? "";
  const opening = summaryAmount(summary, /^(?:opening balance|balance brought forward):?$/i, currency);
  const reportedClosing = summaryAmount(summary, /^(?:closing balance|ending balance|balance carried forward):?$/i, currency);
  const transactions: LedgerTransaction[] = [];
  let columns: ReturnType<typeof headerColumns> = null;

  for (const rows of tables) {
    for (const row of rows) {
      const header = headerColumns(row);
      if (header) {
        columns = header;
        continue;
      }
      if (!columns || row.length !== columns.width || !row.some(Boolean)) continue;
      if (!transactionDate.test(row[columns.date])) {
        if (row[columns.date] && !/\b(?:opening|closing|subtotals?|totals?|balance)\b/i.test(row[columns.date])) {
          throw new Error("A transaction row has an unreadable date. Check the date column before calculating.");
        }
        continue;
      }
      const debitCell = row[columns.debit];
      const creditCell = row[columns.credit];
      const balanceCell = row[columns.balance];
      const debit = emptyAmount.test(debitCell) ? BigInt(0) : parseMoney(debitCell, currency);
      const credit = emptyAmount.test(creditCell) ? BigInt(0) : parseMoney(creditCell, currency);
      const balance = emptyAmount.test(balanceCell) ? null : parseMoney(balanceCell, currency);
      if (debit === null || debit < 0 || credit === null || credit < 0 || (!emptyAmount.test(balanceCell) && balance === null)) {
        throw new Error(`Transaction ${transactions.length + 1} has an unreadable amount. Use positive entries in the debit and credit columns.`);
      }
      transactions.push({ date: row[columns.date], description: row[columns.description] ?? "", debit, credit, balance });
    }
  }

  const first = transactions[0];
  if (!first) throw new Error("No transaction table was found. Use a Word file with Date, Debit, Credit, and Balance columns.");
  const inferredOpening = first.balance === null ? null : first.balance - first.credit + first.debit;
  if (opening === null && inferredOpening === null) throw new Error("The opening balance could not be found or inferred from the first transaction.");
  return { opening: (opening ?? inferredOpening)!, reportedClosing, currency, transactions };
}

export async function readLedger(file: File) {
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const body = zip.file("word/document.xml");
  if (!body) throw new Error("This file does not contain a readable Word document body.");
  return readLedgerXml(await body.async("string"));
}

function escapeXml(text: string) {
  return text.replace(/[&<>]/g, (character) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;" })[character]!);
}

function paragraph(text: string, bold = false, align = "left", size = 22) {
  return `<w:p><w:pPr><w:jc w:val="${align}"/><w:spacing w:after="80"/></w:pPr><w:r><w:rPr>${bold ? "<w:b/>" : ""}<w:sz w:val="${size}"/></w:rPr><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`;
}

export async function createLedgerReport(ledger: Ledger, projection: LedgerProjection) {
  const currency = ledger.currency ? ` (${ledger.currency})` : "";
  const widths = [600, 1580, 3260, 1350, 1350, 1500];
  const row = (cells: string[], header = false, final = false) => `<w:tr><w:trPr><w:cantSplit/>${header ? "<w:tblHeader/>" : ""}</w:trPr>${cells.map((cell, index) => `<w:tc><w:tcPr><w:tcW w:w="${widths[index]}" w:type="dxa"/>${header || final ? '<w:shd w:fill="EDF1F7"/>' : ""}</w:tcPr>${paragraph(cell, header || final, index >= 3 ? "right" : "left", 20)}</w:tc>`).join("")}</w:tr>`;
  const table = `<w:tbl><w:tblPr><w:tblW w:w="9640" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblBorders><w:bottom w:val="single" w:sz="4" w:color="CBD3E0"/><w:insideH w:val="single" w:sz="4" w:color="CBD3E0"/></w:tblBorders></w:tblPr><w:tblGrid>${widths.map((width) => `<w:gridCol w:w="${width}"/>`).join("")}</w:tblGrid>${row(["#", "Date", "Description", "Debit", "Credit", "Balance"], true)}${projection.transactions.map((transaction, index) => row([
    String(index + 1), transaction.date, transaction.description, formatMoney(transaction.debit), formatMoney(transaction.credit), formatMoney(transaction.balance),
  ], false, index === projection.transactions.length - 1)).join("")}</w:tbl>`;
  const summary = [
    `Opening balance${currency}  ${formatMoney(projection.opening)}`,
    `Total debits${currency}  ${formatMoney(projection.debitTotal)}`,
    `Total credits${currency}  ${formatMoney(projection.creditTotal)}`,
    `Closing balance${currency}  ${formatMoney(projection.closing)}`,
  ].map((line) => paragraph(line, true)).join("");
  const zip = new JSZip();
  zip.file("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>');
  zip.file("_rels/.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file("word/_rels/document.xml.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>');
  zip.file("word/header1.xml", `<w:hdr xmlns:w="${wordNamespace}">${paragraph("DateShift transaction projection", true, "right", 18)}</w:hdr>`);
  zip.file("word/document.xml", `<w:document xmlns:w="${wordNamespace}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${paragraph("Transaction projection", true, "left", 36)}${paragraph("Planning scenario based on the selected target and the uploaded transaction list.")}${summary}${table}<w:sectPr><w:headerReference w:type="default" r:id="rId1"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1260" w:right="1300" w:bottom="1260" w:left="1300" w:header="630" w:footer="630"/></w:sectPr></w:body></w:document>`);
  return zip.generateAsync({ type: "blob", compression: "DEFLATE", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
}
