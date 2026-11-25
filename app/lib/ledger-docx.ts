import JSZip from "jszip";
import { formatMoney, parseMoney, type Ledger, type LedgerProjection, type LedgerTransaction } from "./ledger";
import { openWordDocument, replaceTextInRuns } from "./word-text";
import type { DocumentChange } from "./document-export";

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

type CashRow = { debit: Element; credit: Element; balance: Element };
type Summary = Record<"opening" | "closing" | "debit" | "credit", Element[]>;

const summaryLabels = {
  opening: /^(?:opening balance|balance brought forward):?$/i,
  closing: /^(?:closing balance|ending balance|balance carried forward):?$/i,
  debit: /^(?:total debits?|total withdrawals?):?$/i,
  credit: /^(?:total credits?|total deposits?):?$/i,
};

function summaryAmount(cells: Element[], currency: string) {
  const amounts = cells.map(cellText).filter((value) => !emptyAmount.test(value)).map((value) => parseMoney(value, currency));
  if (!amounts.length) return null;
  if (amounts[0] === null || amounts.some((amount) => amount !== amounts[0])) {
    throw new Error("The Word file contains unreadable or conflicting summary balances. Use a document for one account and statement period.");
  }
  return amounts[0];
}

const emptyAmount = /^(?:[-–—]|\s*)$/;

export function hasTransactionColumns(text: string) {
  return /\b(?:debits?|withdrawals?|money out)\b[\s\S]{0,45}\b(?:credits?|deposits?|money in)\b[\s\S]{0,45}\bbalance\b/i.test(text);
}

function indexLedgerXml(xml: string) {
  const document = new DOMParser().parseFromString(xml, "application/xml");
  if (!document.documentElement || document.getElementsByTagName("parsererror").length) throw new Error("The Word file contains unreadable XML.");
  const tables = elements(document.documentElement, "tbl").map((table) =>
    children(table, "tr").map((row) => {
      const cells = children(row, "tc");
      return { cells, values: cells.map(cellText) };
    }),
  );
  const summaryRows = tables.flat().filter(({ cells }) => cells.length === 2);
  const currencies = new Set(summaryRows.filter(({ values }) => /^currency:?$/i.test(values[0])).map(({ values }) => values[1].toUpperCase()).filter(Boolean));
  if (currencies.size > 1) throw new Error("Use a statement with one currency per projection.");
  const currency = [...currencies][0] ?? "";
  const summaries: Summary = { opening: [], closing: [], debit: [], credit: [] };
  for (const [key, label] of Object.entries(summaryLabels) as [keyof Summary, RegExp][]) {
    summaries[key] = summaryRows.filter(({ values }) => label.test(values[0])).map(({ cells }) => cells[1]);
  }
  const transactions: LedgerTransaction[] = [];
  const cashRows: CashRow[] = [];
  let columns: ReturnType<typeof headerColumns> = null;
  let hasOpeningEntry = false;

  for (const rows of tables) {
    for (const { cells, values: row } of rows) {
      const header = headerColumns(row);
      if (header) {
        columns = header;
        continue;
      }
      if (!columns || row.length !== columns.width || !row.some(Boolean)) continue;
      const debitCell = row[columns.debit];
      const creditCell = row[columns.credit];
      const balanceCell = row[columns.balance];
      const date = row[columns.date];
      const description = row[columns.description] ?? "";
      if (!transactions.length && [description, date].some((label) => summaryLabels.opening.test(label)) && [debitCell, creditCell].every((cell) => emptyAmount.test(cell) || parseMoney(cell, currency) === BigInt(0))) {
        summaries.opening.push(cells[columns.balance]);
        hasOpeningEntry = true;
        continue;
      }
      if ([debitCell, creditCell, balanceCell].every((cell) => emptyAmount.test(cell))) continue;
      if (!date || !/\d/.test(date)) {
        const label = `${date} ${description}`.trim();
        if (summaryLabels.opening.test(label)) { summaries.opening.push(cells[columns.balance]); continue; }
        if (summaryLabels.closing.test(label)) { summaries.closing.push(cells[columns.balance]); continue; }
        if (/^(?:grand\s+)?totals?:?$/i.test(label)) {
          if (!emptyAmount.test(debitCell)) summaries.debit.push(cells[columns.debit]);
          if (!emptyAmount.test(creditCell)) summaries.credit.push(cells[columns.credit]);
          if (!emptyAmount.test(balanceCell)) summaries.closing.push(cells[columns.balance]);
          continue;
        }
        if (/\b(?:subtotals?|totals?|carried forward|brought forward)\b/i.test(label)) continue;
      }
      const debit = emptyAmount.test(debitCell) ? BigInt(0) : parseMoney(debitCell, currency);
      const credit = emptyAmount.test(creditCell) ? BigInt(0) : parseMoney(creditCell, currency);
      const balance = emptyAmount.test(balanceCell) ? null : parseMoney(balanceCell, currency);
      if (debit === null || credit === null || (!emptyAmount.test(balanceCell) && balance === null)) {
        throw new Error(`Transaction ${transactions.length + 1} has an unreadable debit, credit, or balance.`);
      }
      // The debit column defines outflow even in source files that print debits with a minus sign.
      const outflow = (debit < 0 ? -debit : debit) + (credit < 0 ? -credit : BigInt(0));
      transactions.push({ date, description, debit: outflow, credit: credit < 0 ? BigInt(0) : credit, balance });
      cashRows.push({ debit: cells[columns.debit], credit: cells[columns.credit], balance: cells[columns.balance] });
    }
  }

  const first = transactions[0];
  if (!first) throw new Error("No transaction table was found. Use a Word file with Date, Debit, Credit, and Balance columns.");
  const opening = summaryAmount(summaries.opening, currency);
  const reportedClosing = summaryAmount(summaries.closing, currency);
  const inferredOpening = first.balance === null ? null : first.balance - first.credit + first.debit;
  if (opening === null && inferredOpening === null) throw new Error("The opening balance could not be found or inferred from the first transaction.");
  const ledger: Ledger = { opening: (opening ?? inferredOpening)!, reportedClosing, currency, hasOpeningEntry, transactions };
  return { document, ledger, cashRows, summaries };
}

export function readLedgerXml(xml: string): Ledger {
  return indexLedgerXml(xml).ledger;
}

export async function readLedger(file: File) {
  const { xml } = await openWordDocument(file);
  return readLedgerXml(xml);
}

const docxMime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

function directTextRuns(paragraph: Element) {
  return children(paragraph, "r").flatMap((run) => children(run, "t"));
}

function inlineTotals(document: Document) {
  const body = elements(document.documentElement, "body")[0];
  const label = /(?<![A-Za-z])Total[ \t\u00a0]+(Dr|Cr|Debit[s]?|Credit[s]?)[ \t\u00a0]*:[ \t\u00a0]*([-−]?\(?\d[\d,]*(?:\.\d{1,2})?\)?)(?![\d.])/gi;
  return children(body, "p").map((paragraph) => {
    const nodes = directTextRuns(paragraph);
    let text = "";
    const starts = nodes.map((node) => {
      const start = text.length;
      text += node.textContent ?? "";
      return start;
    });
    const matches = [...text.matchAll(label)].map((match) => ({
      kind: match[1].toLowerCase().startsWith("d") ? "debit" as const : "credit" as const,
      start: match.index + match[0].length - match[2].length,
      value: match[2],
    }));
    return { nodes, starts, matches };
  });
}

function cashCellValue(cell: Element, amount: bigint, keepBlank: boolean) {
  const prior = cellText(cell);
  if (keepBlank && emptyAmount.test(prior) && amount === BigInt(0)) return prior;
  if (!/[-−]/.test(prior) && parseMoney(prior) === amount) return prior;
  const currencyPrefix = prior.match(/^[(-]?\s*([₵$€£]|[A-Z]{3})(?=\s*[(\d-])/i)?.[1];
  const amountText = formatMoney(amount);
  return currencyPrefix ? `${currencyPrefix} ${amountText}` : amountText;
}

function setCashCell(cell: Element, amount: bigint, keepBlank = false) {
  const next = cashCellValue(cell, amount, keepBlank);
  if (next === cellText(cell)) return;
  if (elements(cell, "fldSimple").length || elements(cell, "instrText").length || elements(cell, "dataBinding").length) {
    throw new Error("A balance or transaction amount is a Word field. Convert that field to regular text before exporting.");
  }

  const runs = elements(cell, "t");
  const existing = runs.filter((node) => Boolean(node.textContent?.trim()));
  if (existing.length) {
    // Keep the original paragraphs and character styles. Extra characters
    // inherit the last run's formatting, as in the date editor.
    let rest = next;
    existing.forEach((node, index) => {
      const length = node.textContent?.length ?? 0;
      node.textContent = index === existing.length - 1 ? rest : rest.slice(0, length);
      if (/^\s|\s$/.test(node.textContent)) node.setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:space", "preserve");
      rest = rest.slice(length);
    });
    return;
  }
  if (runs[0]) { runs[0].textContent = next; return; }
  const document = cell.ownerDocument!;
  const paragraph = elements(cell, "p")[0];
  if (!paragraph) throw new Error("A transaction amount has no editable Word paragraph.");
  const run = children(paragraph, "r").at(-1) ?? paragraph.appendChild(document.createElementNS(wordNamespace, "w:r"));
  const text = run.appendChild(document.createElementNS(wordNamespace, "w:t"));
  text.textContent = next;
}

function markProjection(document: Document, currency: string, projection: LedgerProjection, summaries: Summary) {
  const body = elements(document.documentElement, "body")[0];
  if (!body) throw new Error("The Word file has no editable document body.");
  const label = "DateShift transaction projection";
  const note = [label];
  const currencyLabel = currency ? `${currency} ` : "";
  if (!summaries.opening.length) note.push(`Opening balance ${currencyLabel}${formatMoney(projection.opening)}`);
  if (!summaries.closing.length) note.push(`Closing balance ${currencyLabel}${formatMoney(projection.closing)}`);
  const paragraphs = children(body, "p");
  for (const paragraph of paragraphs) {
    const nodes = directTextRuns(paragraph);
    const text = nodes.map((node) => node.textContent ?? "").join("");
    const oldHeading = /^(\s*Account Summary)(\s+\(Projection\))\s*$/i.exec(text);
    if (!oldHeading) continue;

    let offset = 0;
    const starts = nodes.map((node) => {
      const start = offset;
      offset += node.textContent?.length ?? 0;
      return start;
    });
    replaceTextInRuns(nodes, starts, oldHeading[1].length, oldHeading[1].length + oldHeading[2].length, "");
  }
  const existing = paragraphs.find((p) => directTextRuns(p)[0]?.textContent?.startsWith(label));
  if (existing) {
    directTextRuns(existing)[0].textContent = note.join(". ");
    return;
  }
  const marker = document.createElementNS(wordNamespace, "w:p");
  const paragraphProperties = marker.appendChild(document.createElementNS(wordNamespace, "w:pPr"));
  const spacing = paragraphProperties.appendChild(document.createElementNS(wordNamespace, "w:spacing"));
  spacing.setAttributeNS(wordNamespace, "w:before", "0");
  spacing.setAttributeNS(wordNamespace, "w:after", "0");
  const run = marker.appendChild(document.createElementNS(wordNamespace, "w:r"));
  const properties = run.appendChild(document.createElementNS(wordNamespace, "w:rPr"));
  properties.appendChild(document.createElementNS(wordNamespace, "w:b"));
  const size = properties.appendChild(document.createElementNS(wordNamespace, "w:sz"));
  size.setAttributeNS(wordNamespace, "w:val", "20");
  run.appendChild(document.createElementNS(wordNamespace, "w:t")).textContent = note.join(". ");
  body.insertBefore(marker, children(body, "sectPr")[0] ?? null);
}

export async function createUpdatedLedgerDocument(file: File, projection: LedgerProjection, dateChanges: DocumentChange[] = []) {
  const { zip, xml } = await openWordDocument(file);
  let sourceXml = xml;
  const selectedDates = dateChanges.filter((change) => change.selected);
  if (selectedDates.length) {
    const { updateDocumentXml } = await import("./document-export");
    const result = updateDocumentXml(xml, selectedDates);
    if (result.appliedCount !== selectedDates.length) throw new Error("Some selected dates could not be written. Review the date selection before downloading.");
    sourceXml = result.xml;
  }
  const { document, ledger, cashRows, summaries } = indexLedgerXml(sourceXml);
  if (cashRows.length !== projection.transactions.length) throw new Error("The Word transaction table no longer matches the projection.");

  for (const [index, row] of cashRows.entries()) {
    const amount = projection.transactions[index];
    setCashCell(row.debit, amount.debit, true);
    setCashCell(row.credit, amount.credit, true);
    setCashCell(row.balance, amount.balance);
  }
  const totals = { opening: projection.opening, closing: projection.closing, debit: projection.debitTotal, credit: projection.creditTotal };
  for (const key of Object.keys(summaries) as (keyof Summary)[]) {
    for (const cell of summaries[key]) setCashCell(cell, totals[key]);
  }
  for (const { nodes, starts, matches } of inlineTotals(document)) {
    for (const { kind, start, value } of matches.reverse()) {
      const next = formatMoney(totals[kind]);
      if (value !== next) replaceTextInRuns(nodes, starts, start, start + value.length, next);
    }
  }

  markProjection(document, ledger.currency, projection, summaries);
  const updatedXml = new XMLSerializer().serializeToString(document);
  const updated = readLedgerXml(updatedXml);
  const matches = updated.transactions.length === cashRows.length && updated.transactions.every((row, index) => {
    const expected = projection.transactions[index];
    return row.debit === expected.debit && row.credit === expected.credit && row.balance === expected.balance;
  });
  const firstBalance = updated.hasOpeningEntry ? updated.opening : updated.transactions[0].balance;
  if (!matches || firstBalance !== projection.opening || updated.opening !== projection.opening || (summaries.closing.length && updated.reportedClosing !== projection.closing)) {
    throw new Error("The projected balances could not be verified in the updated Word file.");
  }

  zip.file("word/document.xml", updatedXml);
  return zip.generateAsync({ type: "blob", compression: "DEFLATE", mimeType: docxMime });
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
  zip.file("word/document.xml", `<w:document xmlns:w="${wordNamespace}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${paragraph("Transaction projection", true, "left", 36)}${paragraph("Planning scenario based on the selected target and the uploaded transaction list. Parentheses indicate negative balances.")}${summary}${table}<w:sectPr><w:headerReference w:type="default" r:id="rId1"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1260" w:right="1300" w:bottom="1260" w:left="1300" w:header="630" w:footer="630"/></w:sectPr></w:body></w:document>`);
  return zip.generateAsync({ type: "blob", compression: "DEFLATE", mimeType: docxMime });
}
