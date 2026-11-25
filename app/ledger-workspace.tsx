import { useMemo, useState } from "react";
import { DocxDocumentPreview } from "./docx-preview";
import type { DocumentChange } from "./lib/document-export";
import { formatMoney, parseMoney, projectLedger, reconcileLedger, type Ledger, type LedgerMethod, type LedgerProjection } from "./lib/ledger";

const sampleLedger: Ledger = {
  currency: "GHS",
  opening: BigInt(100_000),
  reportedClosing: BigInt(125_000),
  transactions: [
    { date: "01-Dec-2025", description: "Deposit", debit: BigInt(0), credit: BigInt(40_000), balance: BigInt(140_000) },
    { date: "02-Dec-2025", description: "Supplies", debit: BigInt(15_000), credit: BigInt(0), balance: BigInt(125_000) },
  ],
};

type LedgerDate = DocumentChange & { id: number; context: string; reason: string };
export type LedgerSource = {
  name: string;
  ledger: Ledger;
  target?: string;
  file?: File;
  dateReview?: { changes: LedgerDate[]; summary: string };
};

export function LedgerWorkspace({ initialSource = null, onBackToSettings }: { initialSource?: LedgerSource | null; onBackToSettings?: () => void }) {
  const [source, setSource] = useState<LedgerSource | null>(initialSource);
  const [openingMode, setOpeningMode] = useState<"calculated" | "fixed">("calculated");
  const [opening, setOpening] = useState(() => initialSource ? formatMoney(initialSource.ledger.opening) : "");
  const [closing, setClosing] = useState(() => initialSource?.target ?? (initialSource ? formatMoney(reconcileLedger(initialSource.ledger).closing) : ""));
  const [method, setMethod] = useState<LedgerMethod>("spread");
  const [reading, setReading] = useState(false);
  const [preparing, setPreparing] = useState<"preview" | "download" | null>(null);
  const [dateChanges, setDateChanges] = useState<LedgerDate[]>(initialSource?.dateReview?.changes ?? []);
  const [prepared, setPrepared] = useState<{ file?: File; projection: LedgerProjection; dates: LedgerDate[]; blob: Blob } | null>(null);
  const [error, setError] = useState("");
  const ledger = source?.ledger;
  const reconciliation = useMemo(() => ledger ? reconcileLedger(ledger) : null, [ledger]);
  const target = parseMoney(closing, ledger?.currency);
  const specifiedOpening = openingMode === "fixed" ? parseMoney(opening, ledger?.currency) : null;
  const canCalculate = target !== null && (openingMode !== "fixed" || specifiedOpening !== null);
  const { projection, projectionError } = useMemo(() => {
    if (!ledger || !canCalculate || target === null) return { projection: null, projectionError: "" };
    try {
      return { projection: projectLedger(ledger, target, specifiedOpening, method), projectionError: "" };
    } catch (failure) {
      return { projection: null, projectionError: failure instanceof Error ? failure.message : "The balances could not be calculated." };
    }
  }, [ledger, canCalculate, target, specifiedOpening, method]);
  const preparedFile = prepared && prepared.file === source?.file && prepared.projection === projection && prepared.dates === dateChanges ? prepared.blob : null;
  const dateReview = source?.dateReview;
  const selectedDates = dateChanges.filter((date) => date.selected).length;

  function loadSource(name: string, next: Ledger, file?: File) {
    setSource({ name, ledger: next, file });
    setOpening(formatMoney(next.opening));
    setClosing(formatMoney(reconcileLedger(next).closing));
    setError("");
  }

  async function loadFile(file: File) {
    setSource(null);
    setError("");
    if (!/\.docx$/i.test(file.name) || file.size > 25 * 1024 * 1024) {
      setError("Choose a Word .docx file up to 25 MB.");
      return;
    }
    setReading(true);
    try {
      const { readLedger } = await import("./lib/ledger-docx");
      loadSource(file.name, await readLedger(file), file);
    } catch (failure) {
      setError(failure instanceof Error ? failure.message : "The transaction table could not be read.");
    } finally {
      setReading(false);
    }
  }

  async function prepare(usage: "preview" | "download") {
    if (!source || !ledger || !projection) return;
    setPreparing(usage);
    setError("");
    try {
      const { createLedgerReport, createUpdatedLedgerDocument } = await import("./lib/ledger-docx");
      const blob = preparedFile ?? (source.file ? await createUpdatedLedgerDocument(source.file, projection, dateChanges) : await createLedgerReport(ledger, projection));
      if (usage === "preview") {
        setPrepared({ file: source.file, projection, dates: dateChanges, blob });
        return;
      }
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = `${source.name.replace(/\.docx$/i, "")}-projection.docx`;
      document.body.append(link);
      link.click();
      link.remove();
      window.setTimeout(() => URL.revokeObjectURL(url), 60_000);
    } catch (failure) {
      setError(failure instanceof Error ? failure.message : "The updated Word file could not be created.");
    } finally {
      setPreparing(null);
    }
  }

  const busy = reading || preparing !== null;
  const currency = ledger?.currency;
  const reportedClosing = ledger?.reportedClosing;
  const displayedMoney = (value: bigint) => `${currency ? `${currency} ` : ""}${formatMoney(value)}`;

  return (
    <section className="ledger-workspace" aria-label="Transaction projection workspace">
      <div className="ledger-settings">
        <article className="panel ledger-panel">
          {dateReview ? <>
            <h2>Review date changes</h2>
            <p>{dateReview.summary}</p>
            <p role="status">{dateChanges.length ? `${selectedDates} of ${dateChanges.length} dates selected` : "No dates matched your instructions."}</p>
            {dateChanges.length > 0 && <details className="ledger-date-review">
              <summary>Choose which dates to update</summary>
              <div className="ledger-date-list">{dateChanges.map((date) => <label key={date.id}>
                <input type="checkbox" checked={date.selected} disabled={busy} onChange={(event) => {
                  const selected = event.target.checked;
                  setDateChanges((changes) => changes.map((change) => change.id === date.id ? { ...change, selected } : change));
                }} />
                <span><s>{date.original}</s> → <strong>{date.updated || "—"}</strong><small>{date.context}</small><small>{date.reason}</small></span>
              </label>)}</div>
            </details>}
            <button className="secondary-button" disabled={busy} onClick={onBackToSettings}>Change date settings</button>
          </> : <>
          <h2>Choose a Word ledger</h2>
          <p>Use a .docx with Date, Debit, Credit, and Balance columns. Table continuations are read in document order.</p>
          <input aria-label="Choose a Word ledger" type="file" accept=".docx" disabled={busy} onChange={(event) => {
            const file = event.currentTarget.files?.[0];
            event.currentTarget.value = "";
            if (file) void loadFile(file);
          }} />
          <button className="secondary-button" disabled={busy} onClick={() => loadSource("Example ledger", sampleLedger)}>Use sample ledger</button>
          {reading && <p role="status">Reading the transaction table…</p>}
          </>}
          {source && <p role="status"><strong>{source.name}</strong> · {source.ledger.transactions.length} transactions{currency ? ` · ${currency}` : ""}</p>}
        </article>
        <article className="panel ledger-panel">
          <h2>Set the projected balances</h2>
          <div className="figure-controls ledger-controls">
            <label><span>Target closing balance {currency && `(${currency})`}</span><input value={closing} onChange={(event) => setClosing(event.target.value)} inputMode="decimal" placeholder="50,000.00" disabled={busy} /></label>
            <label><span>Opening balance</span><select value={openingMode} onChange={(event) => setOpeningMode(event.target.value === "fixed" ? "fixed" : "calculated")} disabled={busy}><option value="calculated">Calculate from target and remaining transactions</option><option value="fixed">Set the opening balance and adjust transactions</option></select></label>
            {openingMode === "fixed" && <>
              <label><span>Set opening balance {currency && `(${currency})`}</span><input value={opening} onChange={(event) => setOpening(event.target.value)} inputMode="decimal" placeholder="10,000.00" disabled={busy} /></label>
              <label><span>Apply the adjustment</span><select value={method} onChange={(event) => setMethod(event.target.value === "last" ? "last" : "spread")} disabled={busy}><option value="spread">Spread across transactions proportionally</option><option value="last">Apply to the final transaction</option></select></label>
            </>}
            {ledger && !canCalculate && <p role="status">Enter a valid amount with up to two decimal places in {currency ? `${currency} or without a currency code` : "the balance fields"}.</p>}
            {projectionError && <p className="notice" role="alert">{projectionError}</p>}
          </div>
        </article>
      </div>
      {error && <p className="notice" role="alert">{error}</p>}
      {reconciliation && ledger && <p className="ledger-source">
        <strong>Source reconciliation.</strong> Opening balance {displayedMoney(ledger.opening)}. Calculated closing balance {displayedMoney(reconciliation.closing)}.
        {reportedClosing != null && ` Printed closing balance ${displayedMoney(reportedClosing)}.`}
        {reconciliation.mismatches > 0 && ` ${reconciliation.mismatches} printed running balances differ from the recalculated values.`}
      </p>}
      {projection && ledger && (
        <article className="panel ledger-panel ledger-result">
          <div className="ledger-result-heading">
            <div><h2>{dateReview ? "Dates and transaction balances" : "Transaction projection"}</h2><p>The first row shows the opening balance with zero debit and credit. Later rows account for cash movement and the final balance matches your closing target. Parentheses indicate a negative balance.</p>{source?.file && <p>{dateReview ? "The updated Word file includes the selected dates and projected balances." : "The download writes these figures into a copy of your uploaded Word document."}</p>}</div>
            <div className="ledger-file-actions">
              {dateReview && <button className="secondary-button" disabled={busy || Boolean(preparedFile)} onClick={() => void prepare("preview")}>{preparing === "preview" ? "Preparing preview…" : preparedFile ? "Preview ready" : "Preview updated Word file"}</button>}
              <button className="primary-button" disabled={busy} onClick={() => void prepare("download")}>{preparing === "download" ? "Preparing Word file…" : source?.file ? "Download updated Word file" : "Download sample projection"}</button>
            </div>
          </div>
          <dl className="ledger-totals">
            <div><dt>Opening balance</dt><dd>{displayedMoney(projection.opening)}</dd></div>
            <div><dt>Total debits</dt><dd>{displayedMoney(projection.debitTotal)}</dd></div>
            <div><dt>Total credits</dt><dd>{displayedMoney(projection.creditTotal)}</dd></div>
            <div><dt>Closing balance</dt><dd>{displayedMoney(projection.closing)}</dd></div>
          </dl>
          {dateReview && preparedFile && <section className="document-preview-card" aria-label="Updated document preview">
            <div className="document-preview-header">
              <div><span className="preview-file-type">DOCX</span><div><strong>Updated Word preview</strong><small>Preview and download use the same prepared file.</small></div></div>
            </div>
            <div className="document-preview-body"><DocxDocumentPreview blob={preparedFile} /></div>
          </section>}
          {/* eslint-disable-next-line jsx-a11y/no-noninteractive-tabindex -- Allow keyboard users to scroll the transaction table. */}
          <div className="ledger-table-scroll" tabIndex={0} role="region" aria-label="Projected transactions">
            <table>
              <thead><tr><th scope="col">{dateReview ? "Date in upload" : "Date"}</th><th scope="col">Description</th><th scope="col">Debit</th><th scope="col">Credit</th><th scope="col">Running balance</th></tr></thead>
              <tbody>{ledger.hasOpeningEntry && <tr><td /><th scope="row">Opening balance</th><td>0.00</td><td>0.00</td><td>{formatMoney(projection.opening)}</td></tr>}{projection.transactions.map((row, index) => {
                const original = ledger.transactions[index];
                return <tr key={index} className={index === projection.transactions.length - 1 ? "ledger-final-row" : undefined}>
                  <td>{row.date}</td><td>{row.description || `Transaction ${index + 1}`}{index === 0 && !ledger.hasOpeningEntry && <small>Opening entry</small>}</td>
                  <td>{formatMoney(row.debit)}{row.debit !== original.debit && <small>Source {formatMoney(original.debit)}</small>}</td>
                  <td>{formatMoney(row.credit)}{row.credit !== original.credit && <small>Source {formatMoney(original.credit)}</small>}</td>
                  <td>{formatMoney(row.balance)}</td>
                </tr>;
              })}</tbody>
            </table>
          </div>
        </article>
      )}
    </section>
  );
}
