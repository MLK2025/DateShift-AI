import { useMemo, useState } from "react";
import { formatMoney, parseMoney, projectLedger, reconcileLedger, type Ledger, type LedgerMethod } from "./lib/ledger";

const sampleLedger: Ledger = {
  currency: "GHS",
  opening: BigInt(100_000),
  reportedClosing: BigInt(125_000),
  transactions: [
    { date: "01-Dec-2025", description: "Deposit", debit: BigInt(0), credit: BigInt(40_000), balance: BigInt(140_000) },
    { date: "02-Dec-2025", description: "Supplies", debit: BigInt(15_000), credit: BigInt(0), balance: BigInt(125_000) },
  ],
};

export type LedgerSource = { name: string; ledger: Ledger; target?: string };

export function LedgerWorkspace({ initialSource = null }: { initialSource?: LedgerSource | null }) {
  const [source, setSource] = useState<LedgerSource | null>(initialSource);
  const [openingMode, setOpeningMode] = useState<"calculated" | "fixed">("calculated");
  const [opening, setOpening] = useState(() => initialSource ? formatMoney(initialSource.ledger.opening) : "");
  const [closing, setClosing] = useState(() => initialSource?.target ?? (initialSource ? formatMoney(reconcileLedger(initialSource.ledger).closing) : ""));
  const [method, setMethod] = useState<LedgerMethod>("spread");
  const [reading, setReading] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [error, setError] = useState("");
  const ledger = source?.ledger;
  const reconciliation = useMemo(() => ledger ? reconcileLedger(ledger) : null, [ledger]);
  const target = parseMoney(closing, ledger?.currency);
  const specifiedOpening = openingMode === "fixed" ? parseMoney(opening, ledger?.currency) : null;
  const canCalculate = target !== null && (openingMode !== "fixed" || specifiedOpening !== null);
  const projection = useMemo(() => ledger && canCalculate && target !== null
    ? projectLedger(ledger, target, specifiedOpening, method)
    : null, [ledger, canCalculate, target, specifiedOpening, method]);

  function loadSource(name: string, next: Ledger) {
    setSource({ name, ledger: next });
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
      loadSource(file.name, await readLedger(file));
    } catch (failure) {
      setError(failure instanceof Error ? failure.message : "The transaction table could not be read.");
    } finally {
      setReading(false);
    }
  }

  async function download() {
    if (!ledger || !projection) return;
    setExporting(true);
    setError("");
    try {
      const { createLedgerReport } = await import("./lib/ledger-docx");
      const url = URL.createObjectURL(await createLedgerReport(ledger, projection));
      const link = document.createElement("a");
      link.href = url;
      link.download = `${source.name.replace(/\.docx$/i, "")}-projection.docx`;
      document.body.append(link);
      link.click();
      link.remove();
      window.setTimeout(() => URL.revokeObjectURL(url), 60_000);
    } catch {
      setError("The Word report could not be created. Try downloading it again.");
    } finally {
      setExporting(false);
    }
  }

  const busy = reading || exporting;
  const currency = ledger?.currency;
  const reportedClosing = ledger?.reportedClosing;
  const displayedMoney = (value: bigint) => `${currency ? `${currency} ` : ""}${formatMoney(value)}`;

  return (
    <section className="ledger-workspace" aria-label="Transaction projection workspace">
      <div className="ledger-settings">
        <article className="panel ledger-panel">
          <h2>Choose a Word ledger</h2>
          <p>Use a .docx with Date, Debit, Credit, and Balance columns. Table continuations are read in document order.</p>
          <input aria-label="Choose a Word ledger" type="file" accept=".docx" disabled={busy} onChange={(event) => {
            const file = event.currentTarget.files?.[0];
            event.currentTarget.value = "";
            if (file) void loadFile(file);
          }} />
          <button className="secondary-button" disabled={busy} onClick={() => loadSource("Example ledger", sampleLedger)}>Use sample ledger</button>
          {reading && <p role="status">Reading the transaction table…</p>}
          {source && <p role="status"><strong>{source.name}</strong> · {source.ledger.transactions.length} transactions{currency ? ` · ${currency}` : ""}</p>}
        </article>
        <article className="panel ledger-panel">
          <h2>Set the projected balances</h2>
          <div className="figure-controls ledger-controls">
            <label><span>Target closing balance {currency && `(${currency})`}</span><input value={closing} onChange={(event) => setClosing(event.target.value)} inputMode="decimal" placeholder="50,000.00" disabled={busy} /></label>
            <label><span>Opening balance</span><select value={openingMode} onChange={(event) => setOpeningMode(event.target.value === "fixed" ? "fixed" : "calculated")} disabled={busy}><option value="calculated">Calculate from the target and existing transactions</option><option value="fixed">Set the opening balance and adjust transactions</option></select></label>
            {openingMode === "fixed" && <>
              <label><span>Set opening balance {currency && `(${currency})`}</span><input value={opening} onChange={(event) => setOpening(event.target.value)} inputMode="decimal" placeholder="10,000.00" disabled={busy} /></label>
              <label><span>Apply the adjustment</span><select value={method} onChange={(event) => setMethod(event.target.value === "last" ? "last" : "spread")} disabled={busy}><option value="spread">Spread across transactions proportionally</option><option value="last">Apply to the final transaction</option></select></label>
            </>}
            {ledger && !canCalculate && <p role="status">Enter a valid amount with up to two decimal places in {currency ? `${currency} or without a currency code` : "the balance fields"}.</p>}
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
            <div><h2>Transaction projection</h2><p>Each running balance adds the row’s credit and subtracts its debit. The final row equals the target closing balance. Parentheses indicate a negative balance.</p></div>
            <button className="primary-button" disabled={busy} onClick={() => void download()}>{exporting ? "Preparing report…" : "Download Word report"}</button>
          </div>
          <dl className="ledger-totals">
            <div><dt>Opening balance</dt><dd>{displayedMoney(projection.opening)}</dd></div>
            <div><dt>Total debits</dt><dd>{displayedMoney(projection.debitTotal)}</dd></div>
            <div><dt>Total credits</dt><dd>{displayedMoney(projection.creditTotal)}</dd></div>
            <div><dt>Closing balance</dt><dd>{displayedMoney(projection.closing)}</dd></div>
          </dl>
          {/* eslint-disable-next-line jsx-a11y/no-noninteractive-tabindex -- Allow keyboard users to scroll the transaction table. */}
          <div className="ledger-table-scroll" tabIndex={0} role="region" aria-label="Projected transactions">
            <table>
              <thead><tr><th scope="col">Date</th><th scope="col">Description</th><th scope="col">Debit</th><th scope="col">Credit</th><th scope="col">Running balance</th></tr></thead>
              <tbody>{projection.transactions.map((row, index) => {
                const original = ledger.transactions[index];
                return <tr key={index} className={index === projection.transactions.length - 1 ? "ledger-final-row" : undefined}>
                  <td>{row.date}</td><td>{row.description || `Transaction ${index + 1}`}</td>
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
