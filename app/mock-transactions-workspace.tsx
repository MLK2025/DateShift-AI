"use client";

import { useMemo, useState } from "react";
import { filterMockTransactions, generateMockTransactions, type MockPattern } from "./lib/mock-transactions";
import { SampleLedgerTemplate } from "./sample-ledger-template";

const amountFormat = new Intl.NumberFormat("en-GB", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const monthFormat = new Intl.DateTimeFormat("en-GB", { month: "long", year: "numeric", timeZone: "UTC" });

export function MockTransactionsWorkspace() {
  const [company, setCompany] = useState("");
  const [pattern, setPattern] = useState<MockPattern>("salary");
  const [batch, setBatch] = useState(1);
  const [search, setSearch] = useState("");
  const [month, setMonth] = useState("");
  const [type, setType] = useState<"all" | "credit" | "debit" | "payroll">("all");
  const rows = useMemo(() => generateMockTransactions(pattern, company, batch), [pattern, company, batch]);
  const months = [...new Set(rows.map(({ date }) => date.slice(0, 7)))];
  const filtered = filterMockTransactions(rows, { search, month, type });
  const hasFilters = Boolean(search || month || type !== "all");

  function clearFilters() {
    setSearch("");
    setMonth("");
    setType("all");
  }

  return (
    <section className="mock-workspace" aria-label="Mock transaction generator">
      <article className="panel mock-controls" id="how-it-works">
        <label className="mock-field">
          <span>Placeholder company name</span>
          <input type="text" value={company} maxLength={64} placeholder="Example Company" onChange={(event) => setCompany(event.target.value)} />
        </label>
        <fieldset className="mock-pattern">
          <legend>Sample pattern</legend>
          <div>
            <label><input type="radio" name="mockPattern" checked={pattern === "salary"} onChange={() => setPattern("salary")} /> Salary structured</label>
            <label><input type="radio" name="mockPattern" checked={pattern === "non-salary"} onChange={() => setPattern("non-salary")} /> Non-salary structured</label>
          </div>
        </fieldset>
        <button type="button" className="primary-button" onClick={() => setBatch((current) => current + 1)}>Regenerate samples</button>
      </article>

      <SampleLedgerTemplate company={company} pattern={pattern} batch={batch} />

      <article className="panel mock-table">
        <div className="mock-table-heading">
          <div><h2>Sample transactions</h2><p>Six completed months. Amounts use DEMO units.</p></div>
          <span>Synthetic data</span>
        </div>
        <div className="mock-filters">
          <label className="mock-field"><span>Search descriptions or references</span><input type="search" value={search} onChange={(event) => setSearch(event.target.value)} placeholder="Search samples" /></label>
          <label className="mock-field"><span>Month</span><select value={month} onChange={(event) => setMonth(event.target.value)}>
            <option value="">All months</option>
            {months.map((period) => <option key={period} value={period}>{monthFormat.format(new Date(`${period}-01T00:00:00Z`))}</option>)}
          </select></label>
          <label className="mock-field"><span>Transaction type</span><select value={type} onChange={(event) => setType(event.target.value as typeof type)}>
            <option value="all">All transactions</option>
            <option value="credit">Credits</option>
            <option value="debit">Debits</option>
            <option value="payroll">Payroll only</option>
          </select></label>
          {hasFilters && <button type="button" className="secondary-button" onClick={clearFilters}>Clear filters</button>}
        </div>
        <p className="mock-row-count" role="status">Showing {filtered.length} of {rows.length} sample transactions</p>
        {/* eslint-disable-next-line jsx-a11y/no-noninteractive-tabindex -- Allow keyboard scrolling of the table on small screens. */}
        <div className="ledger-table-scroll mock-table-scroll" role="region" aria-label="Synthetic transaction rows" tabIndex={0}>
          <table aria-label="Synthetic transactions">
            <thead><tr>
              <th scope="col">Date</th><th scope="col">Description</th><th scope="col">Mock reference</th><th scope="col">Category</th>
              <th className="mock-amount" scope="col">Debit (DEMO)</th><th className="mock-amount" scope="col">Credit (DEMO)</th>
            </tr></thead>
            <tbody>
              {filtered.length ? filtered.map((row) => <tr key={row.reference}>
                <td><time dateTime={row.date}>{row.date}</time></td>
                <td>{row.description}</td><td>{row.reference}</td><td>{row.category}</td>
                <td className="mock-amount">{row.flow === "debit" ? amountFormat.format(row.minorUnits / 100) : "—"}</td>
                <td className="mock-amount">{row.flow === "credit" ? amountFormat.format(row.minorUnits / 100) : "—"}</td>
              </tr>) : <tr><td colSpan={6}>No sample transactions match these filters.</td></tr>}
            </tbody>
          </table>
        </div>
      </article>
    </section>
  );
}
