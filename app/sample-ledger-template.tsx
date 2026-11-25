"use client";

import { useState } from "react";
import { DocxDocumentPreview } from "./docx-preview";
import type { createSampleLedgerDocument } from "./lib/mock-ledger-docx";
import type { MockPattern } from "./lib/mock-transactions";

type SampleSettings = { company: string; pattern: MockPattern; batch: number };
type PreparedSample = Awaited<ReturnType<typeof createSampleLedgerDocument>> & SampleSettings & { source: File };

export function SampleLedgerTemplate({ company, pattern, batch }: SampleSettings) {
  const [file, setFile] = useState<File | null>(null);
  const [prepared, setPrepared] = useState<PreparedSample | null>(null);
  const [preparing, setPreparing] = useState(false);
  const [error, setError] = useState("");
  const sample = prepared?.source === file && prepared.company === company
    && prepared.pattern === pattern && prepared.batch === batch ? prepared : null;

  function choose(nextFile?: File) {
    setPrepared(null);
    setError("");
    setFile(null);
    if (!nextFile) return;
    if (!/\.docx$/i.test(nextFile.name) || nextFile.size > 25 * 1024 * 1024) {
      setError("Choose a Word .docx template up to 25 MB.");
      return;
    }
    setFile(nextFile);
  }

  async function prepare() {
    if (!file || preparing) return;
    setPreparing(true);
    setPrepared(null);
    setError("");
    try {
      const { createSampleLedgerDocument } = await import("./lib/mock-ledger-docx");
      const result = await createSampleLedgerDocument(file, company, pattern, batch);
      setPrepared({ ...result, source: file, company, pattern, batch });
    } catch (failure) {
      setError(failure instanceof Error ? failure.message : "The Word template could not be processed.");
    } finally {
      setPreparing(false);
    }
  }

  function download() {
    if (!sample) return;
    const url = URL.createObjectURL(sample.blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = sample.fileName;
    document.body.appendChild(link);
    link.click();
    link.remove();
    window.setTimeout(() => URL.revokeObjectURL(url), 60_000);
  }

  return (
    <section className="sample-template" aria-label="Word ledger template">
      <article className="panel mock-table sample-template-controls">
        <div className="mock-table-heading">
          <div><h2>Word ledger template</h2><p>Use the company name and sample pattern above to fill an uploaded .docx ledger. The exported copy includes a SAMPLE footer.</p></div>
        </div>
        <div className="sample-template-fields">
          <label className="mock-field"><span>Word template</span><input type="file" accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document" disabled={preparing} onChange={(event) => {
            choose(event.currentTarget.files?.[0]);
          }} /></label>
          <button type="button" className="primary-button" disabled={!file || preparing} onClick={() => void prepare()}>{preparing ? "Preparing sample…" : "Prepare sample DOCX"}</button>
        </div>
        {preparing && <p className="mock-row-count" role="status">Generating transactions, reconciling balances, and adding the SAMPLE footer.</p>}
        {error && <p className="notice" role="alert">{error}</p>}
      </article>
      {sample && <section className="document-preview-card conversion-preview" aria-label="Sample Word document preview">
        <div className="document-preview-header">
          <div><span className="preview-file-type">DOCX</span><div><strong>{sample.fileName}</strong><small role="status">{sample.rowCount} sample rows · {sample.payrollCount} payroll {sample.payrollCount === 1 ? "credit" : "credits"}</small></div></div>
          <button type="button" className="primary-button" onClick={download}>Download sample DOCX</button>
        </div>
        <div className="document-preview-body"><DocxDocumentPreview blob={sample.blob} /></div>
      </section>}
    </section>
  );
}
