"use client";

import { useEffect, useRef, useState } from "react";
import { DocxDocumentPreview } from "./docx-preview";
import type { ConvertedWord, WordConversionMode } from "./lib/pdf-to-word";
export function PdfWordWorkspace({ initialFile = null }: { initialFile?: File | null }) {
  const inputRef = useRef<HTMLInputElement>(null);
  const conversion = useRef<AbortController | null>(null);
  const [file, setFile] = useState(initialFile);
  const [mode, setMode] = useState<WordConversionMode>("editable");
  const [result, setResult] = useState<ConvertedWord | null>(null);
  const [progress, setProgress] = useState("");
  const [error, setError] = useState("");
  const busy = Boolean(progress);

  useEffect(() => () => conversion.current?.abort(), []);

  function chooseFile(nextFile: File) {
    if (busy) return;
    setResult(null);
    if (!/\.pdf$/i.test(nextFile.name) || nextFile.size > 25 * 1024 * 1024) {
      setFile(null);
      setError("Choose a PDF file up to 25 MB.");
      return;
    }
    setFile(nextFile);
    setError("");
  }

  async function convert() {
    if (!file || conversion.current) return;
    const controller = new AbortController();
    conversion.current = controller;
    setResult(null);
    setError("");
    setProgress("Opening the PDF…");
    try {
      const { convertPdfToWord } = await import("./lib/pdf-to-word");
      const output = await convertPdfToWord(file, mode, (done, total) => {
        setProgress(done === total ? "Building the Word file…" : `Converted ${done} of ${total} PDF pages…`);
      }, controller.signal);
      if (!controller.signal.aborted) setResult(output);
    } catch (failure) {
      if (!controller.signal.aborted) {
        setError(failure instanceof Error && failure.name === "PasswordException"
          ? "This PDF is password-protected. Choose an unlocked copy."
          : "This PDF could not be converted. Try a different PDF or change the conversion format.");
      }
    } finally {
      conversion.current = null;
      if (!controller.signal.aborted) setProgress("");
    }
  }

  function download() {
    if (!result) return;
    const url = URL.createObjectURL(result.blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = result.fileName;
    document.body.appendChild(link);
    link.click();
    link.remove();
    window.setTimeout(() => URL.revokeObjectURL(url), 60_000);
  }

  function chooseMode(nextMode: WordConversionMode) {
    setMode(nextMode);
    setResult(null);
    setError("");
  }

  return (
    <section className="pdf-word-workspace" aria-label="PDF to Word conversion">
      <div className="workspace-grid">
        <article className="panel pdf-word-panel">
          <h2>Choose a PDF</h2>
          <p>PDF files up to 25 MB</p>
          <input ref={inputRef} type="file" accept=".pdf,application/pdf" hidden onChange={(event) => {
            const selected = event.currentTarget.files?.[0];
            event.currentTarget.value = "";
            if (selected) chooseFile(selected);
          }} />
          <button type="button" className={`dropzone ${file ? "has-file" : ""}`} disabled={busy} onClick={() => inputRef.current?.click()}
            onDragOver={(event) => event.preventDefault()}
            onDrop={(event) => {
              event.preventDefault();
              const dropped = event.dataTransfer.files[0];
              if (dropped) chooseFile(dropped);
            }}>
            <span className="file-icon" aria-hidden="true"><i>{file ? "✓" : "+"}</i></span>
            <strong>{file ? file.name : "Drop a PDF here"}</strong>
            <small>{file ? "Choose a different file" : "or browse your files"}</small>
          </button>
        </article>
        <article className="panel pdf-word-panel" id="how-it-works">
          <fieldset className="conversion-formats" disabled={busy}>
            <legend>Choose a Word format</legend>
            <label>
              <input type="radio" name="wordFormat" checked={mode === "editable"} onChange={() => chooseMode("editable")} />
              <strong>Editable text with original layout</strong>
              <small>Keep the page artwork and edit selectable text in positioned Word text boxes. Scanned pages are kept as images.</small>
            </label>
            <label>
              <input type="radio" name="wordFormat" checked={mode === "reflow"} onChange={() => chooseMode("reflow")} />
              <strong>Editable paragraphs and tables</strong>
              <small>Rebuild text and simple tables as regular Word content. This can change the layout.</small>
            </label>
            <label>
              <input type="radio" name="wordFormat" checked={mode === "layout"} onChange={() => chooseMode("layout")} />
              <strong>Page images</strong>
              <small>Include each page as an image. Works for scans and keeps the page appearance.</small>
            </label>
          </fieldset>
          <button type="button" className="primary-button" disabled={!file || busy} onClick={() => void convert()}>{busy ? "Converting…" : "Convert to Word"}</button>
          {progress && <p role="status">{progress}</p>}
          {error && <p className="notice" role="alert">{error}</p>}
        </article>
      </div>
      {result && <section className="document-preview-card conversion-preview" aria-label="Converted Word document preview">
        <div className="document-preview-header">
          <div><span className="preview-file-type">DOCX</span><div><strong>{result.fileName}</strong><small role="status">{result.pageCount} PDF {result.pageCount === 1 ? "page" : "pages"} converted{result.imagePages ? ` · ${result.imagePages} as ${result.imagePages === 1 ? "an image" : "images"}` : ""}</small></div></div>
          <button type="button" className="primary-button" onClick={download}>Download Word file</button>
        </div>
        <div className="document-preview-body"><DocxDocumentPreview blob={result.blob} /></div>
      </section>}
    </section>
  );
}
