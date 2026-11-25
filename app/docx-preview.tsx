"use client";

import { useEffect, useRef, useState } from "react";

export function DocxDocumentPreview({ blob }: { blob: Blob }) {
  const containerRef = useRef<HTMLDivElement>(null);
  const [renderState, setRenderState] = useState<"rendering" | "ready" | "error">("rendering");

  useEffect(() => {
    const container = containerRef.current;
    if (!container) return;
    let cancelled = false;
    container.replaceChildren();
    setRenderState("rendering");

    void import("docx-preview")
      .then(async ({ renderAsync }) => {
        await renderAsync(blob, container, container, {
          breakPages: true,
          experimental: true,
          ignoreFonts: false,
          ignoreHeight: false,
          ignoreLastRenderedPageBreak: false,
          ignoreWidth: false,
          inWrapper: true,
          renderAltChunks: true,
          renderEndnotes: true,
          renderFooters: true,
          renderFootnotes: true,
          renderHeaders: true,
          useBase64URL: true,
        });
        if (!cancelled) setRenderState("ready");
      })
      .catch(() => {
        if (cancelled) return;
        container.replaceChildren();
        setRenderState("error");
      });

    return () => {
      cancelled = true;
      container.replaceChildren();
    };
  }, [blob]);

  return (
    <div className={`docx-preview-host ${renderState}`}>
      {renderState === "rendering" && <div className="preview-loading docx-render-loading"><i /><i /><i /><p>Rendering the prepared Word file…</p></div>}
      {renderState === "error" && <div className="preview-empty"><span>⌕</span><p>The Word preview could not be rendered, but the prepared download is still available.</p></div>}
      <div className="docx-render-surface" ref={containerRef} aria-label="Rendered updated Word document" />
    </div>
  );
}
