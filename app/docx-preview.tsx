"use client";

import { useEffect, useRef, useState } from "react";
import type { PositionedPage } from "./lib/pdf-word-positioned";
import { fontFamilyCss, wordFontKey } from "./lib/word-fonts";

function PositionedWordPreview({ pages, families }: { pages: PositionedPage[]; families: Map<string, string> }) {
  return <div className="positioned-word-preview" aria-label="Converted Word pages">
    {pages.map((page, index) => <div key={index} className="positioned-word-page" style={{ width: `${page.width}pt`, height: `${page.height}pt` }}>
      {/* eslint-disable-next-line @next/next/no-img-element -- This browser-local image is embedded in the Word file. */}
      <img src={page.image} alt={page.text.length ? "" : `PDF page ${index + 1}`} />
      {page.text.map((run, runIndex) => <div key={runIndex} className="positioned-word-text" style={{
        left: `${run.x}pt`, top: `${run.y}pt`, width: `${run.width}pt`, height: `${run.height}pt`,
        fontSize: `${run.size}pt`, lineHeight: `${run.height}pt`, fontFamily: fontFamilyCss(families.get(wordFontKey(run)) ?? run.font),
        fontWeight: run.bold ? 700 : 400, fontStyle: run.italic ? "italic" : "normal", color: `#${run.color}`,
        direction: run.rtl ? "rtl" : "ltr", textAlign: run.rtl ? "right" : "left", transform: `rotate(${run.rotation}deg)`,
      }}>
        <span style={{ transform: `scaleX(${run.stretch / 100})`, transformOrigin: run.rtl ? "right" : "left" }}>{run.text}</span>
      </div>)}
    </div>)}
  </div>;
}

export function DocxDocumentPreview({ blob }: { blob: Blob }) {
  const containerRef = useRef<HTMLDivElement>(null);
  const [renderState, setRenderState] = useState<"rendering" | "ready" | "error">("rendering");
  const [positioned, setPositioned] = useState<{ pages: PositionedPage[]; families: Map<string, string> } | null>(null);

  useEffect(() => {
    const container = containerRef.current;
    if (!container) return;
    let cancelled = false;
    const previewFonts: FontFace[] = [];
    container.replaceChildren();
    setRenderState("rendering");
    setPositioned(null);

    void import("./lib/positioned-word-preview")
      .then(async ({ readPositionedWordPreview }) => {
        const preview = await readPositionedWordPreview(blob);
        if (cancelled) return;
        if (preview) {
          const families = new Map<string, string>();
          if (typeof FontFace === "function" && document.fonts) {
            await Promise.all(preview.fonts.map(async (font) => {
              try {
                const face = await new FontFace(`dateshift-preview-${crypto.randomUUID()}`, font.bytes.slice().buffer, {
                  weight: font.bold ? "700" : "400",
                  style: font.italic ? "italic" : "normal",
                }).load();
                if (cancelled) return;
                document.fonts.add(face);
                previewFonts.push(face);
                families.set(wordFontKey(font), face.family);
              } catch {
                // The preview can still use a system font if this face cannot load.
              }
            }));
          }
          if (cancelled) return;
          setPositioned({ pages: preview.pages, families });
          setRenderState("ready");
          return;
        }
        const { renderAsync } = await import("docx-preview");
        if (cancelled) return;
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
      for (const face of previewFonts) document.fonts.delete(face);
      container.replaceChildren();
    };
  }, [blob]);

  return (
    <div className={`docx-preview-host ${renderState}`}>
      {renderState === "rendering" && <div className="preview-loading docx-render-loading"><i /><i /><i /><p>Rendering the prepared Word file…</p></div>}
      {renderState === "error" && <div className="preview-empty"><span>⌕</span><p>The Word preview could not be rendered, but the prepared download is still available.</p></div>}
      {positioned && <PositionedWordPreview {...positioned} />}
      <div className="docx-render-surface" ref={containerRef} hidden={Boolean(positioned)} aria-label="Rendered updated Word document" />
    </div>
  );
}
