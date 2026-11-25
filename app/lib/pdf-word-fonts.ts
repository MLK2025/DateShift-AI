import { PDFDict, PDFDocument, PDFName, PDFRawStream, decodePDFRawStream } from "pdf-lib";
import type { PDFDocumentProxy, PDFPageProxy } from "pdfjs-dist";
import { fontFamilyCss, wordFontKey, type WordFont } from "./word-fonts";

type EmbeddedFont = { bytes: Uint8Array; family: string; covers: (codes: Set<number>) => boolean };
type FontUse = { font: WordFont; candidates: Set<EmbeddedFont>; codes: Set<number> };
type PdfFont = {
  name?: string;
  bold?: boolean;
  italic?: boolean;
  isMonospace?: boolean;
  isSerifFont?: boolean;
  data?: Uint8Array;
  widths?: Record<string, number>;
};

const maxFontBytes = 8 * 1024 * 1024;
const subset = /^[A-Z]{6}\+/;
const styleSuffix = /(?:[-_, ]?(?:(?:semi|demi|extra)[- ]?bold|bold|regular|italic|oblique|slanted))+$/i;

function table(data: DataView, name: string) {
  if (data.byteLength < 12) return null;
  for (let offset = 12, i = 0; i < data.getUint16(4) && offset + 16 <= data.byteLength; i++, offset += 16) {
    if (String.fromCharCode(...new Uint8Array(data.buffer, data.byteOffset + offset, 4)) !== name) continue;
    const start = data.getUint32(offset + 8);
    const length = data.getUint32(offset + 12);
    if (start + length <= data.byteLength) return new DataView(data.buffer, data.byteOffset + start, length);
  }
  return null;
}

function checksum(data: DataView) {
  let sum = 0;
  for (let offset = 0; offset < data.byteLength; offset += 4) {
    let word = 0;
    for (let i = 0; i < 4 && offset + i < data.byteLength; i++) {
      word |= data.getUint8(offset + i) << (24 - i * 8);
    }
    sum = (sum + word) >>> 0;
  }
  return sum;
}

function fixFixedPitchMetrics(bytes: Uint8Array, info: PdfFont) {
  if (!info.isMonospace || !info.widths) return bytes;
  const widths = Object.values(info.widths);
  const width = widths[0];
  if (!width || !widths.every((value) => Math.abs(value - width) < .01)) return bytes;
  const copy = bytes.slice();
  const view = new DataView(copy.buffer);
  const head = table(view, "head");
  const hhea = table(view, "hhea");
  const hmtx = table(view, "hmtx");
  if (!head || head.byteLength < 20 || !hhea || hhea.byteLength < 36 || !hmtx) return bytes;
  const advance = Math.round(width * head.getUint16(18) / 1000);
  const count = hhea.getUint16(34);
  if (!advance || advance > 65535 || !count || count * 4 > hmtx.byteLength) return bytes;
  // Some subset writers put the wrong advance in the last long metric.
  // All subsequent glyphs inherit that value, including visible characters.
  if (Array.from({ length: count }, (_, index) => hmtx.getUint16(index * 4)).every((value) => value === advance)) return bytes;
  for (let i = 0; i < count; i++) hmtx.setUint16(i * 4, advance);
  for (let i = 0, offset = 12; i < view.getUint16(4) && offset + 16 <= view.byteLength; i++, offset += 16) {
    if (view.getUint32(offset + 8) === hmtx.byteOffset) {
      view.setUint32(offset + 4, checksum(hmtx));
      break;
    }
  }
  head.setUint32(8, 0);
  head.setUint32(8, 0xb1b0afba - checksum(view));
  return copy;
}

function unicodeMaps(cmap: DataView) {
  const maps: ((code: number) => boolean)[] = [];
  for (let i = 0, record = 4; i < cmap.getUint16(2) && record + 8 <= cmap.byteLength; i++, record += 8) {
    const platform = cmap.getUint16(record);
    const encoding = cmap.getUint16(record + 2);
    const offset = cmap.getUint32(record + 4);
    if (!(platform === 0 || (platform === 3 && (encoding === 1 || encoding === 10))) || offset + 6 > cmap.byteLength) continue;
    const format = cmap.getUint16(offset);
    if (format === 4 && offset + 14 <= cmap.byteLength) {
      const end = offset + cmap.getUint16(offset + 2);
      const count = cmap.getUint16(offset + 6) / 2;
      const ends = offset + 14;
      const starts = ends + count * 2 + 2;
      const deltas = starts + count * 2;
      const ranges = deltas + count * 2;
      if (end > cmap.byteLength || ranges + count * 2 > end || !Number.isInteger(count)) continue;
      maps.push((code) => {
        if (code > 0xffff) return false;
        for (let segment = 0; segment < count * 2; segment += 2) {
          if (code > cmap.getUint16(ends + segment)) continue;
          const start = cmap.getUint16(starts + segment);
          if (code < start) return false;
          const delta = cmap.getInt16(deltas + segment);
          const range = cmap.getUint16(ranges + segment);
          if (!range) return ((code + delta) & 0xffff) !== 0;
          const address = ranges + segment + range + 2 * (code - start);
          const glyph = address + 2 <= end ? cmap.getUint16(address) : 0;
          return glyph !== 0 && ((glyph + delta) & 0xffff) !== 0;
        }
        return false;
      });
    } else if (format === 12 && offset + 16 <= cmap.byteLength) {
      const groups = cmap.getUint32(offset + 12);
      const end = offset + cmap.getUint32(offset + 4);
      if (end > cmap.byteLength || offset + 16 + groups * 12 > end) continue;
      maps.push((code) => {
        let low = 0;
        let high = groups - 1;
        while (low <= high) {
          const index = (low + high) >>> 1;
          const group = offset + 16 + index * 12;
          const start = cmap.getUint32(group);
          if (code < start) high = index - 1;
          else if (code > cmap.getUint32(group + 4)) low = index + 1;
          else return cmap.getUint32(group + 8) + code - start !== 0;
        }
        return false;
      });
    }
  }
  return maps;
}

function readEmbeddedFont(bytes: Uint8Array) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (view.byteLength < 12 || view.getUint32(0) !== 0x00010000 || !table(view, "glyf")) return null;
  const os2 = table(view, "OS/2");
  const cmap = table(view, "cmap");
  const name = table(view, "name");
  let family = "";
  // The legacy family groups the four Word styles. Keep weights like Light
  // in the family name when the typeface uses a separate Word family for them.
  if (name && name.byteLength >= 6) {
    let best = 0;
    for (let i = 0, record = 6; i < name.getUint16(2) && record + 12 <= name.byteLength; i++, record += 12) {
      const platform = name.getUint16(record);
      const id = name.getUint16(record + 6);
      const score = (id === 1 ? 4 : id === 16 ? 2 : 0) + Number(name.getUint16(record + 4) === 0x0409);
      if ((id !== 1 && id !== 16) || (platform !== 0 && platform !== 3) || score <= best) continue;
      const start = name.getUint16(4) + name.getUint16(record + 10);
      const length = name.getUint16(record + 8);
      if (start + length > name.byteLength || length % 2) continue;
      const value = new TextDecoder("utf-16be").decode(new Uint8Array(name.buffer, name.byteOffset + start, length));
      if (!value.trim()) continue;
      family = cleanName(value);
      best = score;
    }
  }
  const rights = os2 && os2.byteLength >= 10 ? os2.getUint16(8) : -1;
  const maps = cmap && cmap.byteLength >= 4 ? unicodeMaps(cmap) : [];
  const editable = rights === 0 || rights === 8;
  return {
    family,
    editable,
    embedded: editable && maps.length
      ? { bytes, family, covers: (codes: Set<number>) => maps.some((has) => Array.from(codes).every(has)) }
      : undefined,
  };
}

function cleanName(name: string) {
  return name.replace(subset, "").replace(/[\p{Cc}\uFFFE\uFFFF]/gu, "").trim().slice(0, 128);
}

function fontName(name: string, fallback: string) {
  if (!name) return fallback;
  const stem = name.replace(/(?:PS)?MT$/i, "").replace(styleSuffix, "").trim();
  const aliases: Record<string, string> = {
    timesroman: "Times",
    timesnewroman: "Times New Roman",
    couriernew: "Courier New",
    helveticaneue: "Helvetica Neue",
    segoeui: "Segoe UI",
  };
  const compact = stem.replace(/[\s_-]/g, "").toLowerCase();
  return Object.hasOwn(aliases, compact) ? aliases[compact] : (stem.replace(/([a-z0-9])([A-Z])/g, "$1 $2").trim() || name);
}

export class PdfWordFonts {
  private source = new Map<string, NonNullable<ReturnType<typeof readEmbeddedFont>>>();
  private used = new Map<string, FontUse>();
  private normalized = new Map<string, EmbeddedFont | undefined>();
  private measuring = new Map<EmbeddedFont, string>();
  private faces: FontFace[] = [];

  async readSource(pdf: PDFDocumentProxy, signal?: AbortSignal) {
    try {
      const document = await PDFDocument.load(await pdf.getData(), { updateMetadata: false, throwOnInvalidObject: false });
      signal?.throwIfAborted();
      for (const [, object] of document.context.enumerateIndirectObjects()) {
        if (!(object instanceof PDFDict)) continue;
        try {
          const descriptor = object.lookupMaybe(PDFName.of("FontDescriptor"), PDFDict) ?? object;
          const name = descriptor.lookupMaybe(PDFName.of("FontName"), PDFName)?.decodeText();
          const stream = descriptor.lookup(PDFName.of("FontFile2"));
          if (!name || this.source.has(name) || !(stream instanceof PDFRawStream) || stream.getContentsSize() > maxFontBytes) continue;
          const length = stream.dict.get(PDFName.of("Length1"))?.toString();
          if (!length || !Number.isFinite(Number(length)) || Number(length) > maxFontBytes) continue;
          const bytes = decodePDFRawStream(stream).decode();
          const font = bytes.length <= maxFontBytes ? readEmbeddedFont(bytes) : null;
          if (font) {
            this.source.set(name, font);
            if (!this.source.has(cleanName(name))) this.source.set(cleanName(name), font);
          }
        } catch {
          // A malformed or unsupported font must not block the conversion.
        }
      }
    } catch {
      signal?.throwIfAborted();
      // PDF.js can still recover editable text from some PDFs with broken xrefs.
    }
  }

  resolve(page: PDFPageProxy, id: string, generic: string, text: string): WordFont {
    const info = page.commonObjs.has(id) ? page.commonObjs.get(id) as PdfFont | undefined : undefined;
    const name = cleanName(info?.name ?? "");
    const source = this.source.get(info?.name ?? "") ?? this.source.get(name);
    const fallback = info?.isMonospace || /mono/i.test(generic) ? "Courier New"
      : info?.isSerifFont || generic === "serif" ? "Times New Roman" : "Arial";
    const font = {
      font: source?.family || fontName(name, fallback),
      bold: Boolean(info?.bold || /(?:bold|demi)(?:italic|oblique)?(?:PS)?(?:MT)?$/i.test(name)),
      italic: Boolean(info?.italic || /italic|oblique|slanted/i.test(name)),
    };
    const key = wordFontKey(font);
    const use = this.used.get(key) ?? { font, candidates: new Set<EmbeddedFont>(), codes: new Set<number>() };
    // Some subsets represent spaces by positioning only, with no glyph.
    for (const char of text) {
      if (!/[\s\p{Cf}]/u.test(char)) use.codes.add(char.codePointAt(0)!);
    }
    if (source?.embedded) use.candidates.add(source.embedded);
    if (source?.editable && info?.data) {
      // PDF.js can add Unicode mappings to subset fonts that were encoded
      // only for PDF rendering. The original font's embedding rights govern both.
      if (!this.normalized.has(id)) {
        const bytes = fixFixedPitchMetrics(info.data, info);
        const font = bytes.byteLength <= maxFontBytes ? readEmbeddedFont(bytes)?.embedded : undefined;
        this.normalized.set(id, font);
      }
      const candidate = this.normalized.get(id);
      if (candidate) use.candidates.add(candidate);
    }
    this.used.set(key, use);
    return font;
  }

  embedded(font: WordFont) {
    const use = this.used.get(wordFontKey(font));
    return use && Array.from(use.candidates).find((candidate) => candidate.covers(use.codes));
  }

  *embeddingEntries() {
    for (const { font } of this.used.values()) {
      const candidate = this.embedded(font);
      if (candidate) yield { font, bytes: candidate.bytes };
    }
  }

  async measure(context: CanvasRenderingContext2D, font: WordFont, size: number, value: string) {
    const candidate = this.embedded(font);
    let family = candidate && this.measuring.get(candidate);
    if (candidate && !family && typeof FontFace === "function" && document.fonts) {
      try {
        const face = await new FontFace(`dateshift-pdf-${crypto.randomUUID()}`, candidate.bytes.slice().buffer).load();
        document.fonts.add(face);
        this.faces.push(face);
        family = face.family;
        this.measuring.set(candidate, family);
      } catch {
        this.measuring.set(candidate, font.font);
      }
    }
    const style = family && family !== font.font ? "" : `${font.italic ? "italic " : ""}${font.bold ? "bold " : ""}`;
    context.font = `${style}${size}px ${fontFamilyCss(family || font.font)}`;
    return context.measureText(value).width;
  }

  dispose() {
    for (const face of this.faces) document.fonts.delete(face);
  }
}
