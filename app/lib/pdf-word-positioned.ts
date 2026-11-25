import { OPS, type PDFPageProxy } from "pdfjs-dist";
import type { PageViewport } from "pdfjs-dist";

export type PositionedText = {
  text: string;
  x: number;
  y: number;
  width: number;
  height: number;
  size: number;
  rotation: number;
  stretch: number;
  font: string;
  bold: boolean;
  italic: boolean;
  rtl: boolean;
  color: string;
};

export type PositionedPage = {
  width: number;
  height: number;
  image: string;
  text: PositionedText[];
};

export function imageViewport(page: PDFPageProxy) {
  const { width, height } = page.getViewport({ scale: 1 });
  return page.getViewport({ scale: Math.min(2, 2000 / Math.max(width, height)) });
}

export async function renderPdfCanvas(page: PDFPageProxy, viewport: PageViewport, signal?: AbortSignal, hideText = false) {
  signal?.throwIfAborted();
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.ceil(viewport.width));
  canvas.height = Math.max(1, Math.ceil(viewport.height));
  const context = canvas.getContext("2d", { willReadFrequently: hideText });
  if (!context) throw new Error("The browser could not render this PDF page.");
  const render = page.render({
    canvas,
    viewport,
    background: "#ffffff",
    operationsFilter: hideText ? (index, operations) => {
      const opcode = operations.fnArray[index];
      return opcode !== OPS.showText && opcode !== OPS.showSpacedText;
    } : undefined,
  });
  const cancel = () => render.cancel();
  signal?.addEventListener("abort", cancel, { once: true });
  try {
    await render.promise;
    signal?.throwIfAborted();
    return context;
  } catch (failure) {
    canvas.width = 0;
    canvas.height = 0;
    throw failure;
  } finally {
    signal?.removeEventListener("abort", cancel);
  }
}

function colorHex(value: number) {
  return value.toString(16).padStart(2, "0").toUpperCase();
}

// Replace only the pixels under a text run. Text without a usable Unicode
// mapping, and graphics drawn with text clips, can remain in the artwork.
function separateTextPixels(
  original: ImageData,
  background: ImageData,
  output: ImageData,
  scale: number,
  x: number,
  y: number,
  width: number,
  size: number,
  angle: number,
) {
  const cos = Math.cos(angle);
  const sin = Math.sin(angle);
  const left = -.35;
  const right = width + .7;
  const top = -size * 1.06;
  const bottom = size * .37;
  const corners = [[left, top], [right, top], [left, bottom], [right, bottom]];
  const xs = corners.map(([u, v]) => (x + u * cos - v * sin) * scale);
  const ys = corners.map(([u, v]) => (y + u * sin + v * cos) * scale);
  const minX = Math.max(0, Math.floor(Math.min(...xs)));
  const minY = Math.max(0, Math.floor(Math.min(...ys)));
  const maxX = Math.min(original.width, Math.ceil(Math.max(...xs)));
  const maxY = Math.min(original.height, Math.ceil(Math.max(...ys)));
  let strongest = 0;
  let sample = 0;
  for (let row = minY; row < maxY; row += 1) {
    for (let col = minX; col < maxX; col += 1) {
      if (Math.abs(sin) > .001 || cos < 0) {
        const dx = (col + .5) / scale - x;
        const dy = (row + .5) / scale - y;
        const u = dx * cos + dy * sin;
        const v = dy * cos - dx * sin;
        if (u < left || u > right || v < top || v > bottom) continue;
      }
      const i = (row * original.width + col) * 4;
      const difference = Math.abs(original.data[i] - background.data[i])
        + Math.abs(original.data[i + 1] - background.data[i + 1])
        + Math.abs(original.data[i + 2] - background.data[i + 2]);
      if (difference > strongest) {
        strongest = difference;
        sample = i;
      }
      if (difference) {
        output.data[i] = background.data[i];
        output.data[i + 1] = background.data[i + 1];
        output.data[i + 2] = background.data[i + 2];
      }
    }
  }
  return strongest === 0 ? null : colorHex(original.data[sample]) + colorHex(original.data[sample + 1]) + colorHex(original.data[sample + 2]);
}

function wordFont(name: string, family: string) {
  const compact = name.replace(/^[A-Z]{6}\+/, "").toLowerCase().replace(/[^a-z]/g, "");
  const supported = ["Arial", "Aptos", "Calibri", "Cambria", "Courier New", "Georgia", "Segoe UI", "Tahoma", "Times New Roman", "Verdana"];
  const match = supported.find((font) => compact.startsWith(font.toLowerCase().replaceAll(" ", "")));
  if (match) return match;
  if (/courier|mono/i.test(name + family)) return "Courier New";
  if (/^times/i.test(compact) || (family.includes("serif") && !family.includes("sans"))) return "Times New Roman";
  return "Arial";
}

export async function positionPdfPage(page: PDFPageProxy, width: number, height: number, signal?: AbortSignal): Promise<PositionedPage> {
  const viewport = page.getViewport({ scale: 1 });
  const rasterViewport = imageViewport(page);
  const content = await page.getTextContent();
  const original = await renderPdfCanvas(page, rasterViewport, signal);
  let withoutText: CanvasRenderingContext2D | undefined;
  const text: PositionedText[] = [];
  try {
    const items = content.items.filter((item) => "str" in item && item.str.trim() && !item.str.includes("\uFFFD"));
    if (items.length) {
      withoutText = await renderPdfCanvas(page, rasterViewport, signal, true);
      const { width: pixelWidth, height: pixelHeight } = original.canvas;
      const painted = original.getImageData(0, 0, pixelWidth, pixelHeight);
      const background = withoutText.getImageData(0, 0, pixelWidth, pixelHeight);
      const artwork = original.createImageData(pixelWidth, pixelHeight);
      artwork.data.set(painted.data);
      const xScale = width / viewport.width;
      const yScale = height / viewport.height;
      for (const item of items) {
        if (!("str" in item)) continue;
        const [a, b, c, d, e, f] = item.transform;
        const [x, y] = viewport.convertToViewportPoint(e, f);
        const [advanceX, advanceY] = viewport.convertToViewportPoint(e + a, f + b);
        const size = Math.hypot(c, d) || item.height;
        if (size <= 0 || item.width <= 0 || item.dir === "ttb") continue;
        const originalAngle = Math.atan2(advanceY - y, advanceX - x);
        const quarterTurns = Math.round(originalAngle / (Math.PI / 2));
        const angle = quarterTurns * Math.PI / 2;
        // Word text boxes support horizontal and quarter-turn text. Leave
        // arbitrary angles and upside-down text in the page artwork.
        if (Math.abs(originalAngle - angle) > .01 || Math.abs(quarterTurns) > 1) continue;
        const color = separateTextPixels(painted, background, artwork, rasterViewport.scale, x, y, item.width, size, angle);
        if (!color) continue;

        const fontData = page.commonObjs.has(item.fontName) ? page.commonObjs.get(item.fontName) as { name?: string; bold?: boolean; italic?: boolean } : undefined;
        const name = fontData?.name ?? "";
        const family = content.styles[item.fontName]?.fontFamily ?? "";
        const font = wordFont(name, family);
        const bold = Boolean(fontData?.bold || /bold|demi|black|heavy/i.test(name));
        const italic = Boolean(fontData?.italic || /italic|oblique|slanted/i.test(name));
        const wordSize = Math.max(1, Math.round(size * yScale * 2) / 2);
        original.font = `${italic ? "italic " : ""}${bold ? "bold " : ""}${wordSize}px "${font}"`;
        const measured = original.measureText(item.str).width;
        const wordWidth = item.width * xScale;
        const stretch = measured > 0 ? Math.max(25, Math.min(400, Math.round(wordWidth / measured * 100))) : 100;
        const boxWidth = wordWidth + wordSize * .6;
        const boxHeight = wordSize * 1.22;
        const baselineInset = wordSize * .93;
        const rtl = item.dir === "rtl";
        const horizontalCenter = rtl ? wordWidth - boxWidth / 2 : boxWidth / 2;
        const cos = Math.cos(angle);
        const sin = Math.sin(angle);
        const midX = x * xScale + cos * horizontalCenter - sin * (boxHeight / 2 - baselineInset);
        const midY = y * yScale + sin * horizontalCenter + cos * (boxHeight / 2 - baselineInset);
        text.push({
          text: item.str, x: midX - boxWidth / 2, y: midY - boxHeight / 2, width: boxWidth, height: boxHeight,
          size: wordSize, rotation: angle * 180 / Math.PI, stretch, font, bold, italic, rtl, color,
        });
      }
      original.putImageData(artwork, 0, 0);
    }
    signal?.throwIfAborted();
    const image = original.canvas.toDataURL("image/jpeg", .95);
    if (image === "data:,") throw new Error("The browser could not save this PDF page.");
    return { width, height, image, text };
  } finally {
    original.canvas.width = 0;
    original.canvas.height = 0;
    if (withoutText) {
      withoutText.canvas.width = 0;
      withoutText.canvas.height = 0;
    }
  }
}
