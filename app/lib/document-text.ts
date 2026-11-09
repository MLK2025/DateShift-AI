export type PdfColor = {
  red: number;
  green: number;
  blue: number;
};

export type PdfTextSpan = {
  pageNumber: number;
  start: number;
  end: number;
  text: string;
  x: number;
  y: number;
  width: number;
  height: number;
  rotation: number;
  backgroundColor?: PdfColor;
  textColor?: PdfColor;
};

export type DocumentExtraction = {
  kind: "docx" | "pdf";
  text: string;
  pdfSpans: PdfTextSpan[];
};

type PdfTextItem = {
  str?: string;
  hasEOL?: boolean;
  width?: number;
  height?: number;
  transform?: number[];
};

function colorDistance(first: PdfColor, second: PdfColor) {
  return Math.hypot(first.red - second.red, first.green - second.green, first.blue - second.blue);
}

function sampledColorStyle(context: CanvasRenderingContext2D, x: number, baselineY: number, width: number, height: number) {
  const padding = 3;
  const left = Math.max(0, Math.floor(x) - padding);
  const top = Math.max(0, Math.floor(baselineY - height) - padding);
  const right = Math.min(context.canvas.width, Math.ceil(x + width) + padding);
  const bottom = Math.min(context.canvas.height, Math.ceil(baselineY) + padding);
  const sampleWidth = right - left;
  const sampleHeight = bottom - top;
  if (sampleWidth < 2 || sampleHeight < 2) return {};

  const pixels = context.getImageData(left, top, sampleWidth, sampleHeight).data;
  const buckets = new Map<string, { count: number; red: number; green: number; blue: number }>();
  const addPixel = (pixelX: number, pixelY: number, target: Map<string, { count: number; red: number; green: number; blue: number }>) => {
    const index = (pixelY * sampleWidth + pixelX) * 4;
    if (pixels[index + 3] < 220) return;
    const red = pixels[index];
    const green = pixels[index + 1];
    const blue = pixels[index + 2];
    const key = `${Math.round(red / 16)}-${Math.round(green / 16)}-${Math.round(blue / 16)}`;
    const bucket = target.get(key) ?? { count: 0, red: 0, green: 0, blue: 0 };
    bucket.count += 1;
    bucket.red += red;
    bucket.green += green;
    bucket.blue += blue;
    target.set(key, bucket);
  };

  for (let pixelY = 0; pixelY < sampleHeight; pixelY += 1) {
    for (let pixelX = 0; pixelX < sampleWidth; pixelX += 1) {
      const onEdge = pixelX < 2 || pixelY < 2 || pixelX >= sampleWidth - 2 || pixelY >= sampleHeight - 2;
      if (onEdge) addPixel(pixelX, pixelY, buckets);
    }
  }

  const backgroundBucket = [...buckets.values()].sort((first, second) => second.count - first.count)[0];
  if (!backgroundBucket) return {};
  const backgroundColor = {
    red: backgroundBucket.red / backgroundBucket.count / 255,
    green: backgroundBucket.green / backgroundBucket.count / 255,
    blue: backgroundBucket.blue / backgroundBucket.count / 255,
  };

  const foregroundBuckets = new Map<string, { count: number; red: number; green: number; blue: number }>();
  for (let pixelY = padding; pixelY < sampleHeight - padding; pixelY += 1) {
    for (let pixelX = padding; pixelX < sampleWidth - padding; pixelX += 1) addPixel(pixelX, pixelY, foregroundBuckets);
  }
  const textBucket = [...foregroundBuckets.values()]
    .map((bucket) => {
      const color = { red: bucket.red / bucket.count / 255, green: bucket.green / bucket.count / 255, blue: bucket.blue / bucket.count / 255 };
      return { bucket, color, score: colorDistance(color, backgroundColor) * Math.sqrt(bucket.count) };
    })
    .filter((candidate) => colorDistance(candidate.color, backgroundColor) > 0.12)
    .sort((first, second) => second.score - first.score)[0];

  return { backgroundColor, textColor: textBucket?.color };
}

function extensionOf(name: string) {
  return name.split(".").pop()?.toLowerCase() ?? "";
}

async function readDocx(file: File): Promise<DocumentExtraction> {
  const mammoth = await import("mammoth");
  const result = await mammoth.extractRawText({ arrayBuffer: await file.arrayBuffer() });
  return {
    kind: "docx",
    text: result.value.replace(/\u00a0/g, " ").replace(/\n{3,}/g, "\n\n").trim(),
    pdfSpans: [],
  };
}

async function readPdf(file: File): Promise<DocumentExtraction> {
  const pdfjs = await import("pdfjs-dist");
  const workerModule = await import("pdfjs-dist/build/pdf.worker.min.mjs?url");
  pdfjs.GlobalWorkerOptions.workerSrc = workerModule.default;

  const task = pdfjs.getDocument({
    data: new Uint8Array(await file.arrayBuffer()),
    isEvalSupported: false,
    useWorkerFetch: false,
  });
  const pdf = await task.promise;
  const spans: PdfTextSpan[] = [];
  let text = "";

  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    if (text) text += "\n\n";
    const page = await pdf.getPage(pageNumber);
    const content = await page.getTextContent();
    const viewport = page.getViewport({ scale: 1 });
    const canvas = document.createElement("canvas");
    canvas.width = Math.ceil(viewport.width);
    canvas.height = Math.ceil(viewport.height);
    const context = canvas.getContext("2d", { willReadFrequently: true });
    if (context) await page.render({ canvas, viewport }).promise;

    for (const item of content.items as PdfTextItem[]) {
      if (!item.str) continue;
      const transform = item.transform ?? [1, 0, 0, 10, 0, 0];
      const x = transform[4] ?? 0;
      const y = transform[5] ?? 0;
      const width = Math.max(item.width ?? 0, 1);
      const height = Math.max(item.height ?? Math.hypot(transform[2] ?? 0, transform[3] ?? 10), 1);
      const [viewportX, viewportY] = viewport.convertToViewportPoint(x, y);
      const colors = context ? sampledColorStyle(context, viewportX, viewportY, width, height) : {};
      const start = text.length;
      text += item.str;
      spans.push({
        pageNumber,
        start,
        end: text.length,
        text: item.str,
        x,
        y,
        width,
        height,
        rotation: Math.atan2(transform[1] ?? 0, transform[0] ?? 1) * 180 / Math.PI,
        ...colors,
      });
      text += item.hasEOL ? "\n" : " ";
    }
  }

  await task.destroy();
  return { kind: "pdf", text: text.replace(/\u00a0/g, " "), pdfSpans: spans };
}

export async function extractDocument(file: File) {
  const extension = extensionOf(file.name);
  if (extension === "docx") return readDocx(file);
  if (extension === "pdf") return readPdf(file);
  throw new Error("This private mode supports .docx and .pdf files. Save older .doc files as .docx first.");
}
