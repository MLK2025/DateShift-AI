import { openWordDocument } from "./word-text";
import type { PositionedPage, PositionedText } from "./pdf-word-positioned";
import { obfuscateWordFont, wordFontStyles, type WordFont } from "./word-fonts";

const relNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const wordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

function descendants(node: Document | Element, name: string) {
  return Array.from(node.getElementsByTagNameNS("*", name));
}

function first(node: Document | Element, name: string) {
  return descendants(node, name)[0];
}

function points(value: string | null | undefined) {
  return Number(value ?? "0") / 12700;
}

function readTextBox(anchor: Element): PositionedText | null {
  const content = first(anchor, "txbxContent");
  if (!content) return null;
  const wordRuns = descendants(content, "t");
  const text = wordRuns.map((node) => node.textContent ?? "").join("");
  const properties = first(content, "rPr");
  const attr = (name: string, value = "val") => properties ? first(properties, name)?.getAttributeNS(wordNamespace, value) : null;
  const extent = first(anchor, "extent");
  const positionH = first(anchor, "positionH");
  const positionV = first(anchor, "positionV");
  const direction = first(anchor, "bodyPr")?.getAttribute("vert");
  const rotation = direction === "vert270" ? -90 : direction === "vert" ? 90 : 0;
  const width = points(extent?.getAttribute(rotation ? "cy" : "cx"));
  const height = points(extent?.getAttribute(rotation ? "cx" : "cy"));
  const offset = rotation ? (width - height) / 2 : 0;
  return {
    text, width, height, rotation,
    x: points(positionH?.textContent) - offset,
    y: points(positionV?.textContent) + offset,
    font: attr("rFonts", "ascii") || "Arial",
    size: Number(attr("sz") || "24") / 2,
    stretch: Number(attr("w") || "100"),
    color: attr("color") || "000000",
    bold: Boolean(properties && first(properties, "b")),
    italic: Boolean(properties && first(properties, "i")),
    rtl: Boolean(properties && first(properties, "rtl")),
  };
}

export type PreviewFont = WordFont & { bytes: Uint8Array };

async function readEmbeddedFonts(zip: Awaited<ReturnType<typeof openWordDocument>>["zip"], parser: DOMParser) {
  const [tableXml, relsXml] = await Promise.all([
    zip.file("word/fontTable.xml")?.async("string"),
    zip.file("word/_rels/fontTable.xml.rels")?.async("string"),
  ]);
  if (!tableXml || !relsXml) return [];
  const names = descendants(parser.parseFromString(tableXml, "application/xml"), "font");
  const targets = new Map(descendants(parser.parseFromString(relsXml, "application/xml"), "Relationship")
    .map((relation) => [relation.getAttribute("Id"), relation.getAttribute("Target")]));
  const faces = await Promise.all(names.flatMap((node) => wordFontStyles.map(async (style, index): Promise<PreviewFont | null> => {
    const entry = first(node, style);
    const font = node.getAttributeNS(wordNamespace, "name");
    const key = entry?.getAttributeNS(wordNamespace, "fontKey");
    const path = targets.get(entry?.getAttributeNS(relNamespace, "id") ?? null);
    if (!font || !key || !path || !/^fonts\/embeddedFont\d+\.odttf$/.test(path)) return null;
    try {
      const data = await zip.file(`word/${path}`)?.async("uint8array");
      const bytes = data && obfuscateWordFont(data, key);
      if (bytes) return { bytes, font, bold: Boolean(index & 1), italic: Boolean(index & 2) };
    } catch {
      // A missing font needn't prevent the rest of the document from previewing.
    }
    return null;
  })));
  return faces.filter((face) => face !== null);
}

export async function readPositionedWordPreview(file: Blob) {
  const { zip, xml } = await openWordDocument(file);
  if (!/name="PDF page \d+ artwork"/.test(xml)) return null;
  const parser = new DOMParser();
  const document = parser.parseFromString(xml, "application/xml");
  const relations = await zip.file("word/_rels/document.xml.rels")?.async("string");
  if (!relations) return null;
  const targets = new Map(descendants(parser.parseFromString(relations, "application/xml"), "Relationship")
    .map((rel) => [rel.getAttribute("Id"), rel.getAttribute("Target")]));
  const pages = descendants(document, "anchor").filter((anchor) => /^PDF page \d+ artwork$/.test(first(anchor, "docPr")?.getAttribute("name") ?? ""));
  if (!pages.length) return null;
  const positioned = await Promise.all(pages.map(async (page): Promise<PositionedPage | null> => {
    const imageId = first(page, "blip")?.getAttributeNS(relNamespace, "embed");
    const target = imageId ? targets.get(imageId) : null;
    const path = target?.startsWith("/") ? target.slice(1) : `word/${target ?? ""}`;
    if (!/^word\/media\/.+\.jpeg$/i.test(path)) return null;
    const jpeg = await zip.file(path)?.async("base64");
    const extent = first(page, "extent");
    const paragraph = page.parentNode?.parentNode?.parentNode;
    if (!jpeg || !extent || paragraph?.nodeType !== 1) return null;
    return {
      width: points(extent.getAttribute("cx")),
      height: points(extent.getAttribute("cy")),
      image: `data:image/jpeg;base64,${jpeg}`,
      text: descendants(paragraph as Element, "anchor").flatMap((anchor) => {
        const run = readTextBox(anchor);
        return run ? [run] : [];
      }),
    };
  }));
  if (!positioned.every((page) => page !== null)) return null;
  return { pages: positioned, fonts: await readEmbeddedFonts(zip, parser) };
}
