import JSZip from "jszip";

const wordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

export function replaceTextInRuns(nodes: Element[], starts: number[], start: number, end: number, replacement: string) {
  const affected = nodes.flatMap((node, index) => {
    const value = node.textContent ?? "";
    const nodeStart = starts[index];
    const nodeEnd = nodeStart + value.length;
    if (nodeEnd <= start || nodeStart >= end) return [];
    return [{ node, value, from: Math.max(0, start - nodeStart), to: Math.min(value.length, end - nodeStart) }];
  });
  let remaining = replacement;
  affected.forEach(({ node, value, from, to }, index) => {
    const last = index === affected.length - 1;
    const length = last ? remaining.length : to - from;
    const text = value.slice(0, from) + remaining.slice(0, length) + value.slice(to);
    node.textContent = text;
    if (/^\s|\s$/.test(text)) node.setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:space", "preserve");
    remaining = remaining.slice(length);
  });
  return affected[0]?.node ?? null;
}

export async function openWordDocument(file: File) {
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const entry = zip.file("word/document.xml");
  if (!entry) throw new Error("This file does not contain a readable Word document body.");
  return { zip, xml: await entry.async("string") };
}

export function indexWordText(xml: string) {
  const document = new DOMParser().parseFromString(xml, "application/xml");
  if (document.getElementsByTagName("parsererror").length) throw new Error("This Word document contains unreadable XML.");

  const paragraphs: { start: number; end: number; nodes: Element[]; starts: number[] }[] = [];
  let text = "";
  // Group text runs by their own paragraph. A paragraph can contain a drawing
  // with independent nested paragraphs, which must not be indexed a second time.
  let paragraph: Element | null = null;
  let separator = "\n";
  for (const node of Array.from(document.getElementsByTagNameNS(wordNamespace, "*"))) {
    const isBreak = ["tab", "br", "cr"].includes(node.localName) && (node.parentNode as Element | null)?.localName === "r";
    if (!isBreak && node.localName !== "t") continue;
    let owner = node.parentNode as Element | null;
    while (owner && !(owner.localName === "p" && owner.namespaceURI === wordNamespace)) owner = owner.parentNode as Element | null;
    if (!owner) continue;
    if (isBreak) {
      if (owner === paragraph) separator = node.localName === "tab" ? "\t" : "\n";
      paragraph = null;
      continue;
    }
    if (owner !== paragraph) {
      if (paragraphs.length) text += separator;
      separator = "\n";
      paragraphs.push({ start: text.length, end: text.length, nodes: [], starts: [] });
      paragraph = owner;
    }
    const group = paragraphs[paragraphs.length - 1];
    group.nodes.push(node);
    group.starts.push(text.length - group.start);
    text += node.textContent ?? "";
    group.end = text.length;
  }

  return { document, paragraphs, text: text.replace(/\u00a0/g, " ") };
}
