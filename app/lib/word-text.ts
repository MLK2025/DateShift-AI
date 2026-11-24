import JSZip from "jszip";

const wordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

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
