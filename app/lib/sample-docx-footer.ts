import type JSZip from "jszip";

const word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const office = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const relationships = "http://schemas.openxmlformats.org/package/2006/relationships";
const types = "http://schemas.openxmlformats.org/package/2006/content-types";
const relsPath = "word/_rels/document.xml.rels";
const sampleColor = "9C3C1E";
const marker = `<w:p xmlns:w="${word}"><w:pPr><w:spacing w:before="0" w:after="0" w:line="200" w:lineRule="exact"/><w:ind w:right="300"/><w:jc w:val="right"/></w:pPr><w:r><w:rPr><w:b/><w:vanish w:val="0"/><w:color w:val="${sampleColor}"/><w:sz w:val="20"/></w:rPr><w:t>SAMPLE</w:t></w:r></w:p>`;

function xmlDocument(xml: string, namespace: string, root: string) {
  const document = new DOMParser().parseFromString(xml, "application/xml");
  if (document.documentElement?.namespaceURI !== namespace || document.documentElement.localName !== root) {
    throw new Error("The Word template has an unreadable footer or package manifest.");
  }
  return document;
}

function enabled(element?: Element) {
  return Boolean(element && !/^(0|false|off)$/i.test(element.getAttributeNS(word, "val") ?? ""));
}

export async function addSampleFooter(zip: JSZip, document: Document) {
  const sourceRels = await zip.file(relsPath)?.async("string");
  const rels = xmlDocument(sourceRels ?? `<Relationships xmlns="${relationships}"/>`, relationships, "Relationships");
  const entries = Array.from(rels.getElementsByTagNameNS(relationships, "Relationship"));
  const footerPaths = new Map(entries.filter((rel) => rel.getAttribute("Type") === `${office}/footer` && rel.getAttribute("TargetMode") !== "External")
    .map((rel) => {
      const target = rel.getAttribute("Target") ?? "";
      const path = target.startsWith("/") ? target.slice(1) : `word/${target}`;
      return [rel.getAttribute("Id") ?? "", path];
    }));
  const settingsXml = await zip.file("word/settings.xml")?.async("string");
  const settings = settingsXml ? new DOMParser().parseFromString(settingsXml, "application/xml") : null;
  const hasEvenPages = enabled(settings?.getElementsByTagNameNS(word, "evenAndOddHeaders")[0]);
  const sections = Array.from(document.getElementsByTagNameNS(word, "sectPr")).filter((section) =>
    ["body", "pPr"].includes((section.parentNode as Element | null)?.localName ?? ""));
  if (!sections.length) {
    const body = document.getElementsByTagNameNS(word, "body")[0];
    if (!body) throw new Error("The Word template has no document body.");
    sections.push(body.appendChild(document.createElementNS(word, "w:sectPr")));
  }

  let newId: string | undefined;
  let newPath: string | undefined;
  const referenced = new Set<string>();
  const inherited = new Map<string, string>();
  for (const section of sections) {
    for (const reference of Array.from(section.getElementsByTagNameNS(word, "footerReference"))) {
      const kind = reference.getAttributeNS(word, "type") || "default";
      const id = reference.getAttributeNS(office, "id");
      if (!id || !footerPaths.has(id)) throw new Error("A footer linked by the Word template could not be read.");
      inherited.set(kind, id);
      referenced.add(id);
    }
    const variants = ["default"];
    if (hasEvenPages) variants.push("even");
    if (enabled(section.getElementsByTagNameNS(word, "titlePg")[0])) variants.push("first");
    for (const kind of variants) {
      if (inherited.has(kind)) continue;
      if (!newId) {
        let number = 1;
        while (entries.some((entry) => entry.getAttribute("Id") === `rIdDateShiftSample${number}`) || zip.file(`word/dateshift-sample-footer${number}.xml`)) number += 1;
        newId = `rIdDateShiftSample${number}`;
        newPath = `word/dateshift-sample-footer${number}.xml`;
        footerPaths.set(newId, newPath);
      }
      const ref = document.createElementNS(word, "w:footerReference");
      ref.setAttributeNS(word, "w:type", kind);
      ref.setAttributeNS(office, "r:id", newId);
      section.insertBefore(ref, section.firstChild);
      inherited.set(kind, newId);
      referenced.add(newId);
    }
  }

  const serializer = new XMLSerializer();
  const mark = xmlDocument(marker, word, "p").documentElement;
  for (const id of referenced) {
    const path = footerPaths.get(id)!;
    const previous = await zip.file(path)?.async("string");
    if (!previous && id !== newId) throw new Error("The Word template is missing a linked footer.");
    const footer = xmlDocument(previous ?? `<w:ftr xmlns:w="${word}"/>`, word, "ftr");
    const hasSample = Array.from(footer.getElementsByTagNameNS(word, "p")).some((p) =>
      p.parentNode === footer.documentElement && p.textContent?.trim() === "SAMPLE"
      && p.getElementsByTagNameNS(word, "color")[0]?.getAttributeNS(word, "val") === sampleColor);
    if (!hasSample) footer.documentElement.appendChild(footer.importNode(mark, true));
    zip.file(path, serializer.serializeToString(footer));
  }

  if (newId && newPath) {
    const rel = rels.createElementNS(relationships, "Relationship");
    rel.setAttribute("Id", newId);
    rel.setAttribute("Type", `${office}/footer`);
    rel.setAttribute("Target", newPath.slice("word/".length));
    rels.documentElement.appendChild(rel);
    zip.file(relsPath, serializer.serializeToString(rels));
    const manifestXml = await zip.file("[Content_Types].xml")?.async("string");
    if (!manifestXml) throw new Error("The Word template is missing its content types.");
    const manifest = xmlDocument(manifestXml, types, "Types");
    const override = manifest.createElementNS(types, "Override");
    override.setAttribute("PartName", `/${newPath}`);
    override.setAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml");
    manifest.documentElement.appendChild(override);
    zip.file("[Content_Types].xml", serializer.serializeToString(manifest));
  }
}
