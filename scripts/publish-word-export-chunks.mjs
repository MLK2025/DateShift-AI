import { access, readFile, readdir, writeFile } from "node:fs/promises";

const directory = new URL("../dist/client/_next/static/chunks/", import.meta.url);
const entries = await readdir(directory);
const pages = await Promise.all(entries.filter((entry) => /^page-[\w-]+\.js$/.test(entry))
  .map((entry) => readFile(new URL(entry, directory), "utf8")));
const exporters = new Set(pages.flatMap((page) => page.match(/mock-ledger-docx-[\w-]+\.js/g) ?? []));
if (exporters.size !== 1) throw new Error("Expected one Word exporter referenced by the DateShift page.");

const [current] = exporters;
await access(new URL(current, directory));
const bridge = `export { createSampleLedgerDocument } from "./${current}";\n`;
// Keep these previously published import URLs available to open tabs.
const aliases = ["mock-ledger-current.js", "mock-ledger-docx-DiQvRnSa.js", "mock-ledger-docx-C2gC3-c4.js"];
await Promise.all(aliases.filter((entry) => entry !== current)
  .map((entry) => writeFile(new URL(entry, directory), bridge)));
