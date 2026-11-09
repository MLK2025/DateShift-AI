import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

const pageSource = await readFile(new URL("../app/page.tsx", import.meta.url), "utf8");

test("renders and downloads the same prepared DOCX blob", () => {
  assert.match(pageSource, /import\("docx-preview"\)/);
  assert.match(pageSource, /renderAsync\(blob, container, container,/);
  assert.match(pageSource, /<DocxDocumentPreview blob=\{preparedDocument\.blob\} \/>/);
  assert.match(pageSource, /URL\.createObjectURL\(preparedDocument\.blob\)/);
  assert.doesNotMatch(pageSource, /docx-preview-page|previewSegments|previewText/);
});
