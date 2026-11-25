export type WordFont = { font: string; bold: boolean; italic: boolean };

export const wordFontStyles = ["embedRegular", "embedBold", "embedItalic", "embedBoldItalic"] as const;

export function wordFontStyle({ bold, italic }: WordFont) {
  return wordFontStyles[Number(bold) + 2 * Number(italic)];
}

export function wordFontKey(font: WordFont) {
  return JSON.stringify([font.font, wordFontStyle(font)]);
}

export function fontFamilyCss(family: string) {
  return `"${family.replace(/[\\"]/g, "\\$&")}"`;
}

// OOXML obfuscates the first 32 bytes with the font key in reverse byte order.
export function obfuscateWordFont(bytes: Uint8Array, key: string) {
  const hex = key.replace(/[{}-]/g, "");
  if (!/^[a-f0-9]{32}$/i.test(hex) || bytes.length < 32) return null;
  const result = bytes.slice();
  for (let i = 0; i < 32; i++) {
    const position = 30 - (i % 16) * 2;
    result[i] ^= Number.parseInt(hex.slice(position, position + 2), 16);
  }
  return result;
}
