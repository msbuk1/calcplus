/**
 * XML 1.0 text validity. Cell text is stored in XML parts, so an edit that
 * cannot be represented must be refused before a save starts.
 */
export function findInvalidXmlCharacter(value: string): string | null {
  for (let index = 0; index < value.length; index += 1) {
    const code = value.charCodeAt(index);
    if (code >= 0xd800 && code <= 0xdbff) {
      const next = value.charCodeAt(index + 1);
      if (!(next >= 0xdc00 && next <= 0xdfff)) return formatCodePoint(code);
      index += 1;
      continue;
    }
    if (code >= 0xdc00 && code <= 0xdfff) return formatCodePoint(code);
    const allowed =
      code === 0x09 ||
      code === 0x0a ||
      code === 0x0d ||
      (code >= 0x20 && code <= 0xd7ff) ||
      (code >= 0xe000 && code <= 0xfffd);
    if (!allowed) return formatCodePoint(code);
  }
  return null;
}

export function invalidXmlCharacterMessage(value: string, context = 'Cell text'): string | null {
  const invalid = findInvalidXmlCharacter(value);
  return invalid === null
    ? null
    : `${context} contains ${invalid}, which XML cannot store. The file was not written.`;
}

function formatCodePoint(code: number): string {
  return `an invalid XML character (U+${code.toString(16).toUpperCase().padStart(4, '0')})`;
}
