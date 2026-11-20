/**
 * A minimal typing for the two fontkit calls the PDF export uses. fontkit 3
 * ships its implementation without types, and the export only needs to parse a
 * face and ask it about glyph coverage (research/26 §4.2).
 */
declare module 'fontkit' {
  export interface FontkitFont {
    readonly postscriptName?: string;
    readonly familyName?: string;
    hasGlyphForCodePoint(codePoint: number): boolean;
  }

  export function create(buffer: Uint8Array): FontkitFont;
}
