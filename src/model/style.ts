import type {
  AlignmentStyle,
  BorderEdgeStyle,
  BorderStyle,
  BorderWeight,
  FillStyle,
  FontStyle,
  HorizontalAlignment,
  UnderlineStyle,
  VerticalAlignment,
} from './types';

/**
 * Style resolution helpers. The style table holds immutable style records
 * (workbook-model §2); this module resolves the colour, alignment and
 * number-format parts of a record for display.
 */

export const DEFAULT_FONT: FontStyle = Object.freeze({
  name: 'Calibri',
  size: 11,
  bold: false,
  italic: false,
  underline: false,
  underlineStyle: 'none',
  strike: false,
  color: null,
});

export const DEFAULT_FILL: FillStyle = Object.freeze({
  pattern: 'none',
  color: null,
  backgroundColor: null,
});

export const DEFAULT_ALIGNMENT: AlignmentStyle = Object.freeze({
  horizontal: 'general',
  vertical: 'bottom',
  wrapText: false,
  indent: 0,
  textRotation: 0,
  shrinkToFit: false,
});

const NO_BORDER_EDGE: BorderEdgeStyle = Object.freeze({
  style: 'none',
  color: null,
  weight: null,
});

export const DEFAULT_BORDER: BorderStyle = Object.freeze({
  left: NO_BORDER_EDGE,
  right: NO_BORDER_EDGE,
  top: NO_BORDER_EDGE,
  bottom: NO_BORDER_EDGE,
  diagonal: NO_BORDER_EDGE,
  diagonalUp: false,
  diagonalDown: false,
});

/** The legacy indexed colour palette used by `indexed="n"` colour attributes. */
export const INDEXED_PALETTE: readonly string[] = Object.freeze([
  '000000',
  'FFFFFF',
  'FF0000',
  '00FF00',
  '0000FF',
  'FFFF00',
  'FF00FF',
  '00FFFF',
  '000000',
  'FFFFFF',
  'FF0000',
  '00FF00',
  '0000FF',
  'FFFF00',
  'FF00FF',
  '00FFFF',
  '800000',
  '008000',
  '000080',
  '808000',
  '800080',
  '008080',
  'C0C0C0',
  '808080',
  '9999FF',
  '993366',
  'FFFFCC',
  'CCFFFF',
  '660066',
  'FF8080',
  '0066CC',
  'CCCCFF',
  '000080',
  'FF00FF',
  'FFFF00',
  '00FFFF',
  '800080',
  '800000',
  '008080',
  '0000FF',
  '00CCFF',
  'CCFFFF',
  'CCFFCC',
  'FFFF99',
  '99CCFF',
  'FF99CC',
  'CC99FF',
  'FFCC99',
  '3366FF',
  '33CCCC',
  '99CC00',
  'FFCC00',
  'FF6600',
  '666699',
  '969696',
  '003366',
  '339966',
  '003300',
  '333300',
  '993300',
  '993366',
  '333399',
  '333333',
]);

/**
 * The default Office theme palette, in the order OOXML theme indices use:
 * background 1, text 1, background 2, text 2, accents 1-6, hyperlink,
 * followed hyperlink. A workbook-specific theme is not resolved here yet; its
 * references are preserved on save.
 */
export const THEME_PALETTE: readonly string[] = Object.freeze([
  'FFFFFF',
  '000000',
  'E7E6E6',
  '44546A',
  '4472C4',
  'ED7D31',
  'A5A5A5',
  'FFC000',
  '5B9BD5',
  '70AD47',
  '0563C1',
  '954F72',
]);

export interface ColorAttributes {
  readonly rgb?: string | null | undefined;
  readonly indexed?: string | null | undefined;
  readonly theme?: string | null | undefined;
  readonly tint?: string | null | undefined;
  readonly auto?: string | null | undefined;
}

/**
 * Resolve an OOXML colour attribute set to a CSS colour, or `null`.
 * `indexedPalette` overrides the legacy built-in palette when the workbook
 * declares `<colors><indexedColors>`.
 */
export function resolveColor(
  attributes: ColorAttributes,
  indexedPalette: readonly string[] | null = null,
): string | null {
  const { rgb, indexed, theme, tint, auto } = attributes;
  let hex: string | null = null;
  if (rgb !== undefined && rgb !== null && rgb !== '') {
    hex = normalizeHex(rgb);
  } else if (indexed !== undefined && indexed !== null && indexed !== '') {
    const index = Number(indexed);
    const palette = indexedPalette ?? INDEXED_PALETTE;
    if (Number.isInteger(index) && index >= 0 && index < palette.length) {
      hex = palette[index] ?? null;
    }
  } else if (theme !== undefined && theme !== null && theme !== '') {
    const index = Number(theme);
    if (Number.isInteger(index) && index >= 0 && index < THEME_PALETTE.length) {
      hex = THEME_PALETTE[index] ?? null;
    }
  } else if (auto !== undefined && auto !== null && auto !== '0') {
    return null;
  }
  if (hex === null) return null;
  const tintValue = tint === undefined || tint === null ? 0 : Number(tint);
  if (Number.isFinite(tintValue) && tintValue !== 0) {
    return applyTint(hex, tintValue);
  }
  return `#${hex}`;
}

function normalizeHex(value: string): string | null {
  const trimmed = value.trim();
  if (/^[0-9a-f]{8}$/i.test(trimmed)) return trimmed.slice(2).toUpperCase();
  if (/^[0-9a-f]{6}$/i.test(trimmed)) return trimmed.toUpperCase();
  return null;
}

/** Excel's tint: a luminance shift applied in HSL space. */
export function applyTint(hex: string, tint: number): string {
  const red = Number.parseInt(hex.slice(0, 2), 16) / 255;
  const green = Number.parseInt(hex.slice(2, 4), 16) / 255;
  const blue = Number.parseInt(hex.slice(4, 6), 16) / 255;
  const [hue, saturation, lightness] = rgbToHsl(red, green, blue);
  const shifted = tint < 0 ? lightness * (1 + tint) : lightness * (1 - tint) + tint;
  const [r, g, b] = hslToRgb(hue, saturation, Math.min(1, Math.max(0, shifted)));
  const toHex = (channel: number): string =>
    Math.round(Math.min(1, Math.max(0, channel)) * 255)
      .toString(16)
      .padStart(2, '0')
      .toUpperCase();
  return `#${toHex(r)}${toHex(g)}${toHex(b)}`;
}

function rgbToHsl(red: number, green: number, blue: number): [number, number, number] {
  const max = Math.max(red, green, blue);
  const min = Math.min(red, green, blue);
  const lightness = (max + min) / 2;
  if (max === min) return [0, 0, lightness];
  const delta = max - min;
  const saturation = lightness > 0.5 ? delta / (2 - max - min) : delta / (max + min);
  let hue: number;
  if (max === red) hue = (green - blue) / delta + (green < blue ? 6 : 0);
  else if (max === green) hue = (blue - red) / delta + 2;
  else hue = (red - green) / delta + 4;
  return [hue / 6, saturation, lightness];
}

function hslToRgb(hue: number, saturation: number, lightness: number): [number, number, number] {
  if (saturation === 0) return [lightness, lightness, lightness];
  const q =
    lightness < 0.5
      ? lightness * (1 + saturation)
      : lightness + saturation - lightness * saturation;
  const p = 2 * lightness - q;
  const channel = (t: number): number => {
    let value = t;
    if (value < 0) value += 1;
    if (value > 1) value -= 1;
    if (value < 1 / 6) return p + (q - p) * 6 * value;
    if (value < 1 / 2) return q;
    if (value < 2 / 3) return p + (q - p) * (2 / 3 - value) * 6;
    return p;
  };
  return [channel(hue + 1 / 3), channel(hue), channel(hue - 1 / 3)];
}

/**
 * The built-in number-format codes, by `numFmtId` (number-formats §2).
 *
 * ECMA-376 table ids 0-22 and 37-49 are the all-language entries. The
 * locale-specific ids 27-36, 50-58, 59-62 and 67-81 are resolved to their
 * Gregorian en-US equivalent, which is the v1 pinned policy: the id is
 * preserved on save and Thai/Buddhist-era behaviour is not reproduced.
 */
export const BUILTIN_NUMBER_FORMATS: Readonly<Record<number, string>> = Object.freeze({
  0: 'General',
  1: '0',
  2: '0.00',
  3: '#,##0',
  4: '#,##0.00',
  5: '$#,##0_);($#,##0)',
  6: '$#,##0_);[Red]($#,##0)',
  7: '$#,##0.00_);($#,##0.00)',
  8: '$#,##0.00_);[Red]($#,##0.00)',
  9: '0%',
  10: '0.00%',
  11: '0.00E+00',
  12: '# ?/?',
  13: '# ??/??',
  14: 'm/d/yy',
  15: 'd-mmm-yy',
  16: 'd-mmm',
  17: 'mmm-yy',
  18: 'h:mm AM/PM',
  19: 'h:mm:ss AM/PM',
  20: 'h:mm',
  21: 'h:mm:ss',
  22: 'm/d/yy h:mm',
  // 27-36: East Asian date/time ids, rendered as Gregorian en-US equivalents.
  27: 'm/d/yy',
  28: 'm/d/yy',
  29: 'm/d/yy',
  30: 'm/d/yy',
  31: 'yyyy-mm-dd',
  32: 'h:mm',
  33: 'h:mm:ss',
  34: 'h:mm AM/PM',
  35: 'h:mm:ss AM/PM',
  36: 'm/d/yy',
  37: '#,##0 ;(#,##0)',
  38: '#,##0 ;[Red](#,##0)',
  39: '#,##0.00_);(#,##0.00)',
  40: '#,##0.00_);[Red](#,##0.00)',
  41: '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)',
  42: '_("$"* #,##0_);_("$"* (#,##0);_("$"* "-"_);_(@_)',
  43: '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)',
  44: '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)',
  45: 'mm:ss',
  46: '[h]:mm:ss',
  47: 'mmss.0',
  48: '##0.0E+0',
  49: '@',
  // 50-58: East Asian locale ids, rendered as Gregorian en-US equivalents.
  50: 'm/d/yy',
  51: 'm/d/yy',
  52: 'h:mm',
  53: 'h:mm:ss',
  54: 'm/d/yy',
  55: 'h:mm',
  56: 'h:mm:ss',
  57: 'm/d/yy',
  58: 'm/d/yy',
  // 59-62, 67-81: Thai ids, rendered as Gregorian en-US equivalents.
  59: '0',
  60: '0.00',
  61: '#,##0',
  62: '#,##0.00',
  67: '0%',
  68: '0.00%',
  69: '# ?/?',
  70: '# ??/??',
  71: 'd/m/yyyy',
  72: 'd-mmm-yy',
  73: 'd-mmm',
  74: 'mmm-yy',
  75: 'h:mm',
  76: 'h:mm:ss',
  77: 'd/m/yyyy h:mm',
  78: 'mm:ss',
  79: '[h]:mm:ss',
  80: 'mmss.0',
  81: 'd/m/yyyy',
});

/** The locale-specific built-in ids whose rendering is a pinned v1 deviation. */
export const LOCALE_BUILTIN_NUMBER_FORMAT_IDS: ReadonlySet<number> = new Set([
  27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 67,
  68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81,
]);

const CANONICAL_NUMBER_FORMAT_IDS: ReadonlyMap<string, number> = (() => {
  const map = new Map<string, number>();
  for (const id of Object.keys(BUILTIN_NUMBER_FORMATS).map(Number)) {
    if (LOCALE_BUILTIN_NUMBER_FORMAT_IDS.has(id)) continue;
    const code = BUILTIN_NUMBER_FORMATS[id];
    if (code !== undefined && !map.has(code)) map.set(code, id);
  }
  return map;
})();

/** The built-in format code for an id, or `null` when the id is not built in. */
export function builtinNumberFormat(numFmtId: number): string | null {
  return BUILTIN_NUMBER_FORMATS[numFmtId] ?? null;
}

/** True when the id is one of the locale-specific built-ins with a pinned en-US fallback. */
export function isLocaleBuiltinNumberFormat(numFmtId: number): boolean {
  return LOCALE_BUILTIN_NUMBER_FORMAT_IDS.has(numFmtId);
}

/** The canonical built-in id for a format code, or `null` for a custom code. */
export function builtinNumberFormatId(code: string | null): number | null {
  if (code === null || code.trim() === '' || code.trim() === 'General') return 0;
  return CANONICAL_NUMBER_FORMAT_IDS.get(code) ?? null;
}

/** The visual weight implied by an OOXML border style. */
export function borderWeight(style: string): BorderWeight | null {
  switch (style) {
    case 'hair':
      return 'hair';
    case 'thin':
    case 'dashed':
    case 'dotted':
    case 'dashDot':
    case 'dashDotDot':
      return 'thin';
    case 'medium':
    case 'mediumDashed':
    case 'mediumDashDot':
    case 'mediumDashDotDot':
    case 'slantDashDot':
      return 'medium';
    case 'thick':
    case 'double':
      return 'thick';
    default:
      return null;
  }
}

/** Build a border edge with its derived weight. */
export function borderEdge(style: string, color: string | null): BorderEdgeStyle {
  return { style, color, weight: borderWeight(style) };
}

/** Parse an OOXML `<u val="...">` value into the display underline flags. */
export function underlineFromValue(value: string | null | undefined): {
  underline: boolean;
  underlineStyle: UnderlineStyle;
} {
  switch (value) {
    case null:
    case undefined:
    case '':
    case 'single':
      return { underline: true, underlineStyle: 'single' };
    case 'none':
      return { underline: false, underlineStyle: 'none' };
    case 'double':
      return { underline: true, underlineStyle: 'double' };
    case 'singleAccounting':
      return { underline: true, underlineStyle: 'singleAccounting' };
    case 'doubleAccounting':
      return { underline: true, underlineStyle: 'doubleAccounting' };
    default:
      return { underline: true, underlineStyle: 'single' };
  }
}

export function normalizeHorizontalAlignment(
  value: string | null | undefined,
): HorizontalAlignment {
  switch (value) {
    case 'left':
    case 'center':
    case 'right':
    case 'fill':
    case 'justify':
    case 'centerContinuous':
    case 'distributed':
      return value;
    default:
      return 'general';
  }
}

export function normalizeVerticalAlignment(value: string | null | undefined): VerticalAlignment {
  switch (value) {
    case 'top':
    case 'center':
    case 'bottom':
    case 'justify':
    case 'distributed':
      return value;
    default:
      return 'bottom';
  }
}
