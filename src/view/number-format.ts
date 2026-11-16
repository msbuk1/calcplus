import {
  TOKEN_DURATION,
  format as numfmtFormat,
  formatColor as numfmtFormatColor,
  getFormatInfo,
  isValidFormat,
  tokenize,
} from 'numfmt';
import type { CellValue } from '@model/value';
import { INDEXED_PALETTE } from '@model/style';

/**
 * The number-format seam (technical-design §3.6).
 *
 * `numfmt` owns the Excel grammar. This module owns the Excel seams that
 * `numfmt` does not: the checked-in built-in id table is resolved by the
 * style importer, and here we add the 1904 date adapter, code normalisation,
 * width layout, legacy indexed colours, Indian grouping and the invalid-code
 * fallback to `General`.
 */
export interface NumberFormatter {
  format(value: CellValue, code: string | null, date1904: boolean): string;
}

/** A formatted value plus the layout/colour data a cell painter needs. */
export interface FormattedNumber {
  readonly text: string;
  readonly color: string | null;
  readonly overflow: boolean;
}

export interface LayoutContext {
  /** The cell's content-box width in CSS pixels. */
  readonly widthPx: number;
  /** Measures text in the cell's font. */
  readonly measure: (text: string) => number;
  readonly date1904: boolean;
  readonly indexedPalette?: readonly string[] | null;
  /** Wrapped and shrink-to-fit cells never paint the `#` overflow marker. */
  readonly wrapText?: boolean;
  readonly shrinkToFit?: boolean;
}

const NAMED_COLOURS: Readonly<Record<string, string>> = Object.freeze({
  black: '#000000',
  blue: '#0000FF',
  cyan: '#00FFFF',
  green: '#00FF00',
  magenta: '#FF00FF',
  red: '#FF0000',
  white: '#FFFFFF',
  yellow: '#FFFF00',
});

/** Locales whose convention is Indian 3/2 digit grouping. */
const INDIAN_GROUPING_LOCALES: ReadonlySet<number> = new Set([
  0x4009, // en-IN
  0x439, // hi-IN
  0x445, // bn-IN
  0x446, // pa-IN
  0x447, // gu-IN
  0x448, // or-IN
  0x449, // ta-IN
  0x44a, // te-IN
  0x44b, // kn-IN
  0x44c, // ml-IN
  0x44e, // mr-IN
  0x45a, // as-IN
]);

const BASE_OPTIONS = Object.freeze({
  dateSpanLarge: false,
  dateErrorNumber: false,
  dateErrorThrows: false,
  throws: false,
  leap1900: true,
  nbsp: false,
  grouping: [3, 3] as [number, number],
});

/** The colour a format code applies to the cell text, or `null`. */
export function numberFormatColor(
  code: string | null,
  indexedPalette: readonly string[] | null = null,
): string | null {
  return (
    numberFormatColorForValue(code, 1, indexedPalette) ??
    numberFormatColorForText(code, indexedPalette)
  );
}

/** The colour the selected numeric section applies to one value, or `null`. */
export function numberFormatColorForValue(
  code: string | null,
  value: number,
  indexedPalette: readonly string[] | null = null,
): string | null {
  return resolveFormatColor(code, value, indexedPalette);
}

/** The colour the text section applies, or `null` when it declares none. */
export function numberFormatColorForText(
  code: string | null,
  indexedPalette: readonly string[] | null = null,
): string | null {
  return resolveFormatColor(code, 'text', indexedPalette);
}

function resolveFormatColor(
  code: string | null,
  value: number | string,
  indexedPalette: readonly string[] | null,
): string | null {
  if (code === null || code.trim() === '') return null;
  const normalized = normalizeFormatCode(code);
  try {
    const color = numfmtFormatColor(normalized, value, {
      throws: false,
      indexColors: false,
    });
    if (typeof color === 'number') {
      const palette = indexedPalette ?? INDEXED_PALETTE;
      return palette[color] === undefined ? null : `#${palette[color]!.toUpperCase()}`;
    }
    if (typeof color === 'string') {
      if (color.startsWith('#')) return color.toUpperCase();
      return NAMED_COLOURS[color.toLowerCase()] ?? null;
    }
  } catch {
    return null;
  }
  return null;
}

export function createNumberFormatter(): NumberFormatter {
  return { format: formatValue };
}

export function formatValue(value: CellValue, code: string | null, date1904 = false): string {
  switch (value.type) {
    case 'number':
      return formatNumberValue(value.value, code, date1904);
    case 'text':
      return formatTextValue(value.value, code);
    case 'boolean':
      return value.value ? 'TRUE' : 'FALSE';
    case 'error':
      return value.value;
    case 'blank':
      return '';
  }
}

/** Excel's general format, through the same engine used for cells. */
export function formatGeneral(value: number): string {
  try {
    return numfmtFormat('General', value, { ...BASE_OPTIONS, locale: 'en-US' });
  } catch {
    return String(value);
  }
}

/**
 * Format a numeric cell for layout. This is the width-layout seam: `*x` fills
 * the remaining width, `_x` advances by the width of `x`, and a value wider
 * than the content box becomes Excel's `#` overflow fill.
 */
export function formatForLayout(
  value: CellValue,
  code: string | null,
  context: LayoutContext,
): FormattedNumber {
  if (value.type !== 'number') {
    return {
      text: formatValue(value, code, context.date1904),
      color:
        value.type === 'text'
          ? numberFormatColorForText(code, context.indexedPalette ?? null)
          : numberFormatColor(code, context.indexedPalette ?? null),
      overflow: false,
    };
  }
  const normalized = effectiveCode(code);
  const adjusted = adjustSerial(value.value, normalized, context.date1904);
  const grouping = groupingForCode(normalized);
  const raw = formatValidated(normalized, adjusted, {
    fillChar: '\uE000',
    skipChar: '\uE001',
    grouping,
    locale: localeForCode(normalized),
  });
  const expanded = expandPadding(raw, context.widthPx, context.measure);
  const fits =
    context.wrapText === true ||
    context.shrinkToFit === true ||
    context.measure(expanded) <= context.widthPx;
  return {
    text: fits ? expanded : hashFill(context.widthPx, context.measure),
    color: numberFormatColorForValue(code, value.value, context.indexedPalette ?? null),
    overflow: !fits,
  };
}

function formatNumberValue(value: number, code: string | null, date1904: boolean): string {
  const normalized = effectiveCode(code);
  const adjusted = adjustSerial(value, normalized, date1904);
  return formatValidated(normalized, adjusted, {
    grouping: groupingForCode(normalized),
    locale: localeForCode(normalized),
  });
}

function formatTextValue(text: string, code: string | null): string {
  if (code === null || code.trim() === '' || code.trim() === 'General') return text;
  const normalized = englishLocaleCode(normalizeFormatCode(code));
  if (normalized.trim() === '' || !isValidFormat(normalized)) return text;
  try {
    return numfmtFormat(normalized, text, {
      ...BASE_OPTIONS,
      locale: localeForCode(normalized),
      grouping: groupingForCode(normalized),
    });
  } catch {
    return text;
  }
}

function effectiveCode(code: string | null): string {
  if (code === null || code.trim() === '' || code.trim() === 'General') return 'General';
  const normalized = normalizeFormatCode(code);
  if (normalized.trim() === '' || !isValidFormat(normalized)) return 'General';
  return englishLocaleCode(normalized);
}

function formatValidated(
  code: string,
  value: number,
  options: {
    readonly grouping: [number, number] | [number];
    readonly locale: string | number;
    readonly fillChar?: string;
    readonly skipChar?: string;
  },
): string {
  try {
    return numfmtFormat(code, value, {
      ...BASE_OPTIONS,
      locale: options.locale,
      grouping: options.grouping,
      ...(options.fillChar === undefined ? {} : { fillChar: options.fillChar }),
      ...(options.skipChar === undefined ? {} : { skipChar: options.skipChar }),
    });
  } catch {
    return formatGeneral(value);
  }
}

/** Add 1462 days for calendar formats only; elapsed formats keep their value. */
function adjustSerial(value: number, code: string, date1904: boolean): number {
  if (!date1904) return value;
  if (code === 'General') return value;
  try {
    const info = getFormatInfo(code);
    const calendar = info.type === 'date' || info.type === 'datetime' || info.type === 'time';
    if (!calendar) return value;
    const elapsed = tokenize(code).some((token) => token.type === TOKEN_DURATION);
    return elapsed ? value : value + 1462;
  } catch {
    return value;
  }
}

/**
 * Normalise only outside quoted strings and escapes. `numfmt` currently
 * rejects lowercase `e+` and treats lowercase `a/p` as literal text; CJK
 * numeral modifiers are stripped to their ASCII fallback (a pinned v1 limit).
 */
export function normalizeFormatCode(code: string): string {
  let output = '';
  let inQuotes = false;
  for (let index = 0; index < code.length; index += 1) {
    const character = code[index] ?? '';
    if (character === '"') {
      inQuotes = !inQuotes;
      output += character;
      continue;
    }
    if (inQuotes) {
      output += character;
      continue;
    }
    if (character === '\\') {
      output += character;
      const escaped = code[index + 1];
      if (escaped !== undefined) {
        output += escaped;
        index += 1;
      }
      continue;
    }
    if (character === '[') {
      const close = code.indexOf(']', index);
      if (close < 0) {
        output += code.slice(index);
        break;
      }
      const token = code.slice(index, close + 1);
      const body = token.slice(1, -1);
      if (/^(NatNum|DBNum)/i.test(body)) {
        index = close;
        continue;
      }
      output += token;
      index = close;
      continue;
    }
    if (character === 'e' || character === 'E') {
      const next = code[index + 1];
      if (next === '+' || next === '-') {
        output += 'E';
        continue;
      }
    }
    if (
      (character === 'a' || character === 'A') &&
      code.slice(index, index + 3).toLowerCase() === 'a/p'
    ) {
      output += 'A/P';
      index += 2;
      continue;
    }
    output += character;
  }
  return output;
}

/**
 * v1 renders with English locale semantics (number-formats §8). An authored
 * `[$symbol-LCID]` keeps its symbol but is rendered with the en-US LCID, so a
 * German or Japanese code still displays English separators and month names.
 * Indian locales keep their id because the contract owns 3/2 grouping.
 */
function englishLocaleCode(code: string): string {
  return code.replace(
    /\[\$([^\]]*)-([0-9A-Fa-f]{1,4})\]/g,
    (match, prefix: string, idText: string) => {
      const id = Number.parseInt(idText, 16);
      if (INDIAN_GROUPING_LOCALES.has(id) || id === 0xf800 || id === 0xf400) return match;
      return `[$${prefix}-409]`;
    },
  );
}

function localeForCode(code: string): string | number {
  const lcid = localeIdInFormat(code);
  return lcid ?? 'en-US';
}

function groupingForCode(code: string): [number, number] | [number] {
  const lcid = localeIdInFormat(code);
  if (lcid !== null && INDIAN_GROUPING_LOCALES.has(lcid)) return [3, 2];
  return [3, 3];
}

/** The LCID from an OOXML `[$-xxxx]` locale modifier, when present. */
function localeIdInFormat(code: string): number | null {
  const match = /\[\$[^\]]*-([0-9A-Fa-f]{1,8})\]/.exec(code);
  if (match === null) return null;
  const value = Number.parseInt(match[1]!, 16);
  if (!Number.isFinite(value) || value === 0xf800 || value === 0xf400) return null;
  return value;
}

function expandPadding(raw: string, widthPx: number, measure: (text: string) => number): string {
  if (!raw.includes('\uE000') && !raw.includes('\uE001')) return raw;
  let text = '';
  const fillPositions: number[] = [];
  const fillCharacters: string[] = [];
  for (let index = 0; index < raw.length; index += 1) {
    const character = raw[index] ?? '';
    if (character === '\uE001') {
      const skipped = raw[index + 1] ?? ' ';
      text += ' '.repeat(spacesFor(skipped, measure));
      index += 1;
      continue;
    }
    if (character === '\uE000') {
      const fill = raw[index + 1] ?? ' ';
      fillPositions.push(text.length);
      fillCharacters.push(fill);
      index += 1;
      continue;
    }
    text += character;
  }
  if (fillPositions.length === 0) return text;
  const plainWidth = measure(text);
  const remaining = Math.max(0, widthPx - plainWidth);
  const perFill = fillPositions.length === 0 ? 0 : remaining / fillPositions.length;
  let output = text;
  for (let index = fillPositions.length - 1; index >= 0; index -= 1) {
    const character = fillCharacters[index] ?? ' ';
    const count = repeatCountFor(character, perFill, measure);
    output =
      output.slice(0, fillPositions[index]!) +
      character.repeat(count) +
      output.slice(fillPositions[index]!);
  }
  return output;
}

function spacesFor(character: string, measure: (text: string) => number): number {
  const target = measure(character);
  const one = measure(' ');
  if (one <= 0) return 1;
  return Math.max(1, Math.round(target / one));
}

function repeatCountFor(
  character: string,
  width: number,
  measure: (text: string) => number,
): number {
  if (width <= 0) return 0;
  const one = measure(character);
  if (one <= 0) return 0;
  return Math.floor(width / one);
}

function hashFill(widthPx: number, measure: (text: string) => number): string {
  const hashWidth = Math.max(1, measure('#'));
  return '#'.repeat(Math.max(1, Math.floor(widthPx / hashWidth)));
}

/** Split a format code into sections, ignoring separators in literals/brackets. */
export function splitSections(code: string): string[] {
  const sections: string[] = [];
  let current = '';
  let inQuotes = false;
  let inBrackets = false;
  for (const character of code) {
    if (character === '"') {
      inQuotes = !inQuotes;
      current += character;
      continue;
    }
    if (character === '[' && !inQuotes) inBrackets = true;
    if (character === ']' && !inQuotes) inBrackets = false;
    if (character === ';' && !inQuotes && !inBrackets) {
      sections.push(current);
      current = '';
      continue;
    }
    current += character;
  }
  sections.push(current);
  return sections;
}
