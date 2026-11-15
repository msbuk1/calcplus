import type { CellValue } from '@model/value';

/**
 * The number-format seam (technical-design §3.6).
 *
 * M2 renders the basic formats: the built-in `numFmtId` table and the simple
 * custom codes made of digit placeholders, separators, percentages, literals
 * and date tokens. The full Excel grammar — conditions, sections, colours,
 * padding, elapsed time, locale — arrives with the number-format milestone,
 * which replaces this module behind the same interface.
 */
export interface NumberFormatter {
  format(value: CellValue, code: string | null, date1904: boolean): string;
}

const COLOUR_TOKENS = new Set([
  'Black',
  'Blue',
  'Cyan',
  'Green',
  'Magenta',
  'Red',
  'White',
  'Yellow',
]);

const COLOUR_VALUES: Readonly<Record<string, string>> = Object.freeze({
  Black: '#000000',
  Blue: '#0000ff',
  Cyan: '#00ffff',
  Green: '#00ff00',
  Magenta: '#ff00ff',
  Red: '#ff0000',
  White: '#ffffff',
  Yellow: '#ffff00',
});

/** The colour a format code applies to the cell text, or `null`. */
export function numberFormatColor(code: string | null): string | null {
  if (code === null) return null;
  const match = /^\s*\[([A-Za-z]+)\]/.exec(code);
  if (match === null) return null;
  const name = match[1] ?? '';
  if (!COLOUR_TOKENS.has(name)) return null;
  return COLOUR_VALUES[name] ?? null;
}

export function createNumberFormatter(): NumberFormatter {
  return { format: formatValue };
}

export function formatValue(value: CellValue, code: string | null, date1904 = false): string {
  if (value.type !== 'number') return '';
  if (code === null || code.trim() === '' || code.trim() === 'General') {
    return formatGeneral(value.value);
  }
  const sections = splitSections(code);
  const selected = selectSection(sections, value.value);
  if (selected === null) return formatGeneral(value.value);
  const cleaned = stripBrackets(selected.code);
  if (cleaned.trim() === '') return formatGeneral(value.value);
  if (isDateFormat(cleaned)) return formatDate(value.value, cleaned, date1904);
  // A negative section carries its own sign or parentheses, so the value is
  // formatted as an absolute number.
  return formatNumber(value.value, cleaned, selected.negative);
}

/** Excel's general format: the number as written, exponential when very large. */
export function formatGeneral(value: number): string {
  if (!Number.isFinite(value)) return String(value);
  if (value === 0) return '0';
  const magnitude = Math.abs(value);
  if (magnitude >= 1e11 || magnitude < 1e-10) {
    return value
      .toExponential(5)
      .replace(/e([+-])/, 'E$1')
      .replace(/(\.\d*?)0+E/, '$1E')
      .replace(/\.E/, 'E');
  }
  return String(value);
}

/** Split a format code into its `;` sections, ignoring separators in literals. */
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

interface SelectedSection {
  /** The section code that applies to the value. */
  readonly code: string;
  /** True when the section is the negative section, which supplies the sign. */
  readonly negative: boolean;
}

function selectSection(sections: readonly string[], value: number): SelectedSection | null {
  if (sections.length === 1) {
    const only = sections[0];
    return only === undefined ? null : { code: only, negative: false };
  }
  if (value < 0 && sections.length >= 2) {
    const negative = sections[1];
    return negative === undefined ? null : { code: negative, negative: true };
  }
  if (value === 0 && sections.length >= 3) {
    const zero = sections[2];
    return zero === undefined ? null : { code: zero, negative: false };
  }
  const positive = sections[0];
  return positive === undefined ? null : { code: positive, negative: false };
}

function stripBrackets(code: string): string {
  return code.replace(/\[[^\]]*\]/g, '');
}

function isDateFormat(code: string): boolean {
  const withoutLiterals = code.replace(/"[^"]*"/g, '');
  return /[ydhs]/i.test(withoutLiterals) || /am\/pm/i.test(withoutLiterals);
}

interface NumberSection {
  /** Forced integer digits: the count of `0` and `?` placeholders. */
  readonly integerDigits: number;
  /** Decimal places: the count of every placeholder after the point. */
  readonly decimalDigits: number;
  /** Forced decimal places: the count of `0` and `?` placeholders. */
  readonly decimalRequired: number;
  readonly grouped: boolean;
  readonly percent: boolean;
  readonly scale: number;
  readonly prefix: string;
  readonly suffix: string;
  readonly textPlaceholder: boolean;
}

function parseNumberSection(code: string): NumberSection {
  let integerDigits = 0;
  let decimalDigits = 0;
  let decimalRequired = 0;
  let grouped = false;
  let percent = false;
  let scale = 1;
  let prefix = '';
  let suffix = '';
  let textPlaceholder = false;
  let seenDigit = false;
  let seenDecimal = false;
  let inQuotes = false;
  const characters = [...code];
  for (let index = 0; index < characters.length; index += 1) {
    const character = characters[index] ?? '';
    if (character === '"') {
      inQuotes = !inQuotes;
      continue;
    }
    if (inQuotes) {
      if (seenDigit) suffix += character;
      else prefix += character;
      continue;
    }
    if (character === '\\') {
      const escaped = characters[index + 1];
      index += 1;
      if (escaped === undefined) continue;
      if (seenDigit) suffix += escaped;
      else prefix += escaped;
      continue;
    }
    if (character === '%') {
      percent = true;
      if (seenDigit) suffix += '%';
      else prefix += '%';
      continue;
    }
    if (character === '0' || character === '?') {
      seenDigit = true;
      if (seenDecimal) {
        decimalDigits += 1;
        decimalRequired += 1;
      } else {
        integerDigits += 1;
      }
      continue;
    }
    if (character === '#') {
      seenDigit = true;
      if (seenDecimal) decimalDigits += 1;
      continue;
    }
    if (character === '.') {
      seenDecimal = true;
      seenDigit = true;
      continue;
    }
    if (character === ',') {
      const next = characters[index + 1];
      if (seenDigit && (next === '0' || next === '#' || next === '?')) {
        grouped = true;
        continue;
      }
      if (seenDigit && !seenDecimal) {
        scale *= 1000;
        continue;
      }
      continue;
    }
    if (character === '@') {
      textPlaceholder = true;
      continue;
    }
    if (character === '*' || character === '_') {
      index += 1;
      continue;
    }
    if (character === '[') {
      while (index < characters.length && characters[index] !== ']') index += 1;
      continue;
    }
    if (seenDigit) suffix += character;
    else prefix += character;
  }
  return {
    integerDigits,
    decimalDigits,
    decimalRequired,
    grouped,
    percent,
    scale,
    prefix,
    suffix,
    textPlaceholder,
  };
}

function formatNumber(value: number, code: string, negativeSection = false): string {
  const section = parseNumberSection(code);
  if (section.textPlaceholder) return section.prefix + String(value) + section.suffix;
  let numeric = value;
  if (section.percent) numeric *= 100;
  numeric /= section.scale;
  const negative = numeric < 0;
  const absolute = Math.abs(numeric);
  const fixed = absolute.toFixed(section.decimalDigits);
  const [integerPart = '0', fractionPart = ''] = fixed.split('.');
  let integerText = integerPart;
  if (section.grouped) integerText = groupThousands(integerText);
  if (integerText.length < section.integerDigits) {
    integerText = integerText.padStart(section.integerDigits, '0');
  }
  let fraction = fractionPart === '' ? '' : `.${fractionPart}`;
  if (fraction !== '' && section.decimalRequired < section.decimalDigits) {
    // `#` placeholders drop trailing zeros; `0` placeholders keep them.
    let trimmed = fraction.slice(1).replace(/0+$/, '');
    while (trimmed.length < section.decimalRequired) trimmed += '0';
    fraction = trimmed === '' ? '' : `.${trimmed}`;
  }
  const sign = negative && !negativeSection ? '-' : '';
  return `${sign}${section.prefix}${integerText}${fraction}${section.suffix}`;
}

function groupThousands(text: string): string {
  return text.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

interface DateParts {
  readonly year: number;
  readonly month: number;
  readonly day: number;
  readonly hours: number;
  readonly minutes: number;
  readonly seconds: number;
}

function serialToDateParts(serial: number, date1904: boolean): DateParts {
  const days = Math.floor(serial);
  const fraction = serial - days;
  const base = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, serial >= 61 ? 30 : 31);
  const time = base + days * 86_400_000 + Math.round(fraction * 86_400_000);
  const date = new Date(time);
  return {
    year: date.getUTCFullYear(),
    month: date.getUTCMonth() + 1,
    day: date.getUTCDate(),
    hours: date.getUTCHours(),
    minutes: date.getUTCMinutes(),
    seconds: date.getUTCSeconds(),
  };
}

const MONTH_NAMES = [
  'Jan',
  'Feb',
  'Mar',
  'Apr',
  'May',
  'Jun',
  'Jul',
  'Aug',
  'Sep',
  'Oct',
  'Nov',
  'Dec',
];

function formatDate(serial: number, code: string, date1904: boolean): string {
  const parts = serialToDateParts(serial, date1904);
  const hasAmPm = /am\/pm|a\/p/i.test(code.replace(/"[^"]*"/g, ''));
  let output = '';
  let index = 0;
  const characters = [...code];
  const minutePending = (): boolean => {
    for (let back = index - 1; back >= 0; back -= 1) {
      const previous = characters[back];
      if (previous === undefined) break;
      if (previous === 'h' || previous === 'H') return true;
      if (previous === '"') break;
    }
    return false;
  };
  while (index < characters.length) {
    const character = characters[index] ?? '';
    if (character === '"') {
      index += 1;
      while (index < characters.length && characters[index] !== '"') {
        output += characters[index];
        index += 1;
      }
      index += 1;
      continue;
    }
    if (character === '\\') {
      index += 1;
      output += characters[index] ?? '';
      index += 1;
      continue;
    }
    if (character === '[') {
      while (index < characters.length && characters[index] !== ']') index += 1;
      index += 1;
      continue;
    }
    const lower = character.toLowerCase();
    if (lower === 'y') {
      const run = countRun(characters, index, 'y');
      output += run <= 2 ? String(parts.year % 100).padStart(2, '0') : String(parts.year);
      index += run;
      continue;
    }
    if (lower === 'd') {
      const run = countRun(characters, index, 'd');
      output += run === 1 ? String(parts.day) : String(parts.day).padStart(2, '0');
      index += run;
      continue;
    }
    if (lower === 'h') {
      const run = countRun(characters, index, 'h');
      let hours = parts.hours;
      if (hasAmPm) {
        hours = hours % 12;
        if (hours === 0) hours = 12;
      }
      output += run === 1 ? String(hours) : String(hours).padStart(2, '0');
      index += run;
      continue;
    }
    if (lower === 's') {
      const run = countRun(characters, index, 's');
      output += run === 1 ? String(parts.seconds) : String(parts.seconds).padStart(2, '0');
      index += run;
      continue;
    }
    if (lower === 'm') {
      const run = countRun(characters, index, 'm');
      const isMinute = minutePending() || looksLikeMinute(characters, index, run);
      if (isMinute) {
        output += run === 1 ? String(parts.minutes) : String(parts.minutes).padStart(2, '0');
      } else if (run === 1) {
        output += String(parts.month);
      } else if (run === 2) {
        output += String(parts.month).padStart(2, '0');
      } else if (run === 3) {
        output += MONTH_NAMES[parts.month - 1] ?? '';
      } else {
        output += MONTH_NAMES[parts.month - 1] ?? '';
      }
      index += run;
      continue;
    }
    if (/am\/pm/i.test(characters.slice(index).join('').slice(0, 5))) {
      output += parts.hours < 12 ? 'AM' : 'PM';
      index += 5;
      continue;
    }
    output += character;
    index += 1;
  }
  return output;
}

function countRun(characters: readonly string[], start: number, letter: string): number {
  let run = 0;
  while (characters[start + run]?.toLowerCase() === letter) run += 1;
  return run;
}

/** `m` after an `h` or before an `s` is a minute, not a month. */
function looksLikeMinute(characters: readonly string[], start: number, run: number): boolean {
  for (let ahead = start + run; ahead < characters.length; ahead += 1) {
    const character = characters[ahead];
    if (character === undefined) break;
    if (character.toLowerCase() === 's') return true;
    if (character === '"') break;
    if (!/[a-z0-9:.]/i.test(character)) break;
  }
  return false;
}
