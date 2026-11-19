import type { GridRange, Sheet } from '@model/types';

/**
 * CSV as a value channel (file-compat §11, F24; feature matrix §7).
 *
 * The module is deliberately small: RFC 4180-style quoting, the default
 * comma delimiter and CRLF records on the way out, and a lenient reader for
 * the comma, semicolon and tab dialects producers emit. A workbook is never
 * created or mutated here, and the exporter never reads a cell's formula:
 * a formula cell exports the displayed text its view supplies, because CSV
 * carries values, not workbook semantics.
 */

export type CsvDelimiter = ',' | ';' | '\t';
export type CsvLineEnding = '\r\n' | '\n';

export interface CsvDialect {
  readonly delimiter: CsvDelimiter;
  readonly lineEnding: CsvLineEnding;
}

/** RFC 4180: comma-delimited records ended by CRLF. */
export const DEFAULT_CSV_DIALECT: CsvDialect = Object.freeze({
  delimiter: ',',
  lineEnding: '\r\n',
});

export class CsvParseError extends Error {
  constructor(
    message: string,
    readonly row: number,
    readonly column: number,
  ) {
    super(message);
    this.name = 'CsvParseError';
  }
}

export interface ParseCsvOptions {
  /** Defaults to `DEFAULT_CSV_DIALECT.delimiter`. */
  readonly delimiter?: CsvDelimiter;
}

/**
 * Read RFC 4180-style text into rows of fields. Quoted fields keep their
 * embedded delimiters, CR/LF and doubled quotes; a UTF-8 BOM is ignored;
 * CRLF, LF and lone CR all end a record. An unterminated quote is an error
 * rather than silent truncation, because that would drop data.
 */
export function parseCsv(text: string, options: ParseCsvOptions = {}): string[][] {
  const delimiter = options.delimiter ?? DEFAULT_CSV_DIALECT.delimiter;
  const source = text.startsWith('\uFEFF') ? text.slice(1) : text;
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let inQuotes = false;
  let recordStart = 0;
  let rowIndex = 0;

  const endField = (): void => {
    row.push(field);
    field = '';
  };
  const endRecord = (end: number): void => {
    endField();
    rows.push(row);
    row = [];
    recordStart = end;
    rowIndex += 1;
  };

  let index = 0;
  while (index < source.length) {
    const character = source[index]!;
    if (inQuotes) {
      if (character === '"') {
        if (source[index + 1] === '"') {
          field += '"';
          index += 2;
        } else {
          inQuotes = false;
          index += 1;
        }
      } else {
        field += character;
        index += 1;
      }
      continue;
    }
    if (character === '"' && field === '') {
      inQuotes = true;
      index += 1;
      continue;
    }
    if (character === delimiter) {
      endField();
      index += 1;
      continue;
    }
    if (character === '\r' || character === '\n') {
      const next = character === '\r' && source[index + 1] === '\n' ? index + 2 : index + 1;
      endRecord(next);
      index = next;
      continue;
    }
    field += character;
    index += 1;
  }

  if (inQuotes) {
    throw new CsvParseError(
      'The CSV text ends inside a quoted field; the import was rejected.',
      rowIndex,
      field.length,
    );
  }
  // A final record exists only when characters followed the last record end:
  // a trailing newline does not invent an empty row.
  if (index > recordStart) {
    endField();
    rows.push(row);
  }
  return rows;
}

/**
 * Pick the delimiter a file most likely uses. The candidate that yields the
 * most fields on the first record wins; a consistent field count breaks a
 * tie, and the RFC default (comma) wins the remaining ties.
 */
export function detectDelimiter(text: string): CsvDelimiter {
  const candidates: readonly CsvDelimiter[] = [',', ';', '\t'];
  let best: {
    readonly delimiter: CsvDelimiter;
    readonly fields: number;
    readonly consistent: number;
  } = { delimiter: ',', fields: 1, consistent: 0 };
  for (const candidate of candidates) {
    let rows: string[][];
    try {
      rows = parseCsv(text, { delimiter: candidate });
    } catch {
      continue;
    }
    const first = rows[0];
    if (first === undefined || first.length <= 1) continue;
    const consistent = rows.every((record) => record.length === first.length) ? 1 : 0;
    if (
      first.length > best.fields ||
      (first.length === best.fields && consistent > best.consistent)
    ) {
      best = { delimiter: candidate, fields: first.length, consistent };
    }
  }
  return best.delimiter;
}

/** One field, quoted only when RFC 4180 requires it. */
export function formatCsvField(field: string, delimiter: CsvDelimiter = ','): string {
  if (
    field.includes(delimiter) ||
    field.includes('"') ||
    field.includes('\r') ||
    field.includes('\n')
  ) {
    return `"${field.replaceAll('"', '""')}"`;
  }
  return field;
}

/**
 * Write rows as CSV. Every record is terminated, including the last, which is
 * what Excel and RFC 4180 produce and what keeps the fixture byte-stable.
 */
export function serialiseCsv(
  rows: readonly (readonly string[])[],
  dialect: Partial<CsvDialect> = {},
): string {
  const resolved: CsvDialect = { ...DEFAULT_CSV_DIALECT, ...dialect };
  if (rows.length === 0) return '';
  const records = rows.map((row) =>
    row.map((field) => formatCsvField(field, resolved.delimiter)).join(resolved.delimiter),
  );
  return `${records.join(resolved.lineEnding)}${resolved.lineEnding}`;
}

export interface CsvSheetExportOptions {
  /** Cells to export; defaults to the sheet's used range. */
  readonly range?: GridRange;
  /**
   * The displayed text of one cell, as the grid paints it: the view supplies
   * this so a formula exports its displayed value, never its formula text.
   * `null` or `undefined` means an empty field.
   */
  readonly displayText: (row: number, col: number) => string | null | undefined;
  readonly dialect?: Partial<CsvDialect>;
}

/**
 * Export one sheet's values. The sheet is read only for its used range; every
 * field comes from `displayText`, because CSV is a value channel and a
 * formula has no meaning in it.
 */
export function exportSheetToCsv(sheet: Sheet, options: CsvSheetExportOptions): string {
  const range = options.range ?? csvUsedRange(sheet);
  if (range === null) return '';
  const rows: string[][] = [];
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    const fields: string[] = [];
    for (let col = range.startCol; col <= range.endCol; col += 1) {
      fields.push(options.displayText(row, col) ?? '');
    }
    rows.push(fields);
  }
  return serialiseCsv(rows, options.dialect ?? {});
}

/** The bounding box of the sheet's non-blank cells, or `null` for an empty sheet. */
export function csvUsedRange(sheet: Sheet): GridRange | null {
  let startRow = Number.POSITIVE_INFINITY;
  let startCol = Number.POSITIVE_INFINITY;
  let endRow = -1;
  let endCol = -1;
  for (const [key, cell] of sheet.cells) {
    if (cell.value.type === 'blank' && cell.formula === null) continue;
    const separator = key.indexOf(':');
    const row = Number(key.slice(0, separator));
    const col = Number(key.slice(separator + 1));
    if (!Number.isInteger(row) || !Number.isInteger(col)) continue;
    startRow = Math.min(startRow, row);
    startCol = Math.min(startCol, col);
    endRow = Math.max(endRow, row);
    endCol = Math.max(endCol, col);
  }
  if (endRow < 0) return null;
  return { startRow, startCol, endRow, endCol };
}
