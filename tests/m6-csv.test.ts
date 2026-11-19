import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import {
  CsvParseError,
  DEFAULT_CSV_DIALECT,
  csvUsedRange,
  detectDelimiter,
  exportSheetToCsv,
  formatCsvField,
  parseCsv,
  serialiseCsv,
} from '@io/csv';
import { createEmptyWorkbook, createSheet } from '@model/workbook';
import { BLANK, formatCellValue, numberValue, textValue } from '@model/value';
import type { Cell } from '@model/types';
import { createCell } from '@model/types';

/**
 * M6 lane 6 — F24 (file-compat §9, §11: "CSV is a value channel, not a
 * workbook format").
 *
 * The tests pin RFC 4180 quoting, the comma default and delimiter detection,
 * embedded newlines and quotes, non-ASCII text, and the export rule that a
 * formula cell contributes the displayed text its view supplies, never its
 * formula text.
 */

const fixture = readFileSync(new URL('../fixtures/compat/csv-roundtrip.csv', import.meta.url));
const FIXTURE_TEXT = fixture.toString('utf8');

function fixtureRows(): string[][] {
  return [
    ['Product', 'Region', 'Units', 'Unit price', 'Total', 'Note'],
    ['Widget', 'North', '3', '12.5', '37.50', 'Gift wrap, blue'],
    ['Gadget', 'South', '2', '1234.5', '2,469.00', 'Quoted "as-is"'],
    ['Thing', 'East', '0', '0.5', '0.00', 'Multi-line\r\nnote with a comma, and "quotes"'],
    ['Café', 'West', '1', '€2.50', '€2.50', 'Non-ASCII snowman ☃'],
  ];
}

describe('CSV parsing (RFC 4180)', () => {
  it('reads the F24 fixture into the authored rows', () => {
    expect(parseCsv(FIXTURE_TEXT)).toEqual(fixtureRows());
  });

  it('keeps an embedded delimiter, quote and newline inside a quoted field', () => {
    expect(parseCsv('a,"b,c","d""e","f\r\ng"\r\n')).toEqual([['a', 'b,c', 'd"e', 'f\r\ng']]);
  });

  it('accepts CRLF, LF and lone CR record endings', () => {
    expect(parseCsv('a,b\r\nc,d')).toEqual([
      ['a', 'b'],
      ['c', 'd'],
    ]);
    expect(parseCsv('a,b\nc,d\n')).toEqual([
      ['a', 'b'],
      ['c', 'd'],
    ]);
    expect(parseCsv('a,b\rc,d')).toEqual([
      ['a', 'b'],
      ['c', 'd'],
    ]);
  });

  it('ignores a UTF-8 BOM', () => {
    expect(parseCsv('\uFEFFa,b\r\n')).toEqual([['a', 'b']]);
  });

  it('preserves empty fields, a trailing delimiter and a quoted empty field', () => {
    expect(parseCsv('a,,c\r\n')).toEqual([['a', '', 'c']]);
    expect(parseCsv('a,b,\r\n')).toEqual([['a', 'b', '']]);
    expect(parseCsv('a,"",c\r\n')).toEqual([['a', '', 'c']]);
  });

  it('returns no rows for empty text, and no invented row for a trailing newline', () => {
    expect(parseCsv('')).toEqual([]);
    expect(parseCsv('\r\n')).toEqual([['']]);
    expect(parseCsv('a\r\n')).toEqual([['a']]);
  });

  it('rejects an unterminated quoted field instead of truncating it', () => {
    expect(() => parseCsv('a,"b,c\r\n')).toThrow(CsvParseError);
  });

  it('handles non-ASCII text unchanged', () => {
    expect(parseCsv('café,☃,"€1,00"\r\n', { delimiter: ',' })).toEqual([['café', '☃', '€1,00']]);
  });

  it('parses semicolon and tab dialects', () => {
    expect(parseCsv('a;b;c\r\n', { delimiter: ';' })).toEqual([['a', 'b', 'c']]);
    expect(parseCsv('a\tb\tc\r\n', { delimiter: '\t' })).toEqual([['a', 'b', 'c']]);
  });
});

describe('delimiter detection', () => {
  it('defaults to a comma', () => {
    expect(DEFAULT_CSV_DIALECT.delimiter).toBe(',');
    expect(DEFAULT_CSV_DIALECT.lineEnding).toBe('\r\n');
    expect(detectDelimiter('plain text without delimiters')).toBe(',');
  });

  it('recognises semicolon and tab exports', () => {
    expect(detectDelimiter('a;b;c\r\n1;2;3\r\n')).toBe(';');
    expect(detectDelimiter('a\tb\tc\r\n1\t2\t3\r\n')).toBe('\t');
  });

  it('ignores delimiters inside quotes', () => {
    expect(detectDelimiter('"a;b";c\r\n"d;e";f\r\n')).toBe(';');
    expect(detectDelimiter('"a,b",c\r\n"d,e",f\r\n')).toBe(',');
  });
});

describe('CSV writing (RFC 4180)', () => {
  it('quotes only the fields that need it', () => {
    expect(formatCsvField('plain')).toBe('plain');
    expect(formatCsvField('a,b')).toBe('"a,b"');
    expect(formatCsvField('a"b')).toBe('"a""b"');
    expect(formatCsvField('a\r\nb')).toBe('"a\r\nb"');
    expect(formatCsvField('a;b')).toBe('a;b');
    expect(formatCsvField('a;b', ';')).toBe('"a;b"');
  });

  it('writes CRLF by default and terminates the last record', () => {
    expect(
      serialiseCsv([
        ['a', 'b'],
        ['1', '2'],
      ]),
    ).toBe('a,b\r\n1,2\r\n');
    expect(serialiseCsv([])).toBe('');
  });

  it('round-trips the F24 fixture byte-for-byte', () => {
    const rows = parseCsv(FIXTURE_TEXT);
    expect(serialiseCsv(rows)).toBe(FIXTURE_TEXT);
  });

  it('round-trips quotes, newlines and non-ASCII through parse and write', () => {
    const rows = [['quoted "field"', 'multi\r\nline', '☃', 'semi;colon', 'comma,field']];
    const text = serialiseCsv(rows);
    expect(text).toBe('"quoted ""field""","multi\r\nline",☃,semi;colon,"comma,field"\r\n');
    expect(parseCsv(text)).toEqual(rows);
  });
});

describe('sheet export is a value channel', () => {
  it('exports displayed text, so a formula cell contributes its result and never its text', () => {
    const workbook = createEmptyWorkbook();
    const sheet = createSheet('Sheet1', 'xl/worksheets/sheet1.xml');
    const formulaCell: Cell = {
      ...createCell(numberValue(37.5)),
      formula: 'A1*3',
      savedResult: numberValue(37.5),
    };
    sheet.cells.set('0:0', formulaCell);
    sheet.cells.set('0:1', createCell(textValue('Gift wrap, blue')));
    workbook.sheets.push(sheet);

    const displayed = (row: number, col: number): string | null => {
      const cell: Cell | undefined = sheet.cells.get(`${row}:${col}`);
      if (cell === undefined) return null;
      // The view formats the cached result; the formula text is never used.
      return cell.value.type === 'number'
        ? cell.value.value.toFixed(2)
        : formatCellValue(cell.value);
    };
    const csv = exportSheetToCsv(sheet, { displayText: displayed });
    expect(csv).toBe('37.50,"Gift wrap, blue"\r\n');
    expect(csv).not.toContain('A1*3');
  });

  it('uses the used range, honours an explicit range and writes an empty sheet as empty text', () => {
    const workbook = createEmptyWorkbook();
    const sheet = createSheet('Sheet1', 'xl/worksheets/sheet1.xml');
    sheet.cells.set('1:1', createCell(textValue('b2')));
    sheet.cells.set('2:2', createCell(numberValue(3)));
    workbook.sheets.push(sheet);
    expect(csvUsedRange(sheet)).toEqual({ startRow: 1, startCol: 1, endRow: 2, endCol: 2 });
    const display = (row: number, col: number): string => {
      const cell = sheet.cells.get(`${row}:${col}`);
      return cell === undefined ? '' : formatCellValue(cell.value);
    };
    expect(exportSheetToCsv(sheet, { displayText: display })).toBe('b2,\r\n,3\r\n');
    expect(
      exportSheetToCsv(sheet, {
        displayText: display,
        range: { startRow: 1, startCol: 1, endRow: 1, endCol: 2 },
      }),
    ).toBe('b2,\r\n');

    const empty = createSheet('Empty', 'xl/worksheets/sheet2.xml');
    expect(csvUsedRange(empty)).toBeNull();
    expect(exportSheetToCsv(empty, { displayText: () => '' })).toBe('');
  });

  it('writes the chosen dialect', () => {
    const workbook = createEmptyWorkbook();
    const sheet = createSheet('Sheet1', 'xl/worksheets/sheet1.xml');
    sheet.cells.set('0:0', createCell(textValue('a')));
    sheet.cells.set('0:1', createCell(textValue('b;c')));
    workbook.sheets.push(sheet);
    const display = (row: number, col: number): string =>
      formatCellValue(sheet.cells.get(`${row}:${col}`)?.value ?? BLANK);
    expect(
      exportSheetToCsv(sheet, {
        displayText: display,
        dialect: { delimiter: ';', lineEnding: '\n' },
      }),
    ).toBe('a;"b;c"\n');
  });
});
