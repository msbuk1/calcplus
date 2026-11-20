import { buildXlsx } from '../helpers/build-xlsx';

/**
 * The M5 data-tools fixtures. They are built at test time so the filter and
 * sort walks do not depend on a binary workbook drifting from the model.
 *
 * `cityTable` is A1:B6: a header row and five city rows, one of them blank in
 * the Sales column so the checklist's `(Blanks)` entry is exercised.
 *
 * `threeKeyTable` is A1:C6 with two rows that tie on City and Dept but differ
 * in Sales, so a three-key sort is observable and a two-key sort is not.
 */

function columnLetters(index: number): string {
  let letters = '';
  let value = index + 1;
  while (value > 0) {
    const remainder = (value - 1) % 26;
    letters = String.fromCharCode(65 + remainder) + letters;
    value = Math.floor((value - 1) / 26);
  }
  return letters;
}

function tableXlsx(
  header: readonly string[],
  rows: readonly (readonly (string | number)[])[],
  startRow = 1,
): Uint8Array {
  const strings: string[] = [];
  const stringIndex = (text: string): number => {
    const existing = strings.indexOf(text);
    if (existing >= 0) return existing;
    strings.push(text);
    return strings.length - 1;
  };
  const allRows = [header, ...rows];
  const rowsXml = allRows
    .map((values, rowIndex) => {
      const cells = values
        .map((value, colIndex) => {
          const ref = `${columnLetters(colIndex)}${rowIndex + startRow}`;
          return typeof value === 'number'
            ? `<c r="${ref}"><v>${value}</v></c>`
            : `<c r="${ref}" t="s"><v>${stringIndex(value)}</v></c>`;
        })
        .join('');
      return `<row r="${rowIndex + startRow}">${cells}</row>`;
    })
    .join('');
  const lastColumn = columnLetters(header.length - 1);
  const endRow = startRow + allRows.length - 1;
  const worksheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A${startRow}:${lastColumn}${endRow}"/>
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetData>${rowsXml}</sheetData>
</worksheet>`;
  return buildXlsx(worksheetXml, { sharedStrings: strings });
}

export function cityTableXlsx(): Uint8Array {
  return tableXlsx(
    ['City', 'Sales'],
    [
      ['Bristol', 120],
      ['Bath', 80],
      ['Leeds', 200],
      ['York', 150],
      ['Cork', 60],
    ],
  );
}

export function threeKeyTableXlsx(): Uint8Array {
  return tableXlsx(
    ['Dept', 'City', 'Sales'],
    [
      ['Sales', 'Bath', 100],
      ['Ops', 'Bath', 120],
      ['Ops', 'Bath', 90],
      ['Sales', 'Bristol', 100],
      ['Ops', 'Bristol', 50],
    ],
  );
}

/** A header and two rows: one value wider than the 248px dropdown, one short. */
export function longValueTableXlsx(longValue: string): Uint8Array {
  return tableXlsx(['Value'], [[longValue], ['Short']]);
}

/** A header and `count` distinct values, for the checklist's virtual window. */
export function manyValuesTableXlsx(count: number): Uint8Array {
  return tableXlsx(
    ['Value'],
    Array.from({ length: count }, (_, index) => [`Value ${String(index + 1).padStart(4, '0')}`]),
  );
}

/** A table whose header can sit deep in the sheet, so the grid scrolls to it. */
export function deepHeaderTableXlsx(headerRow: number): Uint8Array {
  return tableXlsx(['Value'], [['Bath'], ['Bristol'], ['Cork'], ['Leeds'], ['York']], headerRow);
}

/** One header row spanning A to Z, for opening a dropdown at the right edge. */
export function wideTableXlsx(): Uint8Array {
  const header = Array.from({ length: 26 }, (_, index) => `Column ${index + 1}`);
  const values = Array.from({ length: 26 }, (_, index) => index + 1);
  return tableXlsx(header, [values]);
}
