import { detectDelimiter, parseCsv } from '@io/csv';
import { createWorkbookBytesFromSheetData } from './blank';
import { openWorkbook, type OpenedWorkbook } from './open';

/**
 * Open a `.csv` as a one-sheet workbook (feature matrix §7 "CSV import and
 * export"; ticket 49 file drop).
 *
 * CSV is a value channel, not a workbook format (file-compat §11): fields come
 * in as numbers or text, never as formulas, exactly as the CSV round-trip
 * writes displayed values. The rows are packed into the same blank-workbook
 * template and read through the one `openWorkbook` path, so a CSV workbook is
 * modelled exactly like any other file.
 */
export function openCsvWorkbook(bytes: Uint8Array, fileName: string): OpenedWorkbook {
  const text = new TextDecoder().decode(bytes);
  const rows = parseCsv(text, { delimiter: detectDelimiter(text) });
  return openWorkbook(
    createWorkbookBytesFromSheetData(csvSheetName(fileName), csvSheetData(rows)),
    fileName,
  );
}

/** The sheet takes the file's stem, as Excel names a CSV's first sheet. */
function csvSheetName(fileName: string): string {
  const stem = fileName
    .replace(/\.[^.]*$/, '')
    .replace(/[[\]:*?/\\]/g, '')
    .trim();
  const clipped = stem.slice(0, 31);
  return clipped === '' ? 'Sheet1' : clipped;
}

const PLAIN_NUMBER = /^-?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?$/;

function csvSheetData(rows: readonly (readonly string[])[]): string {
  if (rows.length === 0) return '<sheetData/>';
  const parts: string[] = ['<sheetData>'];
  for (let row = 0; row < rows.length; row += 1) {
    const fields = rows[row] ?? [];
    const cells: string[] = [];
    for (let col = 0; col < fields.length; col += 1) {
      const field = fields[col] ?? '';
      if (field === '') continue;
      const ref = cellRef(row, col);
      if (PLAIN_NUMBER.test(field)) {
        const value = Number(field);
        if (Number.isFinite(value)) {
          cells.push(`<c r="${ref}"><v>${field.trim()}</v></c>`);
          continue;
        }
      }
      cells.push(
        `<c r="${ref}" t="inlineStr"><is><t${field === field.trim() ? '' : ' xml:space="preserve"'}>${escapeXmlText(
          field,
        )}</t></is></c>`,
      );
    }
    if (cells.length > 0) {
      parts.push(`<row r="${row + 1}">${cells.join('')}</row>`);
    }
  }
  parts.push('</sheetData>');
  return parts.join('');
}

function cellRef(row: number, col: number): string {
  let letters = '';
  let value = col;
  while (value >= 0) {
    letters = String.fromCharCode(65 + (value % 26)) + letters;
    value = Math.floor(value / 26) - 1;
  }
  return `${letters}${row + 1}`;
}

function escapeXmlText(text: string): string {
  return text.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;');
}
