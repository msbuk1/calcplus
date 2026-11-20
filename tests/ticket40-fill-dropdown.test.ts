import { describe, expect, it } from 'vitest';
import { DocumentSession } from '@commands/session';
import { ribbonControlValues } from '@chrome/ribbon-readouts';
import type { GridRange } from '@model/types';
import { createCell, setCellSnapshot } from '@model/types';
import { textValue } from '@model/value';
import { createEmptyWorkbook, createSheet } from '@model/workbook';
import { findOrAppendStyle, patchedStyleInput } from '@model/style-table';
import { createSheetProjection } from '@view/projection';

/**
 * The ribbon Fill drop-down's Current swatch (ticket 40). The swatch previews
 * the active cell — the selection's anchor (ticket 31) — and a merged cell is
 * one cell, so any cell inside a merge previews the anchor's fill.
 *
 * `ribbonControlValues` is the one source for every ribbon style readout, so
 * these tests cover the other readouts' consistency at the same time.
 */

const single = (row: number, col: number): GridRange => ({
  startRow: row,
  startCol: col,
  endRow: row,
  endCol: col,
});

function cellStyle(
  workbook: ReturnType<typeof createEmptyWorkbook>,
  color: string,
  size = 11,
  numberFormat: string | null = null,
): number {
  const base = workbook.styles[0]!;
  return findOrAppendStyle(
    workbook,
    patchedStyleInput(base, {
      fill: { pattern: 'solid', color },
      font: { size },
      ...(numberFormat === null ? {} : { numFmtCode: numberFormat }),
    }),
  );
}

/** A sheet with a red A1, a merged D1:E1 green range, and a plain C3. */
function workbookWithFills() {
  const workbook = createEmptyWorkbook();
  const sheet = createSheet('Sheet1', 'xl/worksheets/sheet1.xml');
  workbook.sheets.push(sheet);
  const red = cellStyle(workbook, '#FF0000', 14, '0.00');
  const green = cellStyle(workbook, '#00B050', 18);
  setCellSnapshot(sheet, 0, 0, createCell(textValue('red'), red));
  setCellSnapshot(sheet, 0, 3, createCell(textValue('merged'), green));
  sheet.merges.push({
    range: { startRow: 0, startCol: 3, endRow: 0, endCol: 4 },
  });
  return { workbook, sheet, red, green };
}

function readout(
  workbook: ReturnType<typeof createEmptyWorkbook>,
  sheetId: string,
  row: number,
  col: number,
  version = 1,
): Readonly<Record<string, string>> {
  const projection = createSheetProjection(workbook, sheetId, { version });
  return ribbonControlValues(projection, { row, col });
}

describe('Fill drop-down Current swatch source', () => {
  it('previews a single cell that has a fill', () => {
    const { workbook, sheet } = workbookWithFills();
    expect(readout(workbook, sheet.id, 0, 0)).toMatchObject({
      'fill-color': '#FF0000',
      // The other readouts preview the same active cell.
      'font-size': '14',
      'number-format': '0.00',
    });
  });

  it('previews No Fill for a cell without a fill', () => {
    const { workbook, sheet } = workbookWithFills();
    const values = readout(workbook, sheet.id, 2, 2);
    expect(values['fill-color']).toBe('');
    expect(values['font-family']).toBe('Calibri');
    expect(values['number-format']).toBe('General');
  });

  it('previews the anchor fill of a merged cell from the anchor', () => {
    const { workbook, sheet } = workbookWithFills();
    expect(readout(workbook, sheet.id, 0, 3)['fill-color']).toBe('#00B050');
  });

  it('reads the merge anchor when the active cell is inside the merge', () => {
    const { workbook, sheet } = workbookWithFills();
    // E1 is covered by the D1:E1 merge and has no stored cell or style.
    expect(readout(workbook, sheet.id, 0, 4)).toMatchObject({
      'fill-color': '#00B050',
      'font-size': '18',
    });
  });

  it('ignores a covered cell’s hidden style in favour of the merge anchor', () => {
    const { workbook, sheet, red } = workbookWithFills();
    // A covered cell may keep a style of its own (a merge hides formatting, it
    // does not erase it); the merged cell still previews its anchor.
    setCellSnapshot(sheet, 0, 4, createCell(textValue('hidden'), red));
    expect(readout(workbook, sheet.id, 0, 4)['fill-color']).toBe('#00B050');
  });

  it('previews No Fill for a merged range whose anchor has none', () => {
    const { workbook, sheet } = workbookWithFills();
    sheet.merges.push({ range: { startRow: 4, startCol: 0, endRow: 5, endCol: 1 } });
    expect(readout(workbook, sheet.id, 5, 1)['fill-color']).toBe('');
  });

  it('updates after a fill command without reopening the file', () => {
    const { workbook, sheet } = workbookWithFills();
    const session = new DocumentSession(workbook);
    const context = {
      activeSheetId: sheet.id,
      selection: single(2, 2),
      activeCell: { row: 2, col: 2 },
    };
    expect(
      ribbonControlValues(createSheetProjection(workbook, sheet.id, { version: session.version }), {
        row: 2,
        col: 2,
      })['fill-color'],
    ).toBe('');

    session.dispatch(
      'format.fill',
      { sheetId: sheet.id, range: single(2, 2), pattern: 'solid', color: '#FFC000' },
      context,
    );
    expect(
      createSheetProjection(workbook, sheet.id, { version: session.version }).cellAt(2, 2)?.style
        .fill.color,
    ).toBe('#FFC000');
    expect(readout(workbook, sheet.id, 2, 2, session.version)['fill-color']).toBe('#FFC000');

    // No Fill clears the preview again, on the same path.
    session.dispatch(
      'format.fill',
      { sheetId: sheet.id, range: single(2, 2), pattern: 'none', color: null },
      context,
    );
    expect(readout(workbook, sheet.id, 2, 2, session.version)['fill-color']).toBe('');
  });

  it('updates a merged cell’s preview after a fill over the merge', () => {
    const { workbook, sheet } = workbookWithFills();
    const session = new DocumentSession(workbook);
    const range = { startRow: 4, startCol: 0, endRow: 5, endCol: 1 };
    sheet.merges.push({ range });
    session.dispatch(
      'format.fill',
      { sheetId: sheet.id, range, pattern: 'solid', color: '#1F4E78' },
      { activeSheetId: sheet.id, selection: range, activeCell: { row: 4, col: 0 } },
    );
    // The active cell inside the merge (not the anchor) still previews it.
    expect(readout(workbook, sheet.id, 5, 1, session.version)['fill-color']).toBe('#1F4E78');
  });

  it('has no readouts before a workbook is open', () => {
    const values = ribbonControlValues(null, { row: 0, col: 0 });
    expect(values).toEqual({
      'font-family': 'Calibri',
      'font-size': '11',
      'number-format': 'General',
      'font-color': '',
      'fill-color': '',
    });
  });
});
