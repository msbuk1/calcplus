import { describe, expect, it } from 'vitest';
import { DocumentSession } from '@commands/session';
import { getCell } from '@model/types';
import { numberValue } from '@model/value';
import { createEmptyWorkbook, createSheet } from '@model/workbook';

function sessionWithValue() {
  const workbook = createEmptyWorkbook();
  const sheet = createSheet('Sheet1', 'xl/worksheets/sheet1.xml');
  workbook.sheets.push(sheet);
  sheet.cells.set('0:0', {
    value: numberValue(10),
    formula: null,
    formulaDerived: false,
    savedResult: null,
    styleId: 0,
    richText: null,
  });
  return { session: new DocumentSession(workbook), workbook, sheet };
}

const target = {
  sheetId: '',
  range: { startRow: 0, startCol: 0, endRow: 0, endCol: 0 },
};

describe('M3C formatting commands', () => {
  it('toggles bold as one transaction and restores the exact prior cell on undo', () => {
    const { session, workbook, sheet } = sessionWithValue();
    const before = getCell(sheet, 0, 0);
    const args = { ...target, sheetId: sheet.id };

    const result = session.dispatch('format.bold', args, {
      activeSheetId: sheet.id,
      selection: args.range,
      activeCell: { row: 0, col: 0 },
    });
    expect(result.changes).toHaveLength(1);
    expect(result.changes[0]?.type).toBe('style.set');
    expect(getCell(sheet, 0, 0)?.styleId).toBeGreaterThan(0);
    expect(workbook.styles[getCell(sheet, 0, 0)!.styleId]?.font.bold).toBe(true);
    expect(session.transactionCount).toBe(1);

    session.dispatch('edit.undo', undefined);
    expect(getCell(sheet, 0, 0)).toEqual(before);
    expect(session.canRedo).toBe(true);

    session.dispatch('edit.redo', undefined);
    expect(workbook.styles[getCell(sheet, 0, 0)!.styleId]?.font.bold).toBe(true);
  });

  it('creates a styled record for an empty cell and removes it on undo', () => {
    const { session, sheet } = sessionWithValue();
    const range = { startRow: 1, startCol: 1, endRow: 1, endCol: 1 };
    const result = session.dispatch(
      'format.fill',
      { sheetId: sheet.id, range, pattern: 'solid', color: '#FFF2CC' },
      { activeSheetId: sheet.id, selection: range, activeCell: { row: 1, col: 1 } },
    );
    expect(result.changes).toHaveLength(1);
    expect(getCell(sheet, 1, 1)?.styleId).toBeGreaterThan(0);

    session.dispatch('edit.undo', undefined);
    expect(getCell(sheet, 1, 1)).toBeNull();

    session.dispatch('edit.redo', undefined);
    expect(getCell(sheet, 1, 1)?.styleId).toBeGreaterThan(0);
  });

  it('keeps a number-format change and its undo exact', () => {
    const { session, workbook, sheet } = sessionWithValue();
    const range = target.range;
    session.dispatch(
      'format.numberFormat',
      { sheetId: sheet.id, range, code: '#,##0.00' },
      { activeSheetId: sheet.id, selection: range, activeCell: { row: 0, col: 0 } },
    );
    const cell = getCell(sheet, 0, 0)!;
    expect(workbook.styles[cell.styleId]?.numFmtCode).toBe('#,##0.00');
    expect(cell.value).toEqual(numberValue(10));

    session.dispatch('edit.undo', undefined);
    expect(getCell(sheet, 0, 0)?.styleId).toBe(0);
  });

  it('applies a range format as one transaction', () => {
    const { session, sheet } = sessionWithValue();
    const range = { startRow: 0, startCol: 0, endRow: 0, endCol: 2 };
    session.dispatch(
      'format.italic',
      { sheetId: sheet.id, range },
      {
        activeSheetId: sheet.id,
        selection: range,
        activeCell: { row: 0, col: 0 },
      },
    );
    expect(session.transactionCount).toBe(1);
    expect(getCell(sheet, 0, 0)?.styleId).toBeGreaterThan(0);
    expect(getCell(sheet, 0, 1)?.styleId).toBe(getCell(sheet, 0, 0)?.styleId);
    expect(getCell(sheet, 0, 2)?.styleId).toBe(getCell(sheet, 0, 0)?.styleId);
  });
});
