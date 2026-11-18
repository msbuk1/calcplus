import { describe, expect, it } from 'vitest';
import { createCommandRegistry } from '@commands/registry';
import { DocumentSession } from '@commands/session';
import { adjustDecimalPlaces } from '@commands/format';
import { getCell, setCellSnapshot, createCell } from '@model/types';
import { numberValue, textValue } from '@model/value';
import { createEmptyWorkbook, createSheet } from '@model/workbook';

/**
 * The M5 chrome-parity commands: Wrap Text and the two decimal steps. Each is
 * one transaction (command-system §4) over the existing style machinery, so the
 * before/after style ids are the whole undo record.
 */

function sessionWithValue() {
  const workbook = createEmptyWorkbook();
  const sheet = createSheet('Sheet1', 'xl/worksheets/sheet1.xml');
  workbook.sheets.push(sheet);
  setCellSnapshot(sheet, 0, 0, createCell(numberValue(10)));
  return { session: new DocumentSession(workbook), workbook, sheet };
}

const range = { startRow: 0, startCol: 0, endRow: 0, endCol: 0 };

function dispatch(session: DocumentSession, sheetId: string, commandId: string) {
  return session.dispatch(
    commandId,
    { sheetId, range },
    { activeSheetId: sheetId, selection: range, activeCell: { row: 0, col: 0 } },
  );
}

describe('format.wrapText', () => {
  it('toggles wrap text as one transaction and undoes exactly', () => {
    const { session, workbook, sheet } = sessionWithValue();
    const before = getCell(sheet, 0, 0);

    const on = dispatch(session, sheet.id, 'format.wrapText');
    expect(on.changes).toHaveLength(1);
    expect(on.changes[0]?.type).toBe('style.set');
    expect(session.transactionCount).toBe(1);
    const wrapped = workbook.styles[getCell(sheet, 0, 0)!.styleId]!;
    expect(wrapped.alignment.wrapText).toBe(true);

    const off = dispatch(session, sheet.id, 'format.wrapText');
    expect(session.transactionCount).toBe(2);
    expect(off.changes).toHaveLength(1);
    expect(workbook.styles[getCell(sheet, 0, 0)!.styleId]?.alignment.wrapText).toBe(false);

    session.dispatch('edit.undo', undefined);
    session.dispatch('edit.undo', undefined);
    expect(getCell(sheet, 0, 0)).toEqual(before);
  });

  it('is checked from the active cell and shared with the alignment command', () => {
    const { session, sheet } = sessionWithValue();
    const registry = createCommandRegistry();
    const command = registry.require('format.wrapText');
    const context = {
      workbook: session.workbook,
      activeSheetId: sheet.id,
      selection: range,
      activeCell: { row: 0, col: 0 },
    };

    expect(command.isChecked?.(context)).toBe(false);
    dispatch(session, sheet.id, 'format.wrapText');
    expect(command.isChecked?.(context)).toBe(true);
    // The same style bit the Format Cells dialog writes.
    expect(command.id).toBe('format.wrapText');
    expect(registry.require('format.alignment').id).toBe('format.alignment');
  });
});

describe('format.increaseDecimal / format.decreaseDecimal', () => {
  it('steps the number format once per dispatch', () => {
    const { session, workbook, sheet } = sessionWithValue();

    dispatch(session, sheet.id, 'format.increaseDecimal');
    expect(workbook.styles[getCell(sheet, 0, 0)!.styleId]?.numFmtCode).toBe('0.00');
    expect(session.transactionCount).toBe(1);

    dispatch(session, sheet.id, 'format.increaseDecimal');
    expect(workbook.styles[getCell(sheet, 0, 0)!.styleId]?.numFmtCode).toBe('0.000');

    dispatch(session, sheet.id, 'format.decreaseDecimal');
    expect(workbook.styles[getCell(sheet, 0, 0)!.styleId]?.numFmtCode).toBe('0.00');
    expect(session.transactionCount).toBe(3);

    // The value is untouched: only the format changed.
    expect(getCell(sheet, 0, 0)?.value).toEqual(numberValue(10));
  });

  it('restores the prior format on undo', () => {
    const { session, workbook, sheet } = sessionWithValue();
    dispatch(session, sheet.id, 'format.increaseDecimal');
    const styleId = getCell(sheet, 0, 0)!.styleId;
    expect(workbook.styles[styleId]?.numFmtCode).toBe('0.00');

    session.dispatch('edit.undo', undefined);
    expect(workbook.styles[getCell(sheet, 0, 0)!.styleId]?.numFmtCode ?? null).toBeNull();
  });

  it('adjusts a patterned code after its decimal point', () => {
    expect(adjustDecimalPlaces('#,##0.00', 1)).toBe('#,##0.000');
    expect(adjustDecimalPlaces('#,##0.00', -1)).toBe('#,##0.0');
    expect(adjustDecimalPlaces('0.0%', 1)).toBe('0.00%');
    expect(adjustDecimalPlaces('0.0%', -1)).toBe('0%');
    expect(adjustDecimalPlaces('0', 1)).toBe('0.0');
    expect(adjustDecimalPlaces('0', -1)).toBeNull();
    expect(adjustDecimalPlaces(null, 1)).toBe('0.00');
    expect(adjustDecimalPlaces('General', -1)).toBe('0');
    expect(adjustDecimalPlaces('"$"#,##0.00', 1)).toBe('"$"#,##0.000');
    // A second section (negatives) is preserved untouched.
    expect(adjustDecimalPlaces('#,##0.00;(#,##0.00)', 1)).toBe('#,##0.000;(#,##0.00)');
  });
});

describe('data.sort quick keys from the Sort & Filter menu', () => {
  it('sorts by the active column without the map naming it', () => {
    const workbook = createEmptyWorkbook();
    const sheet = createSheet('Sheet1', 'xl/worksheets/sheet1.xml');
    workbook.sheets.push(sheet);
    setCellSnapshot(sheet, 0, 0, createCell(textValue('Name')));
    setCellSnapshot(sheet, 0, 1, createCell(textValue('Score')));
    setCellSnapshot(sheet, 1, 0, createCell(textValue('Alice')));
    setCellSnapshot(sheet, 1, 1, createCell(numberValue(30)));
    setCellSnapshot(sheet, 2, 0, createCell(textValue('Bob')));
    setCellSnapshot(sheet, 2, 1, createCell(numberValue(10)));
    setCellSnapshot(sheet, 3, 0, createCell(textValue('Carol')));
    setCellSnapshot(sheet, 3, 1, createCell(numberValue(20)));
    const session = new DocumentSession(workbook);
    const target = { startRow: 0, startCol: 0, endRow: 3, endCol: 1 };

    const result = session.dispatch(
      'data.sort',
      { range: target, keys: [{ direction: 'ascending' }], hasHeaders: null },
      { activeSheetId: sheet.id, selection: target, activeCell: { row: 1, col: 1 } },
    );

    expect(result.sort?.state).toBe('sorted');
    // Ascending by the active column (Score): 10, 20, 30 with their names.
    expect(getCell(sheet, 1, 0)?.value).toEqual(textValue('Bob'));
    expect(getCell(sheet, 2, 0)?.value).toEqual(textValue('Carol'));
    expect(getCell(sheet, 3, 0)?.value).toEqual(textValue('Alice'));
    expect(session.transactionCount).toBe(1);
  });
});
