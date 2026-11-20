import { describe, expect, it, vi } from 'vitest';
import { createCommandRegistry } from '@commands/registry';
import { DocumentSession } from '@commands/session';
import type { GridRange, Sheet } from '@model/types';
import { cloneCell, createCell, setCellSnapshot } from '@model/types';
import type { CellValue } from '@model/value';
import { numberValue, textValue } from '@model/value';
import { appendStyle, createEmptyWorkbook, createSheet } from '@model/workbook';
import { confirmCloseDocument } from '@io/confirm';

/**
 * The confirmed version-one gaps: Clear All is one undoable transaction over
 * the whole cell record, and Close asks before discarding a dirty workbook.
 */

function workbookWithSheet() {
  const workbook = createEmptyWorkbook();
  const sheet = createSheet('Sheet1', 'xl/worksheets/sheet1.xml');
  workbook.sheets.push(sheet);
  return { workbook, sheet };
}

function range(startRow: number, startCol: number, endRow: number, endCol: number): GridRange {
  return { startRow, startCol, endRow, endCol };
}

function put(sheet: Sheet, row: number, col: number, value: CellValue, styleId = 0): void {
  setCellSnapshot(sheet, row, col, createCell(value, styleId));
}

function dispatch(session: DocumentSession, sheet: Sheet, target: GridRange) {
  return session.dispatch(
    'edit.clearAll',
    { sheetId: sheet.id, range: target },
    {
      activeSheetId: sheet.id,
      selection: target,
      activeCell: { row: target.startRow, col: target.startCol },
    },
  );
}

function snapshot(sheet: Sheet) {
  return [...sheet.cells.entries()]
    .sort(([left], [right]) => left.localeCompare(right))
    .map(([address, cell]) => [address, cloneCell(cell)] as const);
}

describe('edit.clearAll', () => {
  it('clears values, formulas, rich text and formats as one transaction', () => {
    const { workbook, sheet } = workbookWithSheet();
    const styleId = appendStyle(workbook, {
      numFmtId: 0,
      fontId: 0,
      fillId: 0,
      borderId: 0,
      xfId: 0,
      rawXml: '',
    });
    put(sheet, 0, 0, textValue('keep'), styleId);
    put(sheet, 0, 1, numberValue(2));
    setCellSnapshot(sheet, 1, 0, {
      value: textValue('rich'),
      formula: null,
      formulaDerived: false,
      savedResult: null,
      styleId: 0,
      richText: { xml: '<is><r><t>rich</t></r></is>', plainText: 'rich' },
      stale: false,
    });
    setCellSnapshot(sheet, 1, 1, {
      value: numberValue(4),
      formula: 'A1&"x"',
      formulaDerived: false,
      savedResult: numberValue(4),
      styleId,
      richText: null,
      stale: false,
    });
    const session = new DocumentSession(workbook);
    const before = snapshot(sheet);
    const target = range(0, 0, 1, 1);

    const result = dispatch(session, sheet, target);
    expect(result.ui).toBeUndefined();
    expect(result.changes).toHaveLength(1);
    expect(result.changes[0]?.type).toBe('cell.set');
    expect(sheet.cells.size).toBe(0);
    expect(session.transactionCount).toBe(1);

    session.dispatch('edit.undo', {}, { activeSheetId: sheet.id });
    expect(snapshot(sheet)).toEqual(before);
  });

  it('creates no transaction when the range is already empty', () => {
    const { workbook, sheet } = workbookWithSheet();
    const session = new DocumentSession(workbook);
    const result = dispatch(session, sheet, range(0, 0, 2, 2));
    expect(result.changes).toHaveLength(0);
    expect(session.transactionCount).toBe(0);
  });

  it('is registered and enabled while a sheet is active', () => {
    const { workbook, sheet } = workbookWithSheet();
    const registry = createCommandRegistry();
    const command = registry.require('edit.clearAll');
    expect(
      command.isEnabled({ workbook, activeSheetId: sheet.id, selection: null, activeCell: null }),
    ).toBe(true);
    expect(
      command.isEnabled({
        workbook: createEmptyWorkbook(),
        activeSheetId: null,
        selection: null,
        activeCell: null,
      }),
    ).toBe(false);
  });
});

describe('confirmCloseDocument', () => {
  it('closes a clean workbook without asking', () => {
    const confirmFn = vi.fn(() => false);
    expect(confirmCloseDocument(false, confirmFn)).toBe(true);
    expect(confirmFn).not.toHaveBeenCalled();
  });

  it('prompts before a dirty workbook and honours the answer', () => {
    const yes = vi.fn((message: string) => {
      expect(message).toContain('unsaved changes');
      return true;
    });
    expect(confirmCloseDocument(true, yes)).toBe(true);
    expect(yes).toHaveBeenCalledTimes(1);

    const no = vi.fn((message: string) => {
      expect(message).toContain('unsaved changes');
      return false;
    });
    expect(confirmCloseDocument(true, no)).toBe(false);
    expect(no).toHaveBeenCalledTimes(1);
  });
});
