import { describe, expect, it } from 'vitest';
import { createCommandRegistry } from '@commands/registry';
import { DocumentSession } from '@commands/session';
import { translateFormulaForPaste } from '@model/clipboard';
import { createEmptyWorkbook, createSheet } from '@model/workbook';
import type { Cell, GridRange, Sheet } from '@model/types';
import { getCell, setCellSnapshot } from '@model/types';
import { BLANK, numberValue, textValue } from '@model/value';

function setup(): { session: DocumentSession; sheet: Sheet } {
  const workbook = createEmptyWorkbook();
  const sheet = createSheet('Sheet1', 'xl/worksheets/sheet1.xml');
  workbook.sheets.push(sheet);
  return { session: new DocumentSession(workbook), sheet };
}

function put(
  sheet: Sheet,
  row: number,
  col: number,
  value: Cell['value'],
  formula: string | null = null,
  styleId = 0,
): void {
  setCellSnapshot(sheet, row, col, {
    value,
    formula,
    formulaDerived: false,
    savedResult: null,
    styleId,
    richText: null,
    stale: formula !== null,
  });
}

function clipboardFor(session: DocumentSession, sheetId: string, range: GridRange) {
  const result = session.dispatch('edit.copy', { sheetId, range });
  expect(result.clipboard).toBeDefined();
  return result.clipboard!;
}

function singleCell(row: number, col: number): GridRange {
  return { startRow: row, startCol: col, endRow: row, endCol: col };
}

describe('clipboard commands', () => {
  it('registers the edit.copy, edit.cut and edit.paste ids', () => {
    const registry = createCommandRegistry();
    expect(registry.require('edit.copy').keybindings).toContain('Ctrl+C');
    expect(registry.require('edit.cut').keybindings).toContain('Ctrl+X');
    expect(registry.require('edit.paste').keybindings).toContain('Ctrl+V');
  });

  it('pastes values, styles and relative formulas as one undoable transaction', () => {
    const { session, sheet } = setup();
    put(sheet, 0, 0, numberValue(10), null, 3);
    put(sheet, 1, 0, BLANK, 'A1*2', 4);
    const snapshot = clipboardFor(session, sheet.id, {
      startRow: 0,
      startCol: 0,
      endRow: 1,
      endCol: 0,
    });
    expect(session.transactionCount).toBe(0);

    const result = session.dispatch('edit.paste', {
      sheetId: sheet.id,
      target: singleCell(2, 2),
      snapshot,
    });
    expect(result.changes.length).toBeGreaterThan(0);
    expect(session.transactionCount).toBe(1);
    expect(getCell(sheet, 2, 2)?.value).toEqual(numberValue(10));
    expect(getCell(sheet, 2, 2)?.styleId).toBe(3);
    expect(getCell(sheet, 3, 2)?.formula).toBe('C3*2');
    expect(getCell(sheet, 3, 2)?.styleId).toBe(4);
    expect(result.selectionAfter).toEqual({
      startRow: 2,
      startCol: 2,
      endRow: 3,
      endCol: 2,
    });

    session.dispatch('edit.undo', undefined);
    expect(session.transactionCount).toBe(0);
    expect(getCell(sheet, 2, 2)).toBeNull();
    expect(getCell(sheet, 3, 2)).toBeNull();
    expect(getCell(sheet, 1, 0)?.formula).toBe('A1*2');
  });

  it('fills a selected range from a copied single value with its style', () => {
    const { session, sheet } = setup();
    put(sheet, 0, 0, numberValue(7), null, 3);
    const snapshot = clipboardFor(session, sheet.id, singleCell(0, 0));
    const target = { startRow: 1, startCol: 1, endRow: 3, endCol: 2 };

    const result = session.dispatch('edit.paste', {
      sheetId: sheet.id,
      target,
      snapshot,
    });
    expect(session.transactionCount).toBe(1);
    expect(result.selectionAfter).toEqual(target);
    for (let row = target.startRow; row <= target.endRow; row += 1) {
      for (let col = target.startCol; col <= target.endCol; col += 1) {
        expect(getCell(sheet, row, col)?.value).toEqual(numberValue(7));
        expect(getCell(sheet, row, col)?.styleId).toBe(3);
      }
    }
    expect(getCell(sheet, 4, 1)).toBeNull();

    session.dispatch('edit.undo', undefined);
    expect(session.transactionCount).toBe(0);
    expect(getCell(sheet, 1, 1)).toBeNull();
    expect(getCell(sheet, 3, 2)).toBeNull();
    expect(getCell(sheet, 0, 0)?.value).toEqual(numberValue(7));
  });

  it('fills a selected range from a copied single formula with per-cell offsets', () => {
    const { session, sheet } = setup();
    put(sheet, 0, 0, BLANK, 'B1+$B$1', 5);
    const snapshot = clipboardFor(session, sheet.id, singleCell(0, 0));

    session.dispatch('edit.paste', {
      sheetId: sheet.id,
      target: { startRow: 1, startCol: 0, endRow: 3, endCol: 1 },
      snapshot,
    });
    expect(session.transactionCount).toBe(1);
    expect(getCell(sheet, 1, 0)?.formula).toBe('B2+$B$1');
    expect(getCell(sheet, 1, 1)?.formula).toBe('C2+$B$1');
    expect(getCell(sheet, 3, 0)?.formula).toBe('B4+$B$1');
    expect(getCell(sheet, 3, 1)?.formula).toBe('C4+$B$1');
    expect(getCell(sheet, 3, 1)?.styleId).toBe(5);
  });

  it('tiles a copied column across a wider target', () => {
    const { session, sheet } = setup();
    put(sheet, 0, 0, numberValue(1));
    put(sheet, 1, 0, numberValue(2));
    const snapshot = clipboardFor(session, sheet.id, {
      startRow: 0,
      startCol: 0,
      endRow: 1,
      endCol: 0,
    });

    const result = session.dispatch('edit.paste', {
      sheetId: sheet.id,
      target: { startRow: 2, startCol: 2, endRow: 3, endCol: 4 },
      snapshot,
    });
    expect(result.selectionAfter).toEqual({
      startRow: 2,
      startCol: 2,
      endRow: 3,
      endCol: 4,
    });
    for (let col = 2; col <= 4; col += 1) {
      expect(getCell(sheet, 2, col)?.value).toEqual(numberValue(1));
      expect(getCell(sheet, 3, col)?.value).toEqual(numberValue(2));
    }
    expect(getCell(sheet, 2, 5)).toBeNull();
  });

  it('repeats a copied row down a taller single-column selection', () => {
    const { session, sheet } = setup();
    for (let col = 0; col < 8; col += 1) put(sheet, 0, col, numberValue(col + 1));
    const snapshot = clipboardFor(session, sheet.id, {
      startRow: 0,
      startCol: 0,
      endRow: 0,
      endCol: 7,
    });

    const result = session.dispatch('edit.paste', {
      sheetId: sheet.id,
      target: { startRow: 1, startCol: 0, endRow: 9, endCol: 0 },
      snapshot,
    });
    expect(session.transactionCount).toBe(1);
    expect(result.selectionAfter).toEqual({
      startRow: 1,
      startCol: 0,
      endRow: 9,
      endCol: 7,
    });
    for (let row = 1; row <= 9; row += 1) {
      for (let col = 0; col < 8; col += 1) {
        expect(getCell(sheet, row, col)?.value).toEqual(numberValue(col + 1));
      }
    }
    expect(getCell(sheet, 1, 8)).toBeNull();

    session.dispatch('edit.undo', undefined);
    expect(session.transactionCount).toBe(0);
    expect(getCell(sheet, 1, 0)).toBeNull();
    expect(getCell(sheet, 9, 7)).toBeNull();
    expect(getCell(sheet, 0, 7)?.value).toEqual(numberValue(8));
  });

  it('translates formulas from the corresponding source cell while repeating a row', () => {
    const { session, sheet } = setup();
    put(sheet, 0, 0, numberValue(1));
    put(sheet, 0, 1, BLANK, 'A1+$B$1');
    const snapshot = clipboardFor(session, sheet.id, {
      startRow: 0,
      startCol: 0,
      endRow: 0,
      endCol: 1,
    });

    session.dispatch('edit.paste', {
      sheetId: sheet.id,
      target: { startRow: 1, startCol: 0, endRow: 3, endCol: 0 },
      snapshot,
    });
    expect(getCell(sheet, 1, 1)?.formula).toBe('A2+$B$1');
    expect(getCell(sheet, 2, 1)?.formula).toBe('A3+$B$1');
    expect(getCell(sheet, 3, 1)?.formula).toBe('A4+$B$1');
    expect(getCell(sheet, 3, 0)?.value).toEqual(numberValue(1));
  });

  it('tiles both directions when neither target dimension divides evenly', () => {
    const { session, sheet } = setup();
    put(sheet, 0, 0, numberValue(1));
    put(sheet, 0, 1, numberValue(2));
    put(sheet, 1, 0, numberValue(3));
    put(sheet, 1, 1, numberValue(4));
    const snapshot = clipboardFor(session, sheet.id, {
      startRow: 0,
      startCol: 0,
      endRow: 1,
      endCol: 1,
    });

    session.dispatch('edit.paste', {
      sheetId: sheet.id,
      target: { startRow: 2, startCol: 2, endRow: 4, endCol: 4 },
      snapshot,
    });
    expect(getCell(sheet, 2, 2)?.value).toEqual(numberValue(1));
    expect(getCell(sheet, 2, 3)?.value).toEqual(numberValue(2));
    expect(getCell(sheet, 2, 4)?.value).toEqual(numberValue(1));
    expect(getCell(sheet, 3, 2)?.value).toEqual(numberValue(3));
    expect(getCell(sheet, 3, 3)?.value).toEqual(numberValue(4));
    expect(getCell(sheet, 3, 4)?.value).toEqual(numberValue(3));
    expect(getCell(sheet, 4, 2)?.value).toEqual(numberValue(1));
    expect(getCell(sheet, 4, 3)?.value).toEqual(numberValue(2));
    expect(getCell(sheet, 4, 4)?.value).toEqual(numberValue(1));
  });

  it('refuses a tiled paste beyond the clipboard budget', () => {
    const { session, sheet } = setup();
    for (let col = 0; col < 8; col += 1) put(sheet, 0, col, numberValue(col + 1));
    const snapshot = clipboardFor(session, sheet.id, {
      startRow: 0,
      startCol: 0,
      endRow: 0,
      endCol: 7,
    });

    const result = session.dispatch('edit.paste', {
      sheetId: sheet.id,
      target: { startRow: 1, startCol: 0, endRow: 100_001, endCol: 0 },
      snapshot,
    });
    expect(session.transactionCount).toBe(0);
    expect(result.changes).toHaveLength(0);
    expect(result.ui?.message).toContain('100,000 cells');
    expect(getCell(sheet, 1, 0)).toBeNull();
  });

  it('cuts as a move: source cells clear and the moved formula text is not offset', () => {
    const { session, sheet } = setup();
    put(sheet, 0, 0, numberValue(5));
    put(sheet, 1, 0, BLANK, 'A1+1');
    const cut = session.dispatch('edit.cut', {
      sheetId: sheet.id,
      range: { startRow: 0, startCol: 0, endRow: 1, endCol: 0 },
    });
    expect(cut.clipboard?.mode).toBe('cut');
    expect(session.transactionCount).toBe(0);

    session.dispatch('edit.paste', {
      sheetId: sheet.id,
      target: singleCell(0, 2),
      snapshot: cut.clipboard!,
    });
    expect(session.transactionCount).toBe(1);
    expect(getCell(sheet, 0, 0)).toBeNull();
    expect(getCell(sheet, 1, 0)).toBeNull();
    expect(getCell(sheet, 0, 2)?.value).toEqual(numberValue(5));
    expect(getCell(sheet, 1, 2)?.formula).toBe('A1+1');

    session.dispatch('edit.undo', undefined);
    expect(getCell(sheet, 0, 0)?.value).toEqual(numberValue(5));
    expect(getCell(sheet, 1, 0)?.formula).toBe('A1+1');
    expect(getCell(sheet, 1, 2)).toBeNull();
  });

  it('clears the source sheet when a cut is pasted on another sheet', () => {
    const { session, sheet } = setup();
    const second = createSheet('Sheet2', 'xl/worksheets/sheet2.xml');
    session.workbook.sheets.push(second);
    put(sheet, 0, 0, textValue('moved'));
    const cut = session.dispatch('edit.cut', {
      sheetId: sheet.id,
      range: singleCell(0, 0),
    });

    session.dispatch('edit.paste', {
      sheetId: second.id,
      target: singleCell(0, 0),
      snapshot: cut.clipboard!,
    });
    expect(getCell(sheet, 0, 0)).toBeNull();
    expect(getCell(second, 0, 0)?.value).toEqual(textValue('moved'));
    expect(session.transactionCount).toBe(1);

    session.dispatch('edit.undo', undefined);
    expect(getCell(sheet, 0, 0)?.value).toEqual(textValue('moved'));
    expect(getCell(second, 0, 0)).toBeNull();
  });

  it('copies merged ranges and restores the merge list on undo', () => {
    const { session, sheet } = setup();
    put(sheet, 0, 0, textValue('merged'));
    sheet.merges.push({ range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 } });
    const snapshot = clipboardFor(session, sheet.id, {
      startRow: 0,
      startCol: 0,
      endRow: 1,
      endCol: 1,
    });

    session.dispatch('edit.paste', {
      sheetId: sheet.id,
      target: singleCell(3, 3),
      snapshot,
    });
    expect(sheet.merges).toHaveLength(2);
    expect(
      sheet.merges.some(
        (merge) =>
          merge.range.startRow === 3 &&
          merge.range.startCol === 3 &&
          merge.range.endRow === 4 &&
          merge.range.endCol === 4,
      ),
    ).toBe(true);
    expect(session.dirtyMergeSheets().has(sheet.id)).toBe(true);
    expect(session.hasUnsavedChanges).toBe(true);

    session.dispatch('edit.undo', undefined);
    expect(sheet.merges).toHaveLength(1);
    expect(sheet.merges[0]?.range).toEqual({
      startRow: 0,
      startCol: 0,
      endRow: 1,
      endCol: 1,
    });
  });
});

describe('formula paste translation', () => {
  it('offsets relative references while keeping absolute and sheet-qualified refs correct', () => {
    expect(translateFormulaForPaste('A1+$B$2+Sheet2!C3', 1, 2)).toBe('C2+$B$2+Sheet2!E4');
  });

  it('returns #REF! when a relative reference leaves the grid', () => {
    expect(translateFormulaForPaste('A1+1', -1, 0)).toBe('#REF!+1');
  });

  it('offsets whole-column references on a column move', () => {
    expect(translateFormulaForPaste('SUM(A:A)', 0, 1)).toBe('SUM(B:B)');
  });
});
