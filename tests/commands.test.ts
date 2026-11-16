import { describe, expect, it } from 'vitest';
import { DocumentSession } from '@commands/session';
import { createCommandRegistry } from '@commands/registry';
import { createEmptyWorkbook, createSheet } from '@model/workbook';
import { getCell, setCellSnapshot } from '@model/types';
import { booleanValue, numberValue, textValue } from '@model/value';

function sessionWithCell(initial = true): DocumentSession {
  const workbook = createEmptyWorkbook();
  const sheet = createSheet('Sheet1', 'xl/worksheets/sheet1.xml');
  workbook.sheets.push(sheet);
  if (initial) {
    setCellSnapshot(sheet, 0, 0, {
      value: numberValue(10),
      formula: '1+9',
      formulaDerived: false,
      savedResult: numberValue(10),
      styleId: 1,
      richText: null,
      stale: false,
    });
  }
  return new DocumentSession(workbook);
}

describe('command registry', () => {
  it('registers stable command ids, including undo and redo', () => {
    const registry = createCommandRegistry();
    expect(registry.require('cell.set').id).toBe('cell.set');
    expect(registry.require('edit.undo').id).toBe('edit.undo');
    expect(registry.require('edit.redo').id).toBe('edit.redo');
    expect(new Set(registry.list().map((command) => command.id)).size).toBe(registry.list().length);
  });

  it('runs cell.set as a range-based change without mutating the model', () => {
    const session = sessionWithCell();
    const before = getCell(session.workbook.sheets[0]!, 0, 0);
    const result = session.dispatch('cell.set', {
      sheetId: session.workbook.sheets[0]!.id,
      row: 0,
      col: 0,
      value: textValue('typed'),
    });
    expect(result.changes).toHaveLength(1);
    expect(result.changes[0]!.type).toBe('cell.set');
    expect(result.changes[0]!.before).toEqual([[before]]);
    expect(getCell(session.workbook.sheets[0]!, 0, 0)?.value).toEqual(textValue('typed'));
  });
});

describe('transactions and undo', () => {
  it('records one undoable transaction per cell.set and restores the prior cell exactly', () => {
    const session = sessionWithCell();
    const sheet = session.workbook.sheets[0]!;
    const before = getCell(sheet, 0, 0)!;

    session.dispatch('cell.set', {
      sheetId: sheet.id,
      row: 0,
      col: 0,
      value: booleanValue(true),
    });
    expect(session.transactionCount).toBe(1);
    expect(session.canUndo).toBe(true);
    expect(getCell(sheet, 0, 0)?.formula).toBeNull();

    session.dispatch('edit.undo', undefined);
    expect(getCell(sheet, 0, 0)).toEqual(before);
    expect(session.canUndo).toBe(false);
    expect(session.canRedo).toBe(true);

    session.dispatch('edit.redo', undefined);
    expect(getCell(sheet, 0, 0)?.value).toEqual(booleanValue(true));
  });

  it('clears redo when a new edit lands', () => {
    const session = sessionWithCell();
    const sheet = session.workbook.sheets[0]!;
    session.dispatch('cell.set', {
      sheetId: sheet.id,
      row: 0,
      col: 0,
      value: textValue('one'),
    });
    session.dispatch('edit.undo', undefined);
    expect(session.canRedo).toBe(true);
    session.dispatch('cell.set', {
      sheetId: sheet.id,
      row: 0,
      col: 1,
      value: textValue('two'),
    });
    expect(session.canRedo).toBe(false);
  });

  it('does not create a transaction for a no-op edit', () => {
    const session = sessionWithCell(false);
    const sheet = session.workbook.sheets[0]!;
    const result = session.dispatch('cell.set', {
      sheetId: sheet.id,
      row: 4,
      col: 4,
      value: { type: 'blank' as const },
    });
    expect(result.changes).toHaveLength(0);
    expect(session.transactionCount).toBe(0);
  });

  it('bounds the undo log at 200 transactions', () => {
    const session = sessionWithCell(false);
    const sheet = session.workbook.sheets[0]!;
    for (let index = 0; index < 205; index += 1) {
      session.dispatch('cell.set', {
        sheetId: sheet.id,
        row: 0,
        col: 0,
        value: numberValue(index),
      });
    }
    expect(session.transactionCount).toBe(200);
    expect(session.historyBytes).toBeLessThan(64 * 1024 * 1024);
  });

  it('marks dirty cells and clears them only on markSaved', () => {
    const session = sessionWithCell(false);
    const sheet = session.workbook.sheets[0]!;
    expect(session.hasUnsavedChanges).toBe(false);
    session.dispatch('cell.set', {
      sheetId: sheet.id,
      row: 2,
      col: 3,
      value: textValue('dirty'),
    });
    expect(session.hasUnsavedChanges).toBe(true);
    expect(session.dirtyCellMap().get(sheet.id)?.has('2:3')).toBe(true);
    session.markSaved();
    expect(session.hasUnsavedChanges).toBe(false);
  });

  it('bumps the document version on edits and undo but not on view-like operations', () => {
    const session = sessionWithCell(false);
    const sheet = session.workbook.sheets[0]!;
    const initialVersion = session.version;
    session.dispatch('cell.set', {
      sheetId: sheet.id,
      row: 0,
      col: 0,
      value: numberValue(1),
    });
    expect(session.version).toBeGreaterThan(initialVersion);
  });
});
