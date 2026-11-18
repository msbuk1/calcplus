import { describe, expect, it } from 'vitest';
import { DocumentSession } from '@commands/session';
import { createCommandRegistry } from '@commands/registry';
import {
  MAX_CONDITIONS_PER_COLUMN,
  computeHiddenLayers,
  createAutoFilter,
  type ColumnFilter,
} from '@model/autofilter';
import type { GridRange, Sheet } from '@model/types';
import { createCell, setCellSnapshot } from '@model/types';
import type { CellValue } from '@model/value';
import { numberValue, textValue } from '@model/value';
import { createEmptyWorkbook, createSheet } from '@model/workbook';

/**
 * AutoFilter fixtures F16 (one filter per sheet) and F18 (apply, change and
 * clear are one undo step each; reapply is a no-op) of sort-autofilter §4, plus
 * the §3.3 layer rule that clearing a filter keeps manual hidden rows hidden.
 * Integration style: the real model, the real command registry, no DOM.
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

function put(sheet: Sheet, row: number, col: number, value: CellValue): void {
  setCellSnapshot(sheet, row, col, createCell(value));
}

/** A1:B5 is one block, D1:E4 another with an empty column between them. */
function filterSheet() {
  const { workbook, sheet } = workbookWithSheet();
  put(sheet, 0, 0, textValue('Name'));
  put(sheet, 0, 1, textValue('Score'));
  put(sheet, 1, 0, textValue('Alice'));
  put(sheet, 1, 1, numberValue(30));
  put(sheet, 2, 0, textValue('Bob'));
  put(sheet, 2, 1, numberValue(10));
  put(sheet, 3, 0, textValue('Carol'));
  put(sheet, 3, 1, numberValue(20));
  put(sheet, 4, 0, textValue('Dave'));
  put(sheet, 4, 1, numberValue(5));

  put(sheet, 0, 3, textValue('Region'));
  put(sheet, 0, 4, textValue('Count'));
  put(sheet, 1, 3, textValue('North'));
  put(sheet, 1, 4, numberValue(1));
  put(sheet, 2, 3, textValue('South'));
  put(sheet, 2, 4, numberValue(2));
  put(sheet, 3, 3, textValue('East'));
  put(sheet, 3, 4, numberValue(3));
  return { workbook, sheet };
}

function manualHiddenRows(sheet: Sheet): Set<number> {
  const hidden = new Set<number>();
  for (const [row, dimension] of sheet.rowDimensions) {
    if (dimension.hidden === true) hidden.add(row);
  }
  return hidden;
}

/** The union of both layers: what a row renderer would hide right now. */
function effectiveHiddenRows(sheet: Sheet): Set<number> {
  return new Set(
    computeHiddenLayers(sheet, sheet.autoFilter, manualHiddenRows(sheet)).effectiveHidden,
  );
}

function filterFor(sheet: Sheet, columns: readonly (readonly [number, ColumnFilter])[]) {
  return createAutoFilter(sheet.id, range(0, 0, 4, 1), columns);
}

function valuesFilter(...values: CellValue[]): ColumnFilter {
  return { kind: 'values', selected: values, includeBlanks: false };
}

function enableFilter(session: DocumentSession, sheet: Sheet, target: GridRange) {
  return session.dispatch('data.filter', { range: target }, { activeSheetId: sheet.id });
}

function setColumn(
  session: DocumentSession,
  sheet: Sheet,
  column: number,
  criterion: ColumnFilter | null,
  confirm?: boolean,
) {
  return session.dispatch(
    'data.filterColumn',
    { column, criterion, ...(confirm === undefined ? {} : { confirm }) },
    { activeSheetId: sheet.id, selection: range(0, 0, 4, 1) },
  );
}

describe('data.filter toggle', () => {
  it('enables on the current region and isChecked follows the sheet', () => {
    const { workbook, sheet } = filterSheet();
    const session = new DocumentSession(workbook);
    const registry = createCommandRegistry();
    const command = registry.require('data.filter');
    const context = {
      workbook,
      activeSheetId: sheet.id,
      selection: range(0, 0, 0, 0),
      activeCell: { row: 0, col: 0 },
    };

    expect(command.isChecked?.(context)).toBe(false);
    const result = session.dispatch(
      'data.filter',
      {},
      { activeSheetId: sheet.id, selection: range(0, 0, 0, 0) },
    );

    expect(result.filter?.state).toBe('applied');
    expect(result.filter?.changed).toBe(true);
    expect(result.changes).toHaveLength(1);
    expect(result.changes[0]!.type).toBe('filter.set');
    // A single cell expands to its own contiguous region, not the next block.
    expect(sheet.autoFilter?.range).toEqual(range(0, 0, 4, 1));
    expect(command.isChecked?.(context)).toBe(true);
    expect(session.transactionCount).toBe(1);

    session.dispatch('edit.undo', undefined);
    expect(sheet.autoFilter).toBeNull();
    expect(command.isChecked?.(context)).toBe(false);
  });

  it('registers every filter command and the sort command under the data ids', () => {
    const registry = createCommandRegistry();
    for (const id of [
      'data.sort',
      'data.filter',
      'data.filterColumn',
      'data.filterClearColumn',
      'data.filterClear',
      'data.filterReapply',
    ]) {
      expect(registry.get(id)?.id).toBe(id);
    }
  });
});

describe('F16 filter-one-per-sheet', () => {
  it('replaces an existing filter on another range in one transaction', () => {
    const { workbook, sheet } = filterSheet();
    const session = new DocumentSession(workbook);

    enableFilter(session, sheet, range(0, 0, 4, 1));
    setColumn(session, sheet, 1, valuesFilter(numberValue(10)));
    expect(session.transactionCount).toBe(2);
    expect(effectiveHiddenRows(sheet)).toEqual(new Set([1, 3, 4]));

    const replacement = enableFilter(session, sheet, range(0, 3, 3, 4));

    // One transaction replaces the range and drops the old criteria.
    expect(replacement.changes).toHaveLength(1);
    expect(session.transactionCount).toBe(3);
    expect(sheet.autoFilter?.range).toEqual(range(0, 3, 3, 4));
    expect(sheet.autoFilter?.columns.size).toBe(0);
    // The old filter-hidden rows are visible again; the new filter hides nothing.
    expect(effectiveHiddenRows(sheet)).toEqual(new Set());

    session.dispatch('edit.undo', undefined);
    expect(sheet.autoFilter?.range).toEqual(range(0, 0, 4, 1));
    expect(sheet.autoFilter?.columns.get(1)).toEqual(valuesFilter(numberValue(10)));
    expect(effectiveHiddenRows(sheet)).toEqual(new Set([1, 3, 4]));
  });
});

describe('F18 filter-undo', () => {
  it('makes apply, change and clear one undo step each, and reapply a no-op', () => {
    const { workbook, sheet } = filterSheet();
    const session = new DocumentSession(workbook);
    enableFilter(session, sheet, range(0, 0, 4, 1));
    const base = session.transactionCount;

    // Apply a value checklist.
    const applied = setColumn(session, sheet, 1, valuesFilter(numberValue(10), numberValue(20)));
    expect(applied.filter?.state).toBe('changed');
    expect(applied.filter?.changed).toBe(true);
    expect(applied.filter?.hiddenChanges).toEqual([
      { row: 1, before: false, after: true },
      { row: 4, before: false, after: true },
    ]);
    expect(applied.filter?.rowsVisible).toBe(2);
    expect(applied.filter?.rowsTotal).toBe(4);
    expect(session.transactionCount).toBe(base + 1);
    expect(effectiveHiddenRows(sheet)).toEqual(new Set([1, 4]));

    session.dispatch('edit.undo', undefined);
    expect(effectiveHiddenRows(sheet)).toEqual(new Set());
    expect(sheet.autoFilter?.columns.size).toBe(0);
    expect(session.transactionCount).toBe(base);

    session.dispatch('edit.redo', undefined);
    expect(sheet.autoFilter?.columns.get(1)).toEqual(
      valuesFilter(numberValue(10), numberValue(20)),
    );
    expect(effectiveHiddenRows(sheet)).toEqual(new Set([1, 4]));

    // Change the criterion to a condition.
    const changed = setColumn(session, sheet, 1, {
      kind: 'conditions',
      join: 'and',
      conditions: [{ kind: 'number', operator: '>', value: 15 }],
    });
    expect(changed.filter?.state).toBe('changed');
    expect(session.transactionCount).toBe(base + 2);
    expect(effectiveHiddenRows(sheet)).toEqual(new Set([2, 4]));

    session.dispatch('edit.undo', undefined);
    expect(effectiveHiddenRows(sheet)).toEqual(new Set([1, 4]));
    expect(sheet.autoFilter?.columns.get(1)?.kind).toBe('values');

    session.dispatch('edit.redo', undefined);
    expect(effectiveHiddenRows(sheet)).toEqual(new Set([2, 4]));
    expect(sheet.autoFilter?.columns.get(1)?.kind).toBe('conditions');

    // Clear All keeps the filter and drops the criteria in one step.
    const cleared = session.dispatch('data.filterClear', {}, { activeSheetId: sheet.id });
    expect(cleared.filter?.state).toBe('cleared');
    expect(cleared.filter?.hiddenChanges).toEqual([
      { row: 2, before: true, after: false },
      { row: 4, before: true, after: false },
    ]);
    expect(sheet.autoFilter).not.toBeNull();
    expect(sheet.autoFilter?.columns.size).toBe(0);
    expect(session.transactionCount).toBe(base + 3);
    expect(effectiveHiddenRows(sheet)).toEqual(new Set());

    session.dispatch('edit.undo', undefined);
    expect(sheet.autoFilter?.columns.get(1)?.kind).toBe('conditions');
    expect(effectiveHiddenRows(sheet)).toEqual(new Set([2, 4]));

    session.dispatch('edit.redo', undefined);
    expect(sheet.autoFilter?.columns.size).toBe(0);

    // Reapply finds no change and creates no transaction.
    const reapplied = session.dispatch('data.filterReapply', {}, { activeSheetId: sheet.id });
    expect(reapplied.changes).toHaveLength(0);
    expect(reapplied.filter?.state).toBe('reapplied');
    expect(reapplied.filter?.changed).toBe(false);
    expect(session.transactionCount).toBe(base + 3);
  });

  it('clearing one column is its own step and leaves the other criteria alone', () => {
    const { workbook, sheet } = filterSheet();
    const session = new DocumentSession(workbook);
    enableFilter(session, sheet, range(0, 0, 4, 1));
    setColumn(session, sheet, 0, valuesFilter(textValue('Alice'), textValue('Bob')));
    setColumn(session, sheet, 1, {
      kind: 'conditions',
      join: 'and',
      conditions: [{ kind: 'number', operator: '>=', value: 10 }],
    });
    // Columns combine with AND: Alice (30), Bob (10) and Carol (20) pass the
    // condition, so only the non-matching name rows stay hidden.
    expect(effectiveHiddenRows(sheet)).toEqual(new Set([3, 4]));

    const cleared = session.dispatch(
      'data.filterClearColumn',
      { column: 0 },
      { activeSheetId: sheet.id },
    );
    expect(cleared.filter?.state).toBe('cleared');
    expect(sheet.autoFilter?.columns.has(0)).toBe(false);
    expect(sheet.autoFilter?.columns.has(1)).toBe(true);
    expect(effectiveHiddenRows(sheet)).toEqual(new Set([4]));

    session.dispatch('edit.undo', undefined);
    expect(sheet.autoFilter?.columns.has(0)).toBe(true);
    expect(effectiveHiddenRows(sheet)).toEqual(new Set([3, 4]));
  });
});

describe('filter hidden-row layers (§3.3)', () => {
  it('toggling off removes only the filter layer and keeps manual hidden rows hidden', () => {
    const { workbook, sheet } = filterSheet();
    sheet.rowDimensions.set(3, { hidden: true });
    const session = new DocumentSession(workbook);

    enableFilter(session, sheet, range(0, 0, 4, 1));
    setColumn(session, sheet, 1, valuesFilter(numberValue(10), numberValue(20)));
    // Applying a filter never writes the filter layer into the row flags.
    expect([...sheet.rowDimensions.keys()]).toEqual([3]);
    expect(effectiveHiddenRows(sheet)).toEqual(new Set([1, 3, 4]));

    const off = session.dispatch('data.filter', { on: false }, { activeSheetId: sheet.id });
    expect(off.filter?.state).toBe('cleared');
    expect(sheet.autoFilter).toBeNull();
    expect(sheet.rowDimensions.get(3)).toEqual({ hidden: true });
    expect(effectiveHiddenRows(sheet)).toEqual(new Set([3]));

    session.dispatch('edit.undo', undefined);
    expect(sheet.autoFilter).not.toBeNull();
    expect(effectiveHiddenRows(sheet)).toEqual(new Set([1, 3, 4]));
  });
});

describe('filter criteria guards', () => {
  it('asks before replacing an imported criterion it cannot edit', () => {
    const { workbook, sheet } = filterSheet();
    sheet.autoFilter = filterFor(sheet, [
      [1, { kind: 'unsupported', raw: { colour: 'FFFF0000' } }],
    ]);
    const session = new DocumentSession(workbook);

    const asked = setColumn(session, sheet, 1, valuesFilter(numberValue(10)));
    expect(asked.changes).toHaveLength(0);
    expect(asked.ui?.type).toBe('dialog');
    expect(asked.ui?.message).toMatch(/cannot edit/);
    expect(session.transactionCount).toBe(0);

    const replaced = setColumn(session, sheet, 1, valuesFilter(numberValue(10)), true);
    expect(replaced.filter?.changed).toBe(true);
    expect(sheet.autoFilter?.columns.get(1)).toEqual(valuesFilter(numberValue(10)));
    expect(session.transactionCount).toBe(1);
  });

  it('refuses an empty checklist, more than two conditions and a column outside the range', () => {
    const { workbook, sheet } = filterSheet();
    const session = new DocumentSession(workbook);
    enableFilter(session, sheet, range(0, 0, 4, 1));
    const base = session.transactionCount;

    const empty = setColumn(session, sheet, 1, valuesFilter());
    expect(empty.changes).toHaveLength(0);
    expect(empty.ui?.message).toMatch(/at least one value/);

    const tooMany = setColumn(session, sheet, 1, {
      kind: 'conditions',
      join: 'or',
      conditions: [
        { kind: 'number', operator: '>', value: 1 },
        { kind: 'number', operator: '<', value: 9 },
        { kind: 'number', operator: '>=', value: 5 },
      ],
    });
    expect(tooMany.changes).toHaveLength(0);
    expect(tooMany.ui?.message).toContain(`at most ${MAX_CONDITIONS_PER_COLUMN} conditions`);

    const outside = setColumn(session, sheet, 4, valuesFilter(numberValue(1)));
    expect(outside.changes).toHaveLength(0);
    expect(outside.ui?.message).toMatch(/outside the AutoFilter range/);

    expect(session.transactionCount).toBe(base);
  });

  it('keeps the criteria when the toggle is asked to turn on for the same range', () => {
    const { workbook, sheet } = filterSheet();
    const session = new DocumentSession(workbook);
    enableFilter(session, sheet, range(0, 0, 4, 1));
    setColumn(session, sheet, 1, valuesFilter(numberValue(10)));
    const base = session.transactionCount;

    const again = session.dispatch(
      'data.filter',
      { range: range(0, 0, 4, 1), on: true },
      { activeSheetId: sheet.id },
    );
    expect(again.changes).toHaveLength(0);
    expect(again.filter?.changed).toBe(false);
    expect(session.transactionCount).toBe(base);
    expect(sheet.autoFilter?.columns.get(1)).toEqual(valuesFilter(numberValue(10)));
  });
});
