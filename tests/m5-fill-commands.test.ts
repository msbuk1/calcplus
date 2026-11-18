import { describe, expect, it } from 'vitest';
import { inventoryCommand } from '@commands/inventory';
import { createCommandRegistry } from '@commands/registry';
import { DocumentSession } from '@commands/session';
import { MERGED_FILL_MESSAGE } from '@model/fill';
import type { Cell, GridRange, Sheet } from '@model/types';
import { cellKey, cloneCell, createCell, getCell, setCellSnapshot } from '@model/types';
import { numberValue, textValue, type CellValue } from '@model/value';
import { createEmptyWorkbook, createSheet } from '@model/workbook';

/**
 * The command half of the fill handle (feature matrix §1; command-system §4:
 * "Fill-handle drag or double-click fill | One"). One gesture, one transaction;
 * undo restores the exact prior state and redo reapplies it. `fill.handle` and
 * the Ctrl+D / Ctrl+R commands all run through the same model plan.
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

function putFormula(sheet: Sheet, row: number, col: number, formula: string): void {
  setCellSnapshot(sheet, row, col, {
    value: numberValue(0),
    formula,
    formulaDerived: false,
    savedResult: null,
    styleId: 0,
    richText: null,
    stale: false,
  } as Cell);
}

/** Everything a fill can touch, deep enough that a lost field fails. */
function sheetState(sheet: Sheet) {
  return {
    cells: [...sheet.cells.entries()]
      .sort(([left], [right]) => left.localeCompare(right))
      .map(([address, cell]) => [address, cloneCell(cell)] as const),
    merges: sheet.merges.map((merge) => ({ ...merge.range })),
  };
}

function valueAt(sheet: Sheet, row: number, col: number): string {
  const cell = getCell(sheet, row, col);
  if (cell === null) return '<empty>';
  if (cell.value.type === 'text') return cell.value.value;
  if (cell.value.type === 'blank') return '<blank>';
  return String(cell.value.value);
}

function formulaAt(sheet: Sheet, row: number, col: number): string | null {
  return getCell(sheet, row, col)?.formula ?? null;
}

function dispatchFill(
  session: DocumentSession,
  sheet: Sheet,
  selection: GridRange,
  args: Record<string, unknown>,
) {
  return session.dispatch('fill.handle', args, {
    activeSheetId: sheet.id,
    selection,
    activeCell: { row: selection.startRow, col: selection.startCol },
  });
}

// ---------------------------------------------------------------------------
// One gesture, one transaction
// ---------------------------------------------------------------------------

describe('fill.handle', () => {
  it('fills the dragged rectangle as one transaction and undoes exactly', () => {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, numberValue(10));
    putFormula(sheet, 0, 1, 'A1*2');
    put(sheet, 1, 0, numberValue(20));
    putFormula(sheet, 1, 1, 'A2*2');
    const session = new DocumentSession(workbook);
    const before = sheetState(sheet);

    const result = dispatchFill(session, sheet, range(1, 1, 1, 1), {
      source: range(0, 0, 1, 1),
      destination: range(0, 0, 5, 1),
    });

    expect(result.fill?.state).toBe('filled');
    expect(result.fill?.cellCount).toBe(8);
    expect(result.fill?.destination).toEqual(range(0, 0, 5, 1));
    // One band: the extension below the source. The whole gesture is one entry.
    expect(result.changes).toHaveLength(1);
    expect(result.changes[0]?.type).toBe('cell.set');
    expect(session.transactionCount).toBe(1);
    expect(valueAt(sheet, 5, 0)).toBe('20');
    expect(formulaAt(sheet, 5, 1)).toBe('A6*2');
    expect(result.selectionAfter).toEqual(range(0, 0, 5, 1));

    session.dispatch('edit.undo', undefined);
    expect(sheetState(sheet)).toEqual(before);
    expect(session.transactionCount).toBe(0);

    session.dispatch('edit.redo', undefined);
    expect(valueAt(sheet, 2, 0)).toBe('10');
    expect(formulaAt(sheet, 4, 1)).toBe('A5*2');
    expect(valueAt(sheet, 5, 0)).toBe('20');
  });

  it('fills across and down in one transaction when the drag covers both axes', () => {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, numberValue(1));
    putFormula(sheet, 0, 1, 'A1');
    const session = new DocumentSession(workbook);

    const result = dispatchFill(session, sheet, range(0, 1, 0, 1), {
      source: range(0, 0, 0, 1),
      destination: range(0, 0, 1, 2),
    });

    expect(result.fill?.state).toBe('filled');
    // The destination minus the source is two disjoint bands (below and right),
    // so the whole gesture is still one transaction with two patches.
    expect(result.changes).toHaveLength(2);
    expect(session.transactionCount).toBe(1);
    expect(valueAt(sheet, 1, 0)).toBe('1');
    expect(formulaAt(sheet, 1, 1)).toBe('A2');
    // Column C maps back to the source's first column: the two-column block
    // wraps, so C is the first column of the next tile.
    expect(valueAt(sheet, 0, 2)).toBe('1');
    expect(valueAt(sheet, 1, 2)).toBe('1');

    session.dispatch('edit.undo', undefined);
    expect(getCell(sheet, 1, 1)).toBeNull();
    expect(getCell(sheet, 0, 2)).toBeNull();
    expect(getCell(sheet, 1, 2)).toBeNull();
  });

  it('is a silent no-op when the handle is clicked without a drag', () => {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, numberValue(1));
    const session = new DocumentSession(workbook);

    const result = dispatchFill(session, sheet, range(0, 0, 0, 0), {
      source: range(0, 0, 0, 0),
      destination: range(0, 0, 0, 0),
    });

    expect(result.fill).toMatchObject({ state: 'unchanged', reason: 'no-destination' });
    expect(result.changes).toHaveLength(0);
    expect(session.transactionCount).toBe(0);
  });

  it('defaults the source to the selection and the destination to the source', () => {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, numberValue(1));
    const session = new DocumentSession(workbook);

    const result = session.dispatch(
      'fill.handle',
      {},
      {
        activeSheetId: sheet.id,
        selection: range(0, 0, 0, 0),
        activeCell: { row: 0, col: 0 },
      },
    );
    expect(result.fill?.state).toBe('unchanged');
    expect(session.transactionCount).toBe(0);
  });

  it('blocks a merged destination with a toast and no transaction', () => {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, numberValue(1));
    sheet.merges.push({ range: range(3, 0, 3, 1) });
    const session = new DocumentSession(workbook);
    const before = sheetState(sheet);

    const result = dispatchFill(session, sheet, range(0, 0, 0, 0), {
      source: range(0, 0, 0, 0),
      destination: range(0, 0, 4, 0),
    });

    expect(result.fill).toMatchObject({ state: 'blocked', code: 'merged' });
    expect(result.ui).toEqual({ type: 'toast', message: MERGED_FILL_MESSAGE });
    expect(result.changes).toHaveLength(0);
    expect(session.transactionCount).toBe(0);
    expect(sheetState(sheet)).toEqual(before);
  });
});

// ---------------------------------------------------------------------------
// Ctrl+D and Ctrl+R
// ---------------------------------------------------------------------------

describe('edit.fillDown', () => {
  it('fills every row below the selection first row from that row', () => {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, textValue('label'));
    put(sheet, 0, 1, numberValue(7), 2);
    const session = new DocumentSession(workbook);

    const result = session.dispatch(
      'edit.fillDown',
      { range: range(0, 0, 2, 1) },
      {
        activeSheetId: sheet.id,
        selection: range(0, 0, 2, 1),
        activeCell: { row: 0, col: 0 },
      },
    );

    expect(result.fill).toMatchObject({ state: 'filled', source: range(0, 0, 0, 1) });
    expect(session.transactionCount).toBe(1);
    expect(valueAt(sheet, 2, 0)).toBe('label');
    expect(valueAt(sheet, 2, 1)).toBe('7');
    expect(getCell(sheet, 2, 1)?.styleId).toBe(2);

    session.dispatch('edit.undo', undefined);
    expect(getCell(sheet, 1, 0)).toBeNull();
    expect(getCell(sheet, 2, 1)).toBeNull();
  });

  it('adjusts a column formula down the selection', () => {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, numberValue(1));
    putFormula(sheet, 0, 1, 'A1*2');
    put(sheet, 1, 0, numberValue(2));
    put(sheet, 2, 0, numberValue(3));
    const session = new DocumentSession(workbook);

    session.dispatch(
      'edit.fillDown',
      { range: range(0, 1, 2, 1) },
      {
        activeSheetId: sheet.id,
        selection: range(0, 1, 2, 1),
        activeCell: { row: 0, col: 1 },
      },
    );

    expect(formulaAt(sheet, 1, 1)).toBe('A2*2');
    expect(formulaAt(sheet, 2, 1)).toBe('A3*2');
  });

  it('is a no-op for a one-row selection', () => {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, numberValue(1));
    const session = new DocumentSession(workbook);

    const result = session.dispatch(
      'edit.fillDown',
      { range: range(0, 0, 0, 0) },
      {
        activeSheetId: sheet.id,
        selection: range(0, 0, 0, 0),
        activeCell: { row: 0, col: 0 },
      },
    );
    expect(result.fill).toMatchObject({ state: 'unchanged', reason: 'no-destination' });
    expect(session.transactionCount).toBe(0);
  });
});

describe('edit.fillRight', () => {
  it('fills every column right of the selection first column from that column', () => {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, numberValue(1));
    putFormula(sheet, 0, 1, 'A1*10');
    put(sheet, 1, 0, numberValue(2));
    putFormula(sheet, 1, 1, 'A2*10');
    const session = new DocumentSession(workbook);

    const result = session.dispatch(
      'edit.fillRight',
      { range: range(0, 1, 1, 2) },
      {
        activeSheetId: sheet.id,
        selection: range(0, 1, 1, 2),
        activeCell: { row: 0, col: 1 },
      },
    );

    expect(result.fill).toMatchObject({ state: 'filled', source: range(0, 1, 1, 1) });
    expect(session.transactionCount).toBe(1);
    expect(formulaAt(sheet, 0, 2)).toBe('B1*10');
    expect(formulaAt(sheet, 1, 2)).toBe('B2*10');

    session.dispatch('edit.undo', undefined);
    expect(getCell(sheet, 0, 2)).toBeNull();
    expect(formulaAt(sheet, 0, 1)).toBe('A1*10');
  });

  it('is a no-op for a one-column selection', () => {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, numberValue(1));
    const session = new DocumentSession(workbook);

    const result = session.dispatch(
      'edit.fillRight',
      { range: range(0, 0, 2, 0) },
      {
        activeSheetId: sheet.id,
        selection: range(0, 0, 2, 0),
        activeCell: { row: 0, col: 0 },
      },
    );
    expect(result.fill).toMatchObject({ state: 'unchanged', reason: 'no-destination' });
    expect(session.transactionCount).toBe(0);
  });
});

// ---------------------------------------------------------------------------
// Registry, map and keyboard parity
// ---------------------------------------------------------------------------

describe('registration and the command map', () => {
  it('registers the fill commands with the keyboard map bindings', () => {
    const registry = createCommandRegistry();
    expect(registry.require('fill.handle').keybindings).toEqual([]);
    expect(registry.require('edit.fillDown').keybindings).toContain('Ctrl+D');
    expect(registry.require('edit.fillRight').keybindings).toContain('Ctrl+R');
    for (const id of ['fill.handle', 'edit.fillDown', 'edit.fillRight']) {
      expect(
        registry.require(id).isEnabled({
          workbook: createEmptyWorkbook(),
          activeSheetId: null,
          selection: null,
          activeCell: null,
        }),
      ).toBe(false);
    }
  });

  it('maps all three commands to the fill-handle matrix row', () => {
    const feature =
      'Fill handle: copy, relative-reference adjustment, double-click to adjacent data';
    const fill = inventoryCommand('fill.handle');
    const down = inventoryCommand('edit.fillDown');
    const right = inventoryCommand('edit.fillRight');
    expect(fill?.implemented).toBe(true);
    expect(down?.implemented).toBe(true);
    expect(right?.implemented).toBe(true);
    expect(down?.matrixFeature).toBe(feature);
    expect(right?.matrixFeature).toBe(feature);
    expect(fill?.matrixFeature).toBe(feature);
    expect(down?.shortcut).toEqual(['Ctrl+D']);
    expect(right?.shortcut).toEqual(['Ctrl+R']);
  });

  it('keeps every filled cell address in the change patch', () => {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, numberValue(1));
    const session = new DocumentSession(workbook);
    const result = dispatchFill(session, sheet, range(0, 0, 0, 0), {
      source: range(0, 0, 0, 0),
      destination: range(0, 0, 2, 0),
    });
    const change = result.changes[0];
    expect(change?.type).toBe('cell.set');
    if (change?.type !== 'cell.set') throw new Error('expected a cell.set change');
    expect(change.range).toEqual(range(1, 0, 2, 0));
    expect(change.before).toHaveLength(2);
    expect(change.after).toHaveLength(2);
    expect(change.after[1]?.[0]?.value).toEqual({ type: 'number', value: 1 });
    expect(sheet.cells.has(cellKey(0, 0))).toBe(true);
  });
});
