import { describe, expect, it } from 'vitest';
import { DocumentSession } from '@commands/session';
import type { GridSelection } from '@grid/frame';
import { resolveNavigation } from '@grid/navigation';
import {
  activeCell,
  collapseSelection,
  createSelection,
  extendSelection,
  isSingleCell,
  moveRange,
  selectionForRange,
  selectionRange,
} from '@grid/selection';
import type { Sheet } from '@model/types';
import { getCell, setCellSnapshot } from '@model/types';
import { numberValue } from '@model/value';
import { createEmptyWorkbook, createSheet } from '@model/workbook';
import { createSheetProjection } from '@view/projection';
import type { SheetProjection } from '@view/types';

/**
 * Ticket 31 A: the active cell of a selection is the cell the gesture began
 * at (the anchor), not the end the pointer or a Shift extension reached. Every
 * consumer — the name box's command context, the formula bar, in-cell editing,
 * formula pointing and the style readouts — reads this one cell.
 */

function projectionFor(): { projection: SheetProjection; sheet: Sheet } {
  const workbook = createEmptyWorkbook();
  const sheet = createSheet('Sheet1', 'xl/worksheets/sheet1.xml');
  workbook.sheets.push(sheet);
  return { projection: createSheetProjection(workbook, sheet.id, { version: 1 }), sheet };
}

/** The grid's keyboard rule: a plain key starts at the active cell. */
function navigate(
  projection: SheetProjection,
  key: string,
  modifiers: { ctrlKey?: boolean; shiftKey?: boolean } = {},
  selection: GridSelection = createSelection({ row: 0, col: 0 }),
) {
  const shiftKey = modifiers.shiftKey ?? false;
  return resolveNavigation({
    key,
    ctrlKey: modifiers.ctrlKey ?? false,
    shiftKey,
    altKey: false,
    selection: shiftKey ? selection : collapseSelection(activeCell(selection)),
    projection,
    pageRows: 20,
    pageCols: 10,
  });
}

function moved(outcome: ReturnType<typeof navigate>): GridSelection {
  if (outcome.kind === 'none') throw new Error('expected a selection outcome');
  return outcome.selection;
}

describe('the active cell is the gesture origin', () => {
  it('is the clicked cell of a collapsed selection', () => {
    const selection = createSelection({ row: 2, col: 1 });
    expect(activeCell(selection)).toEqual({ row: 2, col: 1 });
    expect(isSingleCell(selection)).toBe(true);
  });

  it('stays at the anchor of a drag, whatever corner the pointer reaches', () => {
    const origin = { row: 2, col: 1 };
    // A drag up and to the left still began at B3.
    for (const target of [
      { row: 5, col: 3 },
      { row: 0, col: 0 },
      { row: 2, col: 0 },
      { row: 0, col: 3 },
    ]) {
      const selection = extendSelection(collapseSelection(origin), target);
      expect(activeCell(selection)).toEqual(origin);
      expect(selectionRange(selection)).toEqual({
        startRow: Math.min(origin.row, target.row),
        startCol: Math.min(origin.col, target.col),
        endRow: Math.max(origin.row, target.row),
        endCol: Math.max(origin.col, target.col),
      });
    }
  });

  it('is the origin after a Shift+arrow extension, and Shift extends from the far end', () => {
    const projection = projectionFor().projection;
    const start = createSelection({ row: 2, col: 1 });
    const extended = moved(navigate(projection, 'ArrowRight', { shiftKey: true }, start));
    expect(extended.anchor).toEqual({ row: 2, col: 1 });
    expect(extended.focus).toEqual({ row: 2, col: 2 });
    expect(activeCell(extended)).toEqual({ row: 2, col: 1 });

    // A second Shift+arrow grows the range from the focus end; the active cell
    // does not move with it.
    const further = moved(navigate(projection, 'ArrowDown', { shiftKey: true }, extended));
    expect(selectionRange(further)).toEqual({
      startRow: 2,
      startCol: 1,
      endRow: 3,
      endCol: 2,
    });
    expect(activeCell(further)).toEqual({ row: 2, col: 1 });
  });

  it('moves with a plain arrow, which collapses the selection onto the new cell', () => {
    const projection = projectionFor().projection;
    const extended = moved(
      navigate(projection, 'ArrowRight', { shiftKey: true }, createSelection({ row: 2, col: 1 })),
    );
    // Excel moves from the active cell, not from the extension end.
    const collapsed = moved(navigate(projection, 'ArrowDown', {}, extended));
    expect(collapsed.anchor).toEqual({ row: 3, col: 1 });
    expect(collapsed.focus).toEqual({ row: 3, col: 1 });
    expect(activeCell(collapsed)).toEqual({ row: 3, col: 1 });
    expect(isSingleCell(collapsed)).toBe(true);
  });

  it('stays where it is through Ctrl+A, a whole column, a whole row and a region', () => {
    const projection = projectionFor().projection;
    const start = createSelection({ row: 4, col: 2 });

    const sheet = moved(navigate(projection, 'a', { ctrlKey: true }, start));
    expect(selectionRange(sheet)).toEqual({
      startRow: 0,
      startCol: 0,
      endRow: projection.rowCount - 1,
      endCol: projection.columnCount - 1,
    });
    expect(activeCell(sheet)).toEqual({ row: 4, col: 2 });

    const column = moved(navigate(projection, ' ', { ctrlKey: true }, start));
    expect(activeCell(column)).toEqual({ row: 4, col: 2 });

    const row = moved(navigate(projection, ' ', { shiftKey: true }, start));
    expect(activeCell(row)).toEqual({ row: 4, col: 2 });

    // A range selection keeps the focus the caller named as its active cell.
    const region = selectionForRange(
      { startRow: 1, startCol: 1, endRow: 4, endCol: 3 },
      { row: 2, col: 2 },
    );
    expect(activeCell(region)).toEqual({ row: 2, col: 2 });
    const byDefault = selectionForRange({ startRow: 1, startCol: 1, endRow: 4, endCol: 3 });
    expect(activeCell(byDefault)).toEqual({ row: 4, col: 3 });
  });
});

describe('the pasted range becomes the selection', () => {
  it('makes the range top-left the active cell and one undo restores the paste', () => {
    const workbook = createEmptyWorkbook();
    const sheet = createSheet('Sheet1', 'xl/worksheets/sheet1.xml');
    workbook.sheets.push(sheet);
    setCellSnapshot(sheet, 0, 0, {
      value: numberValue(3),
      formula: null,
      formulaDerived: false,
      savedResult: null,
      styleId: 0,
      richText: null,
      stale: false,
    });
    const session = new DocumentSession(workbook);

    const copy = session.dispatch('edit.copy', {
      sheetId: sheet.id,
      range: { startRow: 0, startCol: 0, endRow: 0, endCol: 0 },
    });
    const snapshot = copy.clipboard;
    expect(snapshot).toBeDefined();
    if (snapshot === undefined) throw new Error('expected a snapshot');

    const paste = session.dispatch('edit.paste', {
      sheetId: sheet.id,
      target: { startRow: 2, startCol: 4, endRow: 2, endCol: 4 },
      snapshot,
    });
    expect(session.transactionCount).toBe(1);
    expect(paste.selectionAfter).toEqual({ startRow: 2, startCol: 4, endRow: 2, endCol: 4 });
    // The chrome selects the pasted range with its top-left cell active.
    const after = paste.selectionAfter;
    if (after === undefined) throw new Error('expected a selection');
    const selection = selectionForRange(after, { row: after.startRow, col: after.startCol });
    expect(activeCell(selection)).toEqual({ row: 2, col: 4 });
    expect(getCell(sheet, 2, 4)?.value).toEqual(numberValue(3));

    session.dispatch('edit.undo', undefined);
    expect(getCell(sheet, 2, 4)).toBeNull();
    expect(getCell(sheet, 0, 0)?.value).toEqual(numberValue(3));
  });

  it('clears a cut source in the same transaction as the paste', () => {
    const workbook = createEmptyWorkbook();
    const sheet = createSheet('Sheet1', 'xl/worksheets/sheet1.xml');
    workbook.sheets.push(sheet);
    for (const [row, value] of [
      [0, 5],
      [1, 6],
    ] as const) {
      setCellSnapshot(sheet, row, 0, {
        value: numberValue(value),
        formula: null,
        formulaDerived: false,
        savedResult: null,
        styleId: 0,
        richText: null,
        stale: false,
      });
    }
    const session = new DocumentSession(workbook);

    const cut = session.dispatch('edit.cut', {
      sheetId: sheet.id,
      range: { startRow: 0, startCol: 0, endRow: 1, endCol: 0 },
    });
    const snapshot = cut.clipboard;
    if (snapshot === undefined) throw new Error('expected a snapshot');

    // The marquee keeps the copied size, so the destination is the same shape.
    const paste = session.dispatch('edit.paste', {
      sheetId: sheet.id,
      target: moveRange(
        { startRow: 0, startCol: 0, endRow: 1, endCol: 0 },
        0,
        2,
        1_048_576,
        16_384,
      ),
      snapshot,
    });
    expect(session.transactionCount).toBe(1);
    expect(paste.selectionAfter).toEqual({ startRow: 0, startCol: 2, endRow: 1, endCol: 2 });
    expect(getCell(sheet, 0, 2)?.value).toEqual(numberValue(5));
    expect(getCell(sheet, 1, 2)?.value).toEqual(numberValue(6));
    expect(getCell(sheet, 0, 0)).toBeNull();
    expect(getCell(sheet, 1, 0)).toBeNull();

    // One undo restores the source and removes the pasted copy.
    session.dispatch('edit.undo', undefined);
    expect(getCell(sheet, 0, 0)?.value).toEqual(numberValue(5));
    expect(getCell(sheet, 1, 0)?.value).toEqual(numberValue(6));
    expect(getCell(sheet, 0, 2)).toBeNull();
    expect(getCell(sheet, 1, 2)).toBeNull();
  });
});

describe('the copy marquee moves one step and clamps at the sheet edge', () => {
  it('moves by whole rows and columns', () => {
    const range = { startRow: 2, startCol: 2, endRow: 3, endCol: 4 };
    expect(moveRange(range, 1, 0, 100, 100)).toEqual({ ...range, startRow: 3, endRow: 4 });
    expect(moveRange(range, 0, -1, 100, 100)).toEqual({ ...range, startCol: 1, endCol: 3 });
    expect(moveRange(range, -2, 0, 100, 100)).toEqual({ ...range, startRow: 0, endRow: 1 });
  });

  it('clamps instead of wrapping at row 1 and the last row or column', () => {
    const top = { startRow: 0, startCol: 0, endRow: 1, endCol: 1 };
    expect(moveRange(top, -1, -1, 100, 100)).toEqual(top);

    const bottom = moveRange(top, 99, 99, 100, 100);
    expect(bottom).toEqual({ startRow: 98, startCol: 98, endRow: 99, endCol: 99 });
    expect(moveRange(bottom, 1, 1, 100, 100)).toEqual(bottom);
  });

  it('cannot leave the sheet with the real dimensions', () => {
    const single = { startRow: 1_048_575, startCol: 16_383, endRow: 1_048_575, endCol: 16_383 };
    expect(moveRange(single, 1, 1, 1_048_576, 16_384)).toEqual(single);
    const back = moveRange(single, -1, -1, 1_048_576, 16_384);
    expect(back).toEqual({
      startRow: 1_048_574,
      startCol: 16_382,
      endRow: 1_048_574,
      endCol: 16_382,
    });
  });
});
