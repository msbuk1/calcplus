import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { openWorkbook } from '@ooxml/open';
import { createSheetProjection } from '@view/projection';
import type { SheetProjection } from '@view/types';
import { resolveNavigation } from '@grid/navigation';
import { createSelection, selectionRange } from '@grid/selection';

const fixture = (name: string): Uint8Array =>
  new Uint8Array(readFileSync(new URL(`../fixtures/compat/${name}`, import.meta.url)));

function m2Projection(): SheetProjection {
  const opened = openWorkbook(fixture('m2-grid.xlsx'));
  return createSheetProjection(opened.workbook, opened.initialSheetId, { version: 1 });
}

function navigate(
  projection: SheetProjection,
  key: string,
  modifiers: { ctrlKey?: boolean; shiftKey?: boolean } = {},
  selection = createSelection({ row: 0, col: 0 }),
) {
  return resolveNavigation({
    key,
    ctrlKey: modifiers.ctrlKey ?? false,
    shiftKey: modifiers.shiftKey ?? false,
    altKey: false,
    selection,
    projection,
    pageRows: 20,
    pageCols: 10,
  });
}

describe('grid navigation', () => {
  it('moves the active cell with the arrow keys and stops at the edges', () => {
    const projection = m2Projection();
    const down = navigate(projection, 'ArrowDown');
    expect(down.kind).toBe('move');
    if (down.kind !== 'move') throw new Error('expected a move');
    expect(down.selection.focus).toEqual({ row: 1, col: 0 });

    const up = navigate(projection, 'ArrowUp', {}, createSelection({ row: 0, col: 0 }));
    if (up.kind !== 'move') throw new Error('expected a move');
    expect(up.selection.focus).toEqual({ row: 0, col: 0 });

    const right = navigate(projection, 'ArrowRight');
    if (right.kind !== 'move') throw new Error('expected a move');
    // A1 is merged with B1, so the arrow moves past the merged range.
    expect(right.selection.focus).toEqual({ row: 0, col: 2 });

    const left = navigate(projection, 'ArrowLeft', {}, createSelection({ row: 5, col: 0 }));
    if (left.kind !== 'move') throw new Error('expected a move');
    expect(left.selection.focus).toEqual({ row: 5, col: 0 });
  });

  it('moves with Tab and Enter and keeps the focus in the grid', () => {
    const projection = m2Projection();
    const tab = navigate(projection, 'Tab', {}, createSelection({ row: 2, col: 1 }));
    if (tab.kind !== 'move') throw new Error('expected a move');
    expect(tab.selection.focus).toEqual({ row: 2, col: 2 });

    const shiftTab = navigate(
      projection,
      'Tab',
      { shiftKey: true },
      createSelection({ row: 2, col: 1 }),
    );
    if (shiftTab.kind !== 'move') throw new Error('expected a move');
    expect(shiftTab.selection.focus).toEqual({ row: 2, col: 0 });

    const enter = navigate(projection, 'Enter', {}, createSelection({ row: 2, col: 1 }));
    if (enter.kind !== 'move') throw new Error('expected a move');
    expect(enter.selection.focus).toEqual({ row: 3, col: 1 });

    const shiftEnter = navigate(
      projection,
      'Enter',
      { shiftKey: true },
      createSelection({ row: 2, col: 1 }),
    );
    if (shiftEnter.kind !== 'move') throw new Error('expected a move');
    expect(shiftEnter.selection.focus).toEqual({ row: 1, col: 1 });
  });

  it('moves past a merged range instead of into the cells it covers', () => {
    const projection = m2Projection();
    const right = navigate(projection, 'ArrowRight', {}, createSelection({ row: 0, col: 0 }));
    if (right.kind !== 'move') throw new Error('expected a move');
    expect(right.selection.focus).toEqual({ row: 0, col: 2 });

    const left = navigate(projection, 'ArrowLeft', {}, createSelection({ row: 0, col: 2 }));
    if (left.kind !== 'move') throw new Error('expected a move');
    expect(left.selection.focus).toEqual({ row: 0, col: 0 });
  });

  it('extends the selection with Shift and the arrow keys', () => {
    const projection = m2Projection();
    const selection = createSelection({ row: 2, col: 1 });
    const outcome = navigate(projection, 'ArrowDown', { shiftKey: true }, selection);
    expect(outcome.kind).toBe('select');
    if (outcome.kind !== 'select') throw new Error('expected a select');
    expect(outcome.selection.anchor).toEqual({ row: 2, col: 1 });
    expect(outcome.selection.focus).toEqual({ row: 3, col: 1 });
    expect(selectionRange(outcome.selection)).toEqual({
      startRow: 2,
      startCol: 1,
      endRow: 3,
      endCol: 1,
    });
  });

  it('jumps to the edge of a block of data with Ctrl and the arrow keys', () => {
    const projection = m2Projection();
    const down = navigate(projection, 'ArrowDown', { ctrlKey: true });
    if (down.kind !== 'move') throw new Error('expected a move');
    // Column A holds data from row 0 to row 301.
    expect(down.selection.focus).toEqual({ row: 301, col: 0 });

    const past = navigate(
      projection,
      'ArrowDown',
      { ctrlKey: true },
      createSelection({ row: 301, col: 0 }),
    );
    if (past.kind !== 'move') throw new Error('expected a move');
    expect(past.selection.focus).toEqual({ row: projection.rowCount - 1, col: 0 });

    const right = navigate(
      projection,
      'ArrowRight',
      { ctrlKey: true },
      createSelection({ row: 2, col: 0 }),
    );
    if (right.kind !== 'move') throw new Error('expected a move');
    // Row 3 holds Item, Amount, Share and a checked note.
    expect(right.selection.focus).toEqual({ row: 2, col: 3 });
  });

  it('moves to the ends of the data with Home, End, Ctrl+Home and Ctrl+End', () => {
    const projection = m2Projection();
    const home = navigate(projection, 'Home', {}, createSelection({ row: 4, col: 1 }));
    if (home.kind !== 'move') throw new Error('expected a move');
    expect(home.selection.focus).toEqual({ row: 4, col: 0 });

    const homeFromColumnD = navigate(projection, 'Home', {}, createSelection({ row: 4, col: 3 }));
    if (homeFromColumnD.kind !== 'move') throw new Error('expected a move');
    expect(homeFromColumnD.selection.focus).toEqual({ row: 4, col: 0 });

    const end = navigate(projection, 'End', {}, createSelection({ row: 4, col: 0 }));
    if (end.kind !== 'move') throw new Error('expected a move');
    expect(end.selection.focus).toEqual({ row: 4, col: 2 });

    const ctrlHome = navigate(projection, 'Home', { ctrlKey: true });
    if (ctrlHome.kind !== 'move') throw new Error('expected a move');
    expect(ctrlHome.selection.focus).toEqual({ row: 0, col: 0 });

    const ctrlEnd = navigate(projection, 'End', { ctrlKey: true });
    if (ctrlEnd.kind !== 'move') throw new Error('expected a move');
    expect(ctrlEnd.selection.focus).toEqual({ row: 301, col: 3 });
  });

  it('pages with Page Up and Page Down', () => {
    const projection = m2Projection();
    const down = navigate(projection, 'PageDown', {}, createSelection({ row: 0, col: 0 }));
    if (down.kind !== 'move') throw new Error('expected a move');
    expect(down.selection.focus).toEqual({ row: 20, col: 0 });

    const up = navigate(projection, 'PageUp', {}, createSelection({ row: 100, col: 0 }));
    if (up.kind !== 'move') throw new Error('expected a move');
    expect(up.selection.focus).toEqual({ row: 80, col: 0 });
  });

  it('selects the used range, a column, a row and the whole sheet', () => {
    const projection = m2Projection();
    const all = navigate(projection, 'a', { ctrlKey: true }, createSelection({ row: 4, col: 1 }));
    if (all.kind !== 'select') throw new Error('expected a select');
    expect(selectionRange(all.selection)).toEqual({
      startRow: 0,
      startCol: 0,
      endRow: 301,
      endCol: 3,
    });
    // The active cell stays where it is, so the view does not scroll away.
    expect(all.selection.focus).toEqual({ row: 4, col: 1 });

    const column = navigate(
      projection,
      ' ',
      { ctrlKey: true },
      createSelection({ row: 4, col: 2 }),
    );
    if (column.kind !== 'select') throw new Error('expected a select');
    expect(selectionRange(column.selection)).toEqual({
      startRow: 0,
      startCol: 2,
      endRow: projection.rowCount - 1,
      endCol: 2,
    });
    // The active cell stays where it is, so the view does not scroll away.
    expect(column.selection.focus).toEqual({ row: 4, col: 2 });

    const row = navigate(projection, ' ', { shiftKey: true }, createSelection({ row: 4, col: 2 }));
    if (row.kind !== 'select') throw new Error('expected a select');
    expect(selectionRange(row.selection)).toEqual({
      startRow: 4,
      startCol: 0,
      endRow: 4,
      endCol: projection.columnCount - 1,
    });
    expect(row.selection.focus).toEqual({ row: 4, col: 2 });

    const sheet = navigate(projection, ' ', { ctrlKey: true, shiftKey: true });
    if (sheet.kind !== 'select') throw new Error('expected a select');
    expect(selectionRange(sheet.selection)).toEqual({
      startRow: 0,
      startCol: 0,
      endRow: projection.rowCount - 1,
      endCol: projection.columnCount - 1,
    });
  });

  it('leaves keys that are not part of the navigation subset alone', () => {
    const projection = m2Projection();
    expect(navigate(projection, 'F6').kind).toBe('none');
    expect(navigate(projection, 'g', { ctrlKey: true }).kind).toBe('none');
    expect(navigate(projection, 'PageUp', { ctrlKey: true }).kind).toBe('none');
  });
});
