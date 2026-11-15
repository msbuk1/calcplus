import type { GridCellAddress, GridRange, GridSelection, SelectionExtent } from './frame';

export function createSelection(cell: GridCellAddress): GridSelection {
  return { anchor: cell, focus: cell, extent: { kind: 'cells' } };
}

export function selectionRange(selection: GridSelection): GridRange {
  if (selection.extent.kind === 'range') return selection.extent.range;
  const { anchor, focus } = selection;
  return {
    startRow: Math.min(anchor.row, focus.row),
    startCol: Math.min(anchor.col, focus.col),
    endRow: Math.max(anchor.row, focus.row),
    endCol: Math.max(anchor.col, focus.col),
  };
}

export function selectionContains(selection: GridSelection, row: number, col: number): boolean {
  const range = selectionRange(selection);
  return (
    row >= range.startRow && row <= range.endRow && col >= range.startCol && col <= range.endCol
  );
}

/** Move the selection to a new active cell, keeping the anchor. */
export function moveSelection(selection: GridSelection, cell: GridCellAddress): GridSelection {
  return { anchor: selection.anchor, focus: cell, extent: { kind: 'cells' } };
}

/** Move the anchor and the focus together, as a plain click does. */
export function collapseSelection(cell: GridCellAddress): GridSelection {
  return createSelection(cell);
}

/** Extend the selection from the anchor to a new focus cell. */
export function extendSelection(selection: GridSelection, cell: GridCellAddress): GridSelection {
  return { anchor: selection.anchor, focus: cell, extent: { kind: 'cells' } };
}

/**
 * Select a rectangular range. The focus stays where the user left it unless
 * one is given, so a whole-row, whole-column or region selection does not
 * move the active cell to the far edge of the sheet.
 */
export function selectionForRange(range: GridRange, focus?: GridCellAddress): GridSelection {
  const extent: SelectionExtent = { kind: 'range', range };
  const active = focus ?? { row: range.endRow, col: range.endCol };
  return { anchor: active, focus: active, extent };
}

/** The active cell of a selection. */
export function activeCell(selection: GridSelection): GridCellAddress {
  return selection.focus;
}

/** True when the selection covers exactly one cell. */
export function isSingleCell(selection: GridSelection): boolean {
  return (
    selection.extent.kind === 'cells' &&
    selection.anchor.row === selection.focus.row &&
    selection.anchor.col === selection.focus.col
  );
}
