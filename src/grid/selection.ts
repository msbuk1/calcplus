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

/** Move the extension end of a selection, keeping the cell it began at. */
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

/**
 * The active cell is the one the gesture started at — the anchor (ticket 31).
 *
 * A click, a drag, a Shift extension and Ctrl+A all keep it where the gesture
 * began: the drag draws a range whose focus is the far corner, so the focus is
 * the extension end, never the active cell. A plain arrow collapses the
 * selection onto a new anchor and the active cell moves with it. Every surface
 * that reads "the active cell" — the name box's command context, the formula
 * bar, in-cell editing, formula pointing and the style readouts — reads this.
 */
export function activeCell(selection: GridSelection): GridCellAddress {
  return selection.anchor;
}

/** True when the selection covers exactly one cell. */
export function isSingleCell(selection: GridSelection): boolean {
  return (
    selection.extent.kind === 'cells' &&
    selection.anchor.row === selection.focus.row &&
    selection.anchor.col === selection.focus.col
  );
}

/**
 * Shift a rectangle by whole rows and columns without leaving the sheet
 * (ticket 31 B: the copy marquee clamps at the edge, it does not wrap).
 * A rectangle that already spans the sheet on an axis cannot move on it.
 */
export function moveRange(
  range: GridRange,
  rowDelta: number,
  colDelta: number,
  rowCount: number,
  columnCount: number,
): GridRange {
  const rowShift = clampedShift(range.startRow, range.endRow, rowDelta, rowCount);
  const colShift = clampedShift(range.startCol, range.endCol, colDelta, columnCount);
  if (rowShift === 0 && colShift === 0) return range;
  return {
    startRow: range.startRow + rowShift,
    startCol: range.startCol + colShift,
    endRow: range.endRow + rowShift,
    endCol: range.endCol + colShift,
  };
}

function clampedShift(start: number, end: number, delta: number, count: number): number {
  if (delta === 0) return 0;
  if (delta > 0) return Math.min(delta, Math.max(0, count - 1 - end));
  return -Math.min(-delta, Math.max(0, start));
}
