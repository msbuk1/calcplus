import type { SheetProjection } from '@view/types';
import type { GridCellAddress, GridRange, GridSelection } from './frame';
import { selectionForRange } from './selection';

/**
 * The navigation subset of the shortcut map (command-system §6): arrows,
 * Ctrl+arrows, Home, Ctrl+Home, End, Ctrl+End, Page Up/Down, Shift-extend,
 * Tab/Enter, Ctrl+A, Ctrl+Space, Shift+Space and Ctrl+Shift+Space.
 *
 * Go To (Ctrl+G) and the sheet-tab shortcuts belong to the chrome milestones
 * and are not resolved here.
 */
export interface NavigationInput {
  readonly key: string;
  readonly ctrlKey: boolean;
  readonly shiftKey: boolean;
  readonly altKey: boolean;
  readonly selection: GridSelection;
  readonly projection: SheetProjection;
  /** How many rows fit in the scrolling area, for Page Up and Page Down. */
  readonly pageRows: number;
  readonly pageCols: number;
}

export type NavigationOutcome =
  | { readonly kind: 'move'; readonly selection: GridSelection }
  | { readonly kind: 'select'; readonly selection: GridSelection }
  | { readonly kind: 'none' };

export function resolveNavigation(input: NavigationInput): NavigationOutcome {
  const { key, ctrlKey, shiftKey, altKey, projection, selection } = input;
  if (altKey) return { kind: 'none' };
  const active = selection.focus;

  switch (key) {
    case 'ArrowUp':
    case 'ArrowDown':
    case 'ArrowLeft':
    case 'ArrowRight': {
      const target = ctrlKey ? dataEdge(projection, active, key) : step(projection, active, key);
      if (target === null) return { kind: 'none' };
      if (shiftKey) {
        return {
          kind: 'select',
          selection: { anchor: selection.anchor, focus: target, extent: { kind: 'cells' } },
        };
      }
      return {
        kind: 'move',
        selection: { anchor: target, focus: target, extent: { kind: 'cells' } },
      };
    }
    case 'Tab': {
      // Tab and Shift+Tab move the active cell and keep the focus in the grid.
      const direction = shiftKey ? -1 : 1;
      const target = {
        row: active.row,
        col: nextVisibleColumn(
          projection,
          Math.min(projection.columnCount - 1, Math.max(0, active.col + direction)),
          direction,
        ),
      };
      return {
        kind: 'move',
        selection: { anchor: target, focus: target, extent: { kind: 'cells' } },
      };
    }
    case 'Enter': {
      // Enter and Shift+Enter move the active cell; the editor handles the
      // commit when it is open.
      const direction = shiftKey ? -1 : 1;
      const target = {
        row: nextVisibleRow(
          projection,
          Math.min(projection.rowCount - 1, Math.max(0, active.row + direction)),
          direction,
        ),
        col: active.col,
      };
      return {
        kind: 'move',
        selection: { anchor: target, focus: target, extent: { kind: 'cells' } },
      };
    }
    case 'Home': {
      if (ctrlKey) {
        const target = {
          row: nextVisibleRow(projection, 0, 1),
          col: nextVisibleColumn(projection, 0, 1),
        };
        return shiftAware(selection, target, shiftKey);
      }
      const target = { row: active.row, col: nextVisibleColumn(projection, 0, 1) };
      return shiftAware(selection, target, shiftKey);
    }
    case 'End': {
      if (ctrlKey) {
        const used = projection.usedRange();
        const target = {
          row: nextVisibleRow(projection, used?.endRow ?? projection.rowCount - 1, -1),
          col: nextVisibleColumn(projection, used?.endCol ?? projection.columnCount - 1, -1),
        };
        return shiftAware(selection, target, shiftKey);
      }
      const target = {
        row: active.row,
        col: nextVisibleColumn(projection, endColumn(projection, active), -1),
      };
      return shiftAware(selection, target, shiftKey);
    }
    case 'PageUp':
    case 'PageDown': {
      // Ctrl+Page Up and Ctrl+Page Down are browser-reserved sheet-tab
      // shortcuts; the grid leaves them to the browser.
      if (ctrlKey) return { kind: 'none' };
      const rows = Math.max(1, input.pageRows);
      const direction = key === 'PageUp' ? -1 : 1;
      const row = nextVisibleRow(
        projection,
        clamp(active.row + direction * rows, 0, projection.rowCount - 1),
        direction,
      );
      const target = { row, col: active.col };
      return shiftAware(selection, target, shiftKey);
    }
    case 'a':
    case 'A': {
      if (!ctrlKey) return { kind: 'none' };
      const used = projection.usedRange();
      const range: GridRange =
        used === null
          ? {
              startRow: 0,
              startCol: 0,
              endRow: projection.rowCount - 1,
              endCol: projection.columnCount - 1,
            }
          : used;
      // The active cell stays where it is; the region is the extent.
      return { kind: 'select', selection: selectionForRange(range, active) };
    }
    case ' ': {
      if (ctrlKey && shiftKey) {
        return {
          kind: 'select',
          selection: selectionForRange(
            {
              startRow: 0,
              startCol: 0,
              endRow: projection.rowCount - 1,
              endCol: projection.columnCount - 1,
            },
            active,
          ),
        };
      }
      if (ctrlKey) {
        return {
          kind: 'select',
          selection: selectionForRange(
            {
              startRow: 0,
              startCol: active.col,
              endRow: projection.rowCount - 1,
              endCol: active.col,
            },
            active,
          ),
        };
      }
      if (shiftKey) {
        return {
          kind: 'select',
          selection: selectionForRange(
            {
              startRow: active.row,
              startCol: 0,
              endRow: active.row,
              endCol: projection.columnCount - 1,
            },
            active,
          ),
        };
      }
      return { kind: 'none' };
    }
    default:
      return { kind: 'none' };
  }
}

function shiftAware(
  selection: GridSelection,
  target: GridCellAddress,
  shiftKey: boolean,
): NavigationOutcome {
  if (shiftKey) {
    return {
      kind: 'select',
      selection: { anchor: selection.anchor, focus: target, extent: { kind: 'cells' } },
    };
  }
  return { kind: 'move', selection: { anchor: target, focus: target, extent: { kind: 'cells' } } };
}

function clamp(value: number, low: number, high: number): number {
  return Math.min(high, Math.max(low, value));
}

/**
 * The anchor of the merged range that covers a cell, or the cell itself. The
 * grid never puts the active cell inside a merged range.
 */
export function mergeAnchor(projection: SheetProjection, cell: GridCellAddress): GridCellAddress {
  const merge = projection.mergeAt(cell.row, cell.col);
  if (merge === null) return cell;
  return { row: merge.anchorRow, col: merge.anchorCol };
}

/**
 * The cell after a merged range in the direction of travel. Excel moves past
 * a merged range instead of into the cells it covers.
 */
export function skipMerged(
  projection: SheetProjection,
  cell: GridCellAddress,
  direction: 'up' | 'down' | 'left' | 'right',
): GridCellAddress {
  const merge = projection.mergeAt(cell.row, cell.col);
  if (merge === null) return cell;
  switch (direction) {
    case 'right':
      return {
        row: cell.row,
        col: Math.min(projection.columnCount - 1, merge.anchorCol + merge.colSpan),
      };
    case 'left':
      return { row: cell.row, col: Math.max(0, merge.anchorCol - 1) };
    case 'down':
      return {
        row: Math.min(projection.rowCount - 1, merge.anchorRow + merge.rowSpan),
        col: cell.col,
      };
    default:
      return { row: Math.max(0, merge.anchorRow - 1), col: cell.col };
  }
}

function step(projection: SheetProjection, active: GridCellAddress, key: string): GridCellAddress {
  const from = mergeAnchor(projection, active);
  switch (key) {
    case 'ArrowUp': {
      const target = skipMerged(
        projection,
        { row: nextVisibleRow(projection, from.row - 1, -1), col: from.col },
        'up',
      );
      return { row: nextVisibleRow(projection, target.row, -1), col: target.col };
    }
    case 'ArrowDown': {
      const target = skipMerged(
        projection,
        { row: nextVisibleRow(projection, from.row + 1, 1), col: from.col },
        'down',
      );
      return { row: nextVisibleRow(projection, target.row, 1), col: target.col };
    }
    case 'ArrowLeft': {
      const target = skipMerged(
        projection,
        { row: from.row, col: nextVisibleColumn(projection, from.col - 1, -1) },
        'left',
      );
      return { row: target.row, col: nextVisibleColumn(projection, target.col, -1) };
    }
    default: {
      const target = skipMerged(
        projection,
        { row: from.row, col: nextVisibleColumn(projection, from.col + 1, 1) },
        'right',
      );
      return { row: target.row, col: nextVisibleColumn(projection, target.col, 1) };
    }
  }
}

/**
 * The next row in one direction that is not hidden, clamped to the sheet.
 *
 * A boundary row may itself be hidden. In that case the target is the first
 * or last visible row, never the hidden boundary: an arrow key from the
 * visible edge stays there instead of selecting a collapsed row.
 */
function nextVisibleRow(projection: SheetProjection, row: number, direction: 1 | -1): number {
  let target = row;
  while (target >= 0 && target < projection.rowCount && projection.isRowHidden(target)) {
    target += direction;
  }
  if (target < 0) return projection.rowAt(0);
  if (target >= projection.rowCount) {
    return projection.rowAt(Math.max(0, projection.totalHeight() - 1));
  }
  return target;
}

/** The next column in one direction that is not hidden, clamped to the sheet. */
function nextVisibleColumn(projection: SheetProjection, column: number, direction: 1 | -1): number {
  let target = column;
  while (target >= 0 && target < projection.columnCount && projection.isColumnHidden(target)) {
    target += direction;
  }
  if (target < 0) return projection.columnAt(0);
  if (target >= projection.columnCount) {
    return projection.columnAt(Math.max(0, projection.totalWidth() - 1));
  }
  return target;
}

/**
 * The Ctrl+arrow target: the edge of the block of data, as Excel finds it.
 */
export function dataEdge(
  projection: SheetProjection,
  active: GridCellAddress,
  key: string,
): GridCellAddress {
  const from = mergeAnchor(projection, active);
  switch (key) {
    case 'ArrowUp':
      return {
        row: nextVisibleRow(
          projection,
          edgeInList(projection.columnDataRows(from.col), from.row, -1, projection.rowCount - 1),
          -1,
        ),
        col: from.col,
      };
    case 'ArrowDown':
      return {
        row: nextVisibleRow(
          projection,
          edgeInList(projection.columnDataRows(from.col), from.row, 1, projection.rowCount - 1),
          1,
        ),
        col: from.col,
      };
    case 'ArrowLeft':
      return {
        row: from.row,
        col: nextVisibleColumn(
          projection,
          edgeInList(projection.rowDataColumns(from.row), from.col, -1, projection.columnCount - 1),
          -1,
        ),
      };
    default:
      return {
        row: from.row,
        col: nextVisibleColumn(
          projection,
          edgeInList(projection.rowDataColumns(from.row), from.col, 1, projection.columnCount - 1),
          1,
        ),
      };
  }
}

/**
 * Walk a sorted list of data positions from the current position. When the
 * current cell holds data, the target is the far end of its block; when it
 * does not, the target is the next cell that holds data.
 */
function edgeInList(
  positions: readonly number[],
  current: number,
  direction: 1 | -1,
  limit: number,
): number {
  if (positions.length === 0) return direction === 1 ? limit : 0;
  const index = positions.indexOf(current);
  if (index < 0) {
    if (direction === 1) {
      for (const position of positions) {
        if (position > current) return position;
      }
      return limit;
    }
    for (let position = positions.length - 1; position >= 0; position -= 1) {
      const candidate = positions[position];
      if (candidate !== undefined && candidate < current) return candidate;
    }
    return 0;
  }
  let edge = current;
  let cursor = index;
  while (true) {
    const next = positions[cursor + direction];
    if (next === undefined) break;
    if (next !== edge + direction) break;
    edge = next;
    cursor += direction;
  }
  if (edge !== current) return edge;
  // The cell already sits at the end of its block: Excel jumps to the next
  // cell that holds data, or to the edge of the sheet.
  const neighbour = positions[cursor + direction];
  if (neighbour !== undefined) return neighbour;
  return direction === 1 ? limit : 0;
}

/** The last column of the block of data the cell sits in, or the last column. */
export function endColumn(projection: SheetProjection, active: GridCellAddress): number {
  const columns = projection.rowDataColumns(active.row);
  const index = columns.indexOf(active.col);
  if (index < 0) return projection.columnCount - 1;
  let end = index;
  while (end + 1 < columns.length && (columns[end + 1] ?? -1) === (columns[end] ?? -2) + 1)
    end += 1;
  return columns[end] ?? projection.columnCount - 1;
}
