import type { SheetProjection } from '@view/types';
import { GRID_METRICS, type GridCellAddress } from './frame';

/**
 * Four-pane freeze geometry.
 *
 * The canvas is viewport sized: the row and column headers sit in strips, the
 * frozen rows and columns in the top-left corner, and the scrolling area fills
 * the rest. Nothing is allocated for the sheet itself.
 */
export interface PaneRect {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

export interface PaneLayout {
  readonly headerWidth: number;
  readonly headerHeight: number;
  readonly scrollbarSize: number;
  /** Width of the frozen columns, in pixels. */
  readonly frozenWidth: number;
  /** Height of the frozen rows, in pixels. */
  readonly frozenHeight: number;
  /** Width of the scrolling area, in pixels. */
  readonly scrollRegionWidth: number;
  /** Height of the scrolling area, in pixels. */
  readonly scrollRegionHeight: number;
  /** Total width of the scrolling columns, in pixels. */
  readonly contentWidth: number;
  /** Total height of the scrolling rows, in pixels. */
  readonly contentHeight: number;
  readonly minScrollX: number;
  readonly minScrollY: number;
  readonly maxScrollX: number;
  readonly maxScrollY: number;
}

export interface CellRect {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

export function computePaneLayout(
  projection: SheetProjection,
  width: number,
  height: number,
): PaneLayout {
  const headerWidth = GRID_METRICS.headerWidth;
  const headerHeight = GRID_METRICS.headerHeight;
  const scrollbarSize = GRID_METRICS.scrollbarSize;
  const bodyWidth = Math.max(0, width - headerWidth - scrollbarSize);
  const bodyHeight = Math.max(0, height - headerHeight - scrollbarSize);
  const frozenWidth = projection.columnX(projection.freezeCols);
  const frozenHeight = projection.rowY(projection.freezeRows);
  const scrollRegionWidth = Math.max(0, bodyWidth - frozenWidth);
  const scrollRegionHeight = Math.max(0, bodyHeight - frozenHeight);
  const minScrollX = projection.columnX(projection.freezeCols);
  const minScrollY = projection.rowY(projection.freezeRows);
  const contentWidth = Math.max(0, projection.totalWidth() - minScrollX);
  const contentHeight = Math.max(0, projection.totalHeight() - minScrollY);
  const maxScrollX = minScrollX + Math.max(0, contentWidth - scrollRegionWidth);
  const maxScrollY = minScrollY + Math.max(0, contentHeight - scrollRegionHeight);
  return {
    headerWidth,
    headerHeight,
    scrollbarSize,
    frozenWidth,
    frozenHeight,
    scrollRegionWidth,
    scrollRegionHeight,
    contentWidth,
    contentHeight,
    minScrollX,
    minScrollY,
    maxScrollX,
    maxScrollY,
  };
}

/** The rectangle a cell occupies on the canvas, in CSS pixels. */
export function cellRect(
  projection: SheetProjection,
  layout: PaneLayout,
  scrollX: number,
  scrollY: number,
  row: number,
  col: number,
): CellRect {
  const frozenRow = row < projection.freezeRows;
  const frozenCol = col < projection.freezeCols;
  const x = frozenCol
    ? layout.headerWidth + projection.columnX(col)
    : layout.headerWidth + layout.frozenWidth + projection.columnX(col) - scrollX;
  const y = frozenRow
    ? layout.headerHeight + projection.rowY(row)
    : layout.headerHeight + layout.frozenHeight + projection.rowY(row) - scrollY;
  return {
    x,
    y,
    width: projection.columnWidth(col),
    height: projection.rowHeight(row),
  };
}

/** The rectangle of a merged range, anchored at its top-left cell. */
export function mergedRect(
  projection: SheetProjection,
  layout: PaneLayout,
  scrollX: number,
  scrollY: number,
  anchorRow: number,
  anchorCol: number,
  rowSpan: number,
  colSpan: number,
): CellRect {
  const start = cellRect(projection, layout, scrollX, scrollY, anchorRow, anchorCol);
  let width = start.width;
  let height = start.height;
  for (let col = 1; col < colSpan; col += 1) {
    width += projection.columnWidth(anchorCol + col);
  }
  for (let row = 1; row < rowSpan; row += 1) {
    height += projection.rowHeight(anchorRow + row);
  }
  return { x: start.x, y: start.y, width, height };
}

/**
 * The column under a point in the column header strip, or `null` when the
 * point is outside the strip.
 */
export function columnAtHeader(
  projection: SheetProjection,
  layout: PaneLayout,
  pointX: number,
  scrollX: number,
): number | null {
  if (pointX < layout.headerWidth || pointX > layout.headerWidth + layout.scrollRegionWidth) {
    return null;
  }
  if (pointX < layout.headerWidth + layout.frozenWidth) {
    return projection.columnAt(pointX - layout.headerWidth);
  }
  return projection.columnAt(pointX - layout.headerWidth - layout.frozenWidth + scrollX);
}

/**
 * The row under a point in the row header strip, or `null` when the point is
 * outside the strip.
 */
export function rowAtHeader(
  projection: SheetProjection,
  layout: PaneLayout,
  pointY: number,
  scrollY: number,
): number | null {
  if (pointY < layout.headerHeight || pointY > layout.headerHeight + layout.scrollRegionHeight) {
    return null;
  }
  if (pointY < layout.headerHeight + layout.frozenHeight) {
    return projection.rowAt(pointY - layout.headerHeight);
  }
  return projection.rowAt(pointY - layout.headerHeight - layout.frozenHeight + scrollY);
}

/** Clamp a scroll offset to the range the sheet allows. */
export function clampScroll(
  layout: PaneLayout,
  scrollX: number,
  scrollY: number,
): {
  x: number;
  y: number;
} {
  return {
    x: Math.min(layout.maxScrollX, Math.max(layout.minScrollX, scrollX)),
    y: Math.min(layout.maxScrollY, Math.max(layout.minScrollY, scrollY)),
  };
}

/**
 * Scroll the minimum amount that brings a cell into the scrolling area. The
 * frozen panes always stay visible, so a frozen cell never scrolls.
 */
export function scrollToReveal(
  projection: SheetProjection,
  layout: PaneLayout,
  scrollX: number,
  scrollY: number,
  cell: GridCellAddress,
): { x: number; y: number } {
  let nextX = scrollX;
  let nextY = scrollY;
  if (cell.col >= projection.freezeCols) {
    const left = projection.columnX(cell.col);
    const right = left + projection.columnWidth(cell.col);
    const regionLeft = layout.minScrollX;
    const regionRight = layout.minScrollX + layout.scrollRegionWidth;
    if (left < regionLeft) nextX = left;
    else if (right > regionRight) nextX = right - layout.scrollRegionWidth;
  }
  if (cell.row >= projection.freezeRows) {
    const top = projection.rowY(cell.row);
    const bottom = top + projection.rowHeight(cell.row);
    const regionTop = layout.minScrollY;
    const regionBottom = layout.minScrollY + layout.scrollRegionHeight;
    if (top < regionTop) nextY = top;
    else if (bottom > regionBottom) nextY = bottom - layout.scrollRegionHeight;
  }
  return clampScroll(layout, nextX, nextY);
}
