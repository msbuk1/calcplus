import type { MergeSpan, SheetProjection } from '@view/types';
import { GRID_METRICS, type GridCellAddress, type GridRange, type GridSelection } from './frame';
import type { DimensionAxis } from './resize';

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
  // Excel scales the headers with the cells and removes their strips entirely
  // when headings are hidden; the scrollbars keep their physical size so they
  // stay usable at low zoom.
  const headerWidth = projection.showHeadings ? GRID_METRICS.headerWidth * projection.zoom : 0;
  const headerHeight = projection.showHeadings ? GRID_METRICS.headerHeight * projection.zoom : 0;
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

/** The rows and columns a pane shows, in sheet indices. */
export interface PaneBounds {
  readonly startRow: number;
  readonly endRow: number;
  readonly startCol: number;
  readonly endCol: number;
}

/**
 * Whether a merged range reaches past the pane that holds its anchor cell.
 *
 * A pane only shows the rows and columns inside its own bounds, so a range that
 * spans the freeze divider starts in one pane and ends in another. Painting it
 * inside the starting pane would cut it at the divider.
 */
export function mergeSpillsPane(merge: MergeSpan, pane: PaneBounds): boolean {
  return (
    merge.anchorRow < pane.startRow ||
    merge.anchorRow + merge.rowSpan - 1 > pane.endRow ||
    merge.anchorCol < pane.startCol ||
    merge.anchorCol + merge.colSpan - 1 > pane.endCol
  );
}

/**
 * The area a pane may paint into, in canvas pixels. Along an axis the pane does
 * not scroll it owns the whole body rather than its own rectangle, so a merged
 * range that reaches across the freeze divider is painted whole. Along an axis
 * it does scroll, the pane rectangle is the hard edge: the frozen strips and
 * the panes beside it are already on the canvas.
 */
export function panePaintRect(
  layout: PaneLayout,
  frozen: { readonly frozenRows: boolean; readonly frozenCols: boolean },
): CellRect {
  const bodyWidth = Math.max(0, layout.frozenWidth + layout.scrollRegionWidth);
  const bodyHeight = Math.max(0, layout.frozenHeight + layout.scrollRegionHeight);
  return {
    x: frozen.frozenCols ? layout.headerWidth : layout.headerWidth + layout.frozenWidth,
    y: frozen.frozenRows ? layout.headerHeight : layout.headerHeight + layout.frozenHeight,
    width: frozen.frozenCols ? bodyWidth : layout.scrollRegionWidth,
    height: frozen.frozenRows ? bodyHeight : layout.scrollRegionHeight,
  };
}

/** The overlap of two rectangles, with a zero width or height when they miss. */
export function intersectRects(a: CellRect, b: CellRect): CellRect {
  const x = Math.max(a.x, b.x);
  const y = Math.max(a.y, b.y);
  return {
    x,
    y,
    width: Math.max(0, Math.min(a.x + a.width, b.x + b.width) - x),
    height: Math.max(0, Math.min(a.y + a.height, b.y + b.height) - y),
  };
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
  if (
    pointX < layout.headerWidth ||
    pointX > layout.headerWidth + layout.frozenWidth + layout.scrollRegionWidth
  ) {
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
  if (
    pointY < layout.headerHeight ||
    pointY > layout.headerHeight + layout.frozenHeight + layout.scrollRegionHeight
  ) {
    return null;
  }
  if (pointY < layout.headerHeight + layout.frozenHeight) {
    return projection.rowAt(pointY - layout.headerHeight);
  }
  return projection.rowAt(pointY - layout.headerHeight - layout.frozenHeight + scrollY);
}

/** A row or column whose leading or trailing header edge is under the pointer. */
export interface HeaderResizeTarget {
  readonly axis: DimensionAxis;
  readonly index: number;
}

/**
 * The whole-row or whole-column range a header selection covers, or `null`
 * when the selection is not a header selection on that axis. A whole-sheet
 * selection is excluded: dragging one header edge must not resize all 16,384
 * columns or 1,048,576 rows.
 */
export function headerSelectionRange(
  projection: SheetProjection,
  selection: GridSelection,
  axis: DimensionAxis,
): GridRange | null {
  if (selection.extent.kind !== 'range') return null;
  const range = selection.extent.range;
  if (axis === 'column') {
    if (range.startRow !== 0 || range.endRow !== projection.rowCount - 1) return null;
    if (range.startCol === 0 && range.endCol === projection.columnCount - 1) return null;
  } else {
    if (range.startCol !== 0 || range.endCol !== projection.columnCount - 1) return null;
    if (range.startRow === 0 && range.endRow === projection.rowCount - 1) return null;
  }
  return range;
}

/**
 * The visible rows or columns a header selection covers, in order. Hidden
 * members are skipped so a resize or autofit never changes their stored size
 * or reveals them.
 */
export function headerSelectionIndices(
  projection: SheetProjection,
  selection: GridSelection,
  axis: DimensionAxis,
): readonly number[] | null {
  const range = headerSelectionRange(projection, selection, axis);
  if (range === null) return null;
  const start = axis === 'column' ? range.startCol : range.startRow;
  const end = axis === 'column' ? range.endCol : range.endRow;
  const indices: number[] = [];
  for (let index = start; index <= end; index += 1) {
    const hidden =
      axis === 'column' ? projection.isColumnHidden(index) : projection.isRowHidden(index);
    if (!hidden) indices.push(index);
  }
  return indices.length === 0 ? null : indices;
}

/**
 * The header edge under a pointer position, or `null` when the pointer is
 * elsewhere. The gutter is a few screen pixels wide on either side of a
 * visible boundary; a hidden row or column has no edge and is never a target.
 */
export function headerResizeTarget(
  projection: SheetProjection,
  layout: PaneLayout,
  pointX: number,
  pointY: number,
  scrollX: number,
  scrollY: number,
): HeaderResizeTarget | null {
  if (pointY < layout.headerHeight && pointX >= layout.headerWidth) {
    const col = columnAtHeader(projection, layout, pointX, scrollX);
    if (col === null) return null;
    const index = columnResizeEdge(projection, layout, pointX, scrollX, col);
    return index === null ? null : { axis: 'column', index };
  }
  if (pointX < layout.headerWidth && pointY >= layout.headerHeight) {
    const row = rowAtHeader(projection, layout, pointY, scrollY);
    if (row === null) return null;
    const index = rowResizeEdge(projection, layout, pointY, scrollY, row);
    return index === null ? null : { axis: 'row', index };
  }
  return null;
}

/** The header gutter is measured in screen pixels, like Excel's. */
const RESIZE_GUTTER_PIXELS = 4;

function columnResizeEdge(
  projection: SheetProjection,
  layout: PaneLayout,
  pointX: number,
  scrollX: number,
  col: number,
): number | null {
  const width = projection.columnWidth(col);
  if (width <= 0) return null;
  const x =
    col < projection.freezeCols
      ? layout.headerWidth + projection.columnX(col)
      : layout.headerWidth + layout.frozenWidth + projection.columnX(col) - scrollX;
  const edges: { readonly index: number; readonly position: number }[] = [];
  const start = projection.columnX(col);
  if (start > 0) edges.push({ index: projection.columnAt(start - 1), position: x });
  edges.push({ index: col, position: x + width });
  return closestEdge(
    pointX,
    edges,
    (index) => projection.columnWidth(index),
    layout.headerWidth,
    layout.headerWidth + layout.frozenWidth + layout.scrollRegionWidth,
  );
}

function rowResizeEdge(
  projection: SheetProjection,
  layout: PaneLayout,
  pointY: number,
  scrollY: number,
  row: number,
): number | null {
  const height = projection.rowHeight(row);
  if (height <= 0) return null;
  const y =
    row < projection.freezeRows
      ? layout.headerHeight + projection.rowY(row)
      : layout.headerHeight + layout.frozenHeight + projection.rowY(row) - scrollY;
  const edges: { readonly index: number; readonly position: number }[] = [];
  const start = projection.rowY(row);
  if (start > 0) edges.push({ index: projection.rowAt(start - 1), position: y });
  edges.push({ index: row, position: y + height });
  return closestEdge(
    pointY,
    edges,
    (index) => projection.rowHeight(index),
    layout.headerHeight,
    layout.headerHeight + layout.frozenHeight + layout.scrollRegionHeight,
  );
}

/** The closest visible edge within the gutter, preferring the earlier index on a tie. */
function closestEdge(
  point: number,
  edges: readonly { readonly index: number; readonly position: number }[],
  sizeOf: (index: number) => number,
  minimumPosition: number,
  maximumPosition: number,
): number | null {
  let best: { readonly index: number; readonly distance: number } | null = null;
  for (const edge of edges) {
    if (edge.position < minimumPosition || edge.position > maximumPosition) continue;
    if (sizeOf(edge.index) <= 0) continue;
    const distance = Math.abs(point - edge.position);
    if (distance > RESIZE_GUTTER_PIXELS) continue;
    if (
      best === null ||
      distance < best.distance ||
      (distance === best.distance && edge.index < best.index)
    ) {
      best = { index: edge.index, distance };
    }
  }
  return best === null ? null : best.index;
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
    // The scrolling viewport starts at the current scroll offset, not at the
    // frozen boundary: a cell to the left of the viewport must scroll back
    // into view just as one to the right must scroll forward.
    const viewLeft = scrollX;
    const viewRight = scrollX + layout.scrollRegionWidth;
    if (left < viewLeft) nextX = left;
    else if (right > viewRight) nextX = right - layout.scrollRegionWidth;
  }
  if (cell.row >= projection.freezeRows) {
    const top = projection.rowY(cell.row);
    const bottom = top + projection.rowHeight(cell.row);
    const viewTop = scrollY;
    const viewBottom = scrollY + layout.scrollRegionHeight;
    if (top < viewTop) nextY = top;
    else if (bottom > viewBottom) nextY = bottom - layout.scrollRegionHeight;
  }
  return clampScroll(layout, nextX, nextY);
}
