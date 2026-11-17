import type { ProjectedCell, SheetProjection } from '@view/types';
import {
  columnWidthCharsToPixels,
  columnWidthPixelsToChars,
  pixelsToPoints,
  pointsToPixels,
} from '@view/geometry';
import {
  fontFor,
  fontSizePixels,
  horizontalFor,
  lineHeightFor,
  textInsets,
  wrapTextLines,
  type TextMeasure,
} from './text-layout';

/**
 * Autofit sizing for rows and columns.
 *
 * The grid owns the measurement because only the renderer knows the font
 * metrics. A resize or autofit command applies the result as a document
 * dimension; these functions never mutate the workbook. Widths come back in
 * pixels and in Excel character units, heights in pixels and points, so the
 * caller can store what the model expects (workbook-model §3).
 */

/** Excel's maximum row height is 409 points. */
export const MAX_ROW_HEIGHT_POINTS = 409;
/** Excel's maximum column width is 255 characters. */
export const MAX_COLUMN_WIDTH_CHARACTERS = 255;
const MIN_ROW_HEIGHT_PIXELS = 20;

export interface AutofitResult {
  /** The size to paint at 100%; the grid scales it by the display zoom. */
  readonly pixels: number;
  /** The same size in the unit the model stores: points or characters. */
  readonly modelValue: number;
}

/** One selected row or column and the content size measured for it. */
export interface AutofitTarget {
  readonly index: number;
  readonly result: AutofitResult;
}

/** The fitted width of one column, in pixels, and in Excel characters. */
export function autofitColumnWidth(
  projection: SheetProjection,
  column: number,
  measure: TextMeasure,
): AutofitResult {
  let widest = 0;
  for (const row of projection.columnDataRows(column)) {
    if (projection.isRowHidden(row)) continue;
    const cell = projection.cellAt(row, column);
    if (cell === null || cell.displayText === '') continue;
    const merge = cell.merge;
    if (merge !== null && (merge.anchorRow !== row || merge.anchorCol !== column)) continue;
    const insets = textInsets(cell.style, horizontalFor(cell.style, cell.value));
    const content = textWidth(cell, measure) + insets.left + insets.right;
    widest = Math.max(widest, content);
  }
  if (widest === 0) {
    const fallback = projection.defaultColumnWidthChars ?? 8.43;
    return {
      pixels: columnWidthCharsToPixels(fallback),
      modelValue: fallback,
    };
  }
  const pixels = Math.min(Math.ceil(widest), columnWidthCharsToPixels(MAX_COLUMN_WIDTH_CHARACTERS));
  return {
    pixels,
    modelValue: Math.min(MAX_COLUMN_WIDTH_CHARACTERS, columnWidthPixelsToChars(pixels)),
  };
}

/** The fitted height of one row, in pixels, and in points. */
export function autofitRowHeight(
  projection: SheetProjection,
  row: number,
  measure: TextMeasure,
): AutofitResult {
  const minimum = defaultRowHeightPixels(projection);
  let tallest = minimum;
  for (const column of projection.rowDataColumns(row)) {
    if (projection.isColumnHidden(column)) continue;
    const cell = projection.cellAt(row, column);
    if (cell === null || cell.displayText === '') continue;
    const merge = cell.merge;
    if (merge !== null && (merge.anchorRow !== row || merge.anchorCol !== column)) continue;
    const width = cellWidth(projection, row, column, cell);
    const insets = textInsets(cell.style, horizontalFor(cell.style, cell.value));
    const available = Math.max(1, width - insets.left - insets.right);
    const font = fontFor(cell.style);
    const measureInFont = (text: string): number => measure(text, font);
    const lines = cell.style.alignment.wrapText
      ? wrapTextLines(cell.displayText, available, measureInFont).length
      : 1;
    const lineHeight = lineHeightFor(fontSizePixels(cell.style));
    tallest = Math.max(tallest, Math.ceil(lines * lineHeight + (lines === 1 ? 1 : 2)));
  }
  const pixels = Math.min(Math.max(tallest, minimum), pointsToPixels(MAX_ROW_HEIGHT_POINTS));
  return {
    pixels,
    modelValue: Math.min(MAX_ROW_HEIGHT_POINTS, pixelsToPoints(pixels)),
  };
}

function defaultRowHeightPixels(projection: SheetProjection): number {
  return projection.defaultRowHeightPoints === null
    ? MIN_ROW_HEIGHT_PIXELS
    : pointsToPixels(projection.defaultRowHeightPoints);
}

/** The width the content needs, without the two cell paddings. */
function textWidth(cell: ProjectedCell, measure: TextMeasure): number {
  const font = fontFor(cell.style);
  if (!cell.style.alignment.wrapText) return measure(cell.displayText, font);
  // A wrapped cell autofits to its longest word, not its whole text.
  let widest = 0;
  for (const line of cell.displayText.split(/\r?\n/)) {
    for (const word of line.split(/\s+/)) {
      if (word === '') continue;
      widest = Math.max(widest, measure(word, font));
    }
  }
  return widest;
}

function cellWidth(
  projection: SheetProjection,
  row: number,
  column: number,
  cell: ProjectedCell,
): number {
  const merge = cell.merge;
  if (merge === null || merge.anchorRow !== row || merge.anchorCol !== column) {
    return projection.columnWidth(column) / projection.zoom;
  }
  let width = 0;
  for (let offset = 0; offset < merge.colSpan; offset += 1) {
    width += projection.columnWidth(merge.anchorCol + offset) / projection.zoom;
  }
  return width;
}
