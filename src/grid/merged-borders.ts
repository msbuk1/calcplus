import type { MergeSpan, ProjectedStyle, SheetProjection } from '@view/types';

type BorderStyle = ProjectedStyle['border'];
type BorderEdge = 'left' | 'right' | 'top' | 'bottom';

/**
 * The border of a merged range.
 *
 * A merge paints one rectangle, but the file may store an edge on any
 * constituent cell. Excel's merge formatting keeps the outer edges of every
 * cell in the merged block, so each side is resolved from the cells that touch
 * that side. The anchor remains the fallback, which keeps a border applied to
 * the merged cell itself visible on all four sides.
 */
export function mergedBorderStyle(projection: SheetProjection, merge: MergeSpan): BorderStyle {
  const anchor = projection.cellAt(merge.anchorRow, merge.anchorCol)?.style.border ?? null;
  return {
    left: pickEdge(projection, edgeCells(merge, 'left'), 'left', anchor?.left ?? noneEdge()),
    right: pickEdge(projection, edgeCells(merge, 'right'), 'right', anchor?.right ?? noneEdge()),
    top: pickEdge(projection, edgeCells(merge, 'top'), 'top', anchor?.top ?? noneEdge()),
    bottom: pickEdge(
      projection,
      edgeCells(merge, 'bottom'),
      'bottom',
      anchor?.bottom ?? noneEdge(),
    ),
    diagonal: anchor?.diagonal ?? noneEdge(),
    diagonalUp: anchor?.diagonalUp ?? false,
    diagonalDown: anchor?.diagonalDown ?? false,
  };
}

function* edgeCells(
  merge: MergeSpan,
  edge: BorderEdge,
): Generator<{ readonly row: number; readonly col: number }> {
  const lastRow = merge.anchorRow + merge.rowSpan - 1;
  const lastCol = merge.anchorCol + merge.colSpan - 1;
  if (edge === 'left' || edge === 'right') {
    const col = edge === 'left' ? merge.anchorCol : lastCol;
    for (let row = merge.anchorRow; row <= lastRow; row += 1) yield { row, col };
    return;
  }
  const row = edge === 'top' ? merge.anchorRow : lastRow;
  for (let col = merge.anchorCol; col <= lastCol; col += 1) yield { row, col };
}

function pickEdge(
  projection: SheetProjection,
  cells: Iterable<{ readonly row: number; readonly col: number }>,
  edge: BorderEdge,
  fallback: ProjectedStyle['border'][BorderEdge],
): ProjectedStyle['border'][BorderEdge] {
  for (const { row, col } of cells) {
    const border = projection.cellAt(row, col)?.style.border;
    const candidate = border?.[edge];
    if (candidate !== undefined && candidate.style !== 'none' && candidate.style !== '') {
      return candidate;
    }
  }
  return fallback;
}

function noneEdge(): ProjectedStyle['border'][BorderEdge] {
  return { style: 'none', color: null, weight: null };
}
