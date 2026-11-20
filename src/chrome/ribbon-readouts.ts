import type { GridCellAddress } from '@grid/frame';
import type { ProjectedCell, SheetProjection } from '@view/types';

/**
 * The ribbon's current-style readouts (ui-chrome §4.1). Every readout previews
 * the active cell, which is the selection's anchor (ticket 31), and the colour
 * pickers show the same cell in their Current swatch.
 *
 * A merged cell is one cell: Excel keeps its formatting on the merge's anchor,
 * so a readout for any cell inside a merge shows the anchor's style. The
 * projection answers `cellAt` per stored cell and does not fold the merge in,
 * so this module resolves the anchor itself (ticket 40). A covered cell with no
 * stored style therefore still reads the merge's fill and font.
 */
export function readoutCell(projection: SheetProjection, active: GridCellAddress): GridCellAddress {
  const merge = projection.mergeAt(active.row, active.col);
  if (merge === null) return active;
  return { row: merge.anchorRow, col: merge.anchorCol };
}

/** The cell a ribbon readout previews, or `null` when there is no sheet. */
export function readoutProjectedCell(
  projection: SheetProjection | null,
  active: GridCellAddress,
): ProjectedCell | null {
  if (projection === null) return null;
  const cell = readoutCell(projection, active);
  return projection.cellAt(cell.row, cell.col);
}

/**
 * The ribbon control values, keyed by the control ids in the command map.
 * A blank value means automatic or no fill; the colour picker's Current swatch
 * renders its No Fill and Automatic marks for it.
 */
export function ribbonControlValues(
  projection: SheetProjection | null,
  active: GridCellAddress,
): Readonly<Record<string, string>> {
  const projected = readoutProjectedCell(projection, active);
  return {
    'font-family': projected?.style.font.name ?? 'Calibri',
    'font-size': projected === null ? '11' : String(projected.style.font.size),
    'number-format': projected?.style.numberFormat ?? 'General',
    'font-color': projected?.style.font.color ?? '',
    'fill-color': projected?.style.fill.color ?? '',
  };
}
