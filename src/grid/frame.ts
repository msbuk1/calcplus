import type { SheetProjection } from '@view/types';

export const GRID_METRICS = Object.freeze({
  headerWidth: 48,
  headerHeight: 24,
  rowHeight: 20,
  columnWidth: 64,
  scrollbarSize: 12,
  scrollbarThumb: 8,
});

export interface GridCellAddress {
  readonly row: number;
  readonly col: number;
}

/** A rectangular range, in grid coordinates. */
export interface GridRange {
  readonly startRow: number;
  readonly startCol: number;
  readonly endRow: number;
  readonly endCol: number;
}

/**
 * The selection is an anchor and a focus. The anchor stays where the
 * selection started; the focus is the active cell and moves with the
 * arrow keys.
 */
/**
 * The extent of a selection. `cells` is the rectangle between the anchor and
 * the focus; `range` is a whole row, a whole column, the whole sheet or the
 * used region, which a header click or a shortcut produces. A `range`
 * selection keeps the active cell where the user left it, so selecting a
 * column never scrolls to the far edge of the sheet.
 */
export type SelectionExtent =
  { readonly kind: 'cells' } | { readonly kind: 'range'; readonly range: GridRange };

export interface GridSelection {
  readonly anchor: GridCellAddress;
  readonly focus: GridCellAddress;
  readonly extent: SelectionExtent;
}

export interface GridRenderState {
  readonly projection: SheetProjection;
  readonly selection: GridSelection;
  /** The range a formula edit is pointing at, when it is not the selection. */
  readonly pointing?: GridRange | null;
  readonly scrollX: number;
  readonly scrollY: number;
  readonly width: number;
  readonly height: number;
  readonly dpr: number;
}

/** The renderer seam. A DOM-cell renderer can replace the canvas without touching the model. */
export interface RenderFrame {
  resize(width: number, height: number, dpr: number): void;
  render(state: GridRenderState): void;
  destroy(): void;
}
