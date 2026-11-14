import type { SheetProjection } from '@view/types';

export const GRID_METRICS = Object.freeze({
  headerWidth: 48,
  headerHeight: 24,
  rowHeight: 22,
  columnWidth: 104,
});

export interface GridCellAddress {
  readonly row: number;
  readonly col: number;
}

export interface GridRenderState {
  readonly projection: SheetProjection;
  readonly active: GridCellAddress;
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
