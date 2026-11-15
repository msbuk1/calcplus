import type { PaneLayout } from './panes';

export interface GridDebugRect {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

/**
 * A small published view of the grid viewport. The app exposes it through the
 * debug API so tests and benchmarks can address cells without duplicating the
 * layout maths.
 */
export interface GridDebugState {
  layout: PaneLayout | null;
  scrollX: number;
  scrollY: number;
  /** Where the fill handle was painted, in canvas coordinates. */
  fillHandle: GridDebugRect | null;
}

export const gridDebugState: GridDebugState = {
  layout: null,
  scrollX: 0,
  scrollY: 0,
  fillHandle: null,
};
