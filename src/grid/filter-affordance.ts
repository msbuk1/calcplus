import type { CellRect } from './panes';

/**
 * The small AutoFilter button inside a filter header cell (sort-autofilter
 * §3.8: "Each header cell gets a dropdown button when the filter is on").
 *
 * The canvas paints the arrow or funnel and the grid places a transparent DOM
 * button over the same rectangle, so the affordance has a real accessible name
 * and a keyboard target. The size follows the cell and the grid zoom, and is
 * kept inside a very narrow column without disappearing entirely.
 */
export interface FilterButtonRect {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

const BUTTON_SIZE = 16;
const INSET = 1;

export function filterButtonRect(cell: CellRect, zoom: number): FilterButtonRect {
  const preferred = BUTTON_SIZE * zoom;
  const width = Math.max(8, Math.min(preferred, cell.width - INSET * 2));
  const height = Math.max(8, Math.min(preferred, cell.height - INSET * 2));
  return {
    x: cell.x + cell.width - width - INSET,
    y: cell.y + (cell.height - height) / 2,
    width,
    height,
  };
}

/** `true` when a canvas point is inside the filter button of a header cell. */
export function pointInFilterButton(
  cell: CellRect,
  zoom: number,
  pointX: number,
  pointY: number,
): boolean {
  const button = filterButtonRect(cell, zoom);
  return (
    pointX >= button.x &&
    pointX <= button.x + button.width &&
    pointY >= button.y &&
    pointY <= button.y + button.height
  );
}
