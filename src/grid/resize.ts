import {
  columnWidthCharsToPixels,
  columnWidthPixelsToChars,
  pixelsToPoints,
  pointsToPixels,
} from '@view/geometry';
import { MAX_COLUMN_WIDTH_CHARACTERS, MAX_ROW_HEIGHT_POINTS } from './autofit';

/**
 * Row-height and column-width drag arithmetic.
 *
 * The pointer moves in screen pixels at the live zoom; the model stores row
 * heights in points and column widths in Excel character units. A resize keeps
 * the row or column visible: hiding is a separate v2 feature, so dragging to
 * zero must not turn into an accidental hidden dimension.
 */

export type DimensionAxis = 'row' | 'column';

export interface DimensionResize {
  /** The size to paint at 100%; the projection scales it by the display zoom. */
  readonly pixels: number;
  /** The same size in the unit the model stores: points or characters. */
  readonly modelValue: number;
}

const MIN_COLUMN_WIDTH_PIXELS = 8;
const MIN_ROW_HEIGHT_PIXELS = 4;

/**
 * Translate a pointer delta into a new dimension size. `originPixels` is the
 * size at 100%, so the delta is divided by the display zoom and the result
 * stays anchored to the edge under the pointer at any zoom level.
 */
export function resizeDimension(
  axis: DimensionAxis,
  originPixels: number,
  deltaPixels: number,
  zoom: number,
): DimensionResize {
  const scale = Number.isFinite(zoom) && zoom > 0 ? zoom : 1;
  const requested = originPixels + deltaPixels / scale;
  if (axis === 'column') {
    const pixels = Math.round(
      clamp(
        requested,
        MIN_COLUMN_WIDTH_PIXELS,
        columnWidthCharsToPixels(MAX_COLUMN_WIDTH_CHARACTERS),
      ),
    );
    return {
      pixels,
      modelValue: Math.min(MAX_COLUMN_WIDTH_CHARACTERS, columnWidthPixelsToChars(pixels)),
    };
  }
  const pixels = Math.round(
    clamp(requested, MIN_ROW_HEIGHT_PIXELS, pointsToPixels(MAX_ROW_HEIGHT_POINTS)),
  );
  return {
    pixels,
    modelValue: Math.min(MAX_ROW_HEIGHT_POINTS, pixelsToPoints(pixels)),
  };
}

function clamp(value: number, minimum: number, maximum: number): number {
  return Math.min(maximum, Math.max(minimum, value));
}
