/**
 * The Format Cells fill patterns and the previews the dialog draws for them.
 *
 * Excel's Pattern Style drop-down shows each pattern as a small drawing in the
 * chosen fill colours, not as a line of text. This module owns that artwork: an
 * eight-unit tile per pattern, drawn as SVG dots, strokes or a flat wash and
 * encoded as a data URL. Everything here is original and inline, so no asset is
 * fetched and `npm run licence:check` has nothing new to inspect.
 *
 * The grid paints the same tiles on its canvas; the geometry below mirrors
 * `applyPattern` in `src/grid/canvas-renderer.ts` so the preview and the cell
 * agree. A later ticket could share one tile table between the two surfaces.
 */

export interface FillPatternDefinition {
  readonly value: string;
  readonly label: string;
}

/** Every pattern the dialog offers, in the order Excel's Pattern Style list shows. */
export const FILL_PATTERNS: readonly FillPatternDefinition[] = [
  { value: 'none', label: 'None' },
  { value: 'solid', label: 'Solid' },
  { value: 'gray125', label: '12.5% gray' },
  { value: 'gray0625', label: '6.25% gray' },
  { value: 'lightGray', label: 'Light gray' },
  { value: 'mediumGray', label: 'Medium gray' },
  { value: 'darkGray', label: 'Dark gray' },
  { value: 'lightHorizontal', label: 'Light horizontal' },
  { value: 'darkHorizontal', label: 'Dark horizontal' },
  { value: 'lightVertical', label: 'Light vertical' },
  { value: 'darkVertical', label: 'Dark vertical' },
  { value: 'lightDown', label: 'Light downward diagonal' },
  { value: 'darkDown', label: 'Dark downward diagonal' },
  { value: 'lightUp', label: 'Light upward diagonal' },
  { value: 'darkUp', label: 'Dark upward diagonal' },
  { value: 'lightGrid', label: 'Light grid' },
  { value: 'darkGrid', label: 'Dark grid' },
  { value: 'lightTrellis', label: 'Light trellis' },
  { value: 'darkTrellis', label: 'Dark trellis' },
];

/** One point in the tile, or one line segment (`x1, y1, x2, y2`). */
type TilePoint = readonly [number, number];
type TileSegment = readonly [number, number, number, number];

type TileArt =
  | { readonly kind: 'wash'; readonly opacity: number }
  | { readonly kind: 'dots'; readonly points: readonly TilePoint[] }
  | { readonly kind: 'lines'; readonly segments: readonly TileSegment[] };

const TILE_SIZE = 8;

/**
 * The tile art. Coordinates outside the tile are clipped by the SVG view box,
 * exactly as the canvas clips a tile, which is how the two-line diagonals keep
 * their spacing across repeats.
 */
const PATTERN_ART: Readonly<Record<string, TileArt>> = {
  gray125: {
    kind: 'dots',
    points: [
      [1, 1],
      [5, 5],
    ],
  },
  gray0625: { kind: 'dots', points: [[1, 1]] },
  lightGray: { kind: 'wash', opacity: 0.25 },
  mediumGray: { kind: 'wash', opacity: 0.5 },
  darkGray: { kind: 'wash', opacity: 0.75 },
  lightHorizontal: { kind: 'lines', segments: [[0, 6.5, 8, 6.5]] },
  darkHorizontal: {
    kind: 'lines',
    segments: [
      [0, 2.5, 8, 2.5],
      [0, 6.5, 8, 6.5],
    ],
  },
  lightVertical: { kind: 'lines', segments: [[6.5, 0, 6.5, 8]] },
  darkVertical: {
    kind: 'lines',
    segments: [
      [2.5, 0, 2.5, 8],
      [6.5, 0, 6.5, 8],
    ],
  },
  lightDown: { kind: 'lines', segments: [[0.5, 7.5, 7.5, 0.5]] },
  darkDown: {
    kind: 'lines',
    segments: [
      [-2.5, 7.5, 7.5, -2.5],
      [2.5, 11.5, 11.5, 2.5],
    ],
  },
  lightUp: { kind: 'lines', segments: [[0.5, 0.5, 7.5, 7.5]] },
  darkUp: {
    kind: 'lines',
    segments: [
      [-2.5, 0.5, 7.5, 10.5],
      [2.5, -3.5, 11.5, 6.5],
    ],
  },
  lightGrid: {
    kind: 'lines',
    segments: [
      [0, 3.5, 8, 3.5],
      [3.5, 0, 3.5, 8],
    ],
  },
  darkGrid: {
    kind: 'lines',
    segments: [
      [0, 1.5, 8, 1.5],
      [0, 5.5, 8, 5.5],
      [1.5, 0, 1.5, 8],
      [5.5, 0, 5.5, 8],
    ],
  },
  lightTrellis: {
    kind: 'lines',
    segments: [
      [0, 4, 4, 0],
      [4, 8, 8, 4],
    ],
  },
  darkTrellis: {
    kind: 'lines',
    segments: [
      [0, 2, 2, 0],
      [4, 8, 8, 4],
      [0, 6, 2, 8],
      [6, 0, 8, 2],
    ],
  },
};

/** The white a preview falls back to when the pattern has no background. */
export const PATTERN_PREVIEW_BACKGROUND = '#FFFFFF';

/** An imported pattern the dialog does not list previews like the grid paints it. */
const UNKNOWN_PATTERN_ART: TileArt = { kind: 'wash', opacity: 0.5 };

const HEX_COLOR = /^#[0-9a-f]{6}$/i;

/**
 * The colour a tile is allowed to use. Only a plain six-digit hex passes, so a
 * hostile imported value can never inject markup into the SVG.
 */
function previewColor(value: string | null): string | null {
  if (value === null) return null;
  const trimmed = value.trim();
  return HEX_COLOR.test(trimmed) ? trimmed.toUpperCase() : null;
}

function tileArtFor(pattern: string): TileArt | null {
  if (pattern === '' || pattern === 'none' || pattern === 'solid') return null;
  return PATTERN_ART[pattern] ?? UNKNOWN_PATTERN_ART;
}

function tileBody(art: TileArt, color: string): string {
  switch (art.kind) {
    case 'wash':
      return `<rect x='0' y='0' width='${TILE_SIZE}' height='${TILE_SIZE}' fill='${color}' fill-opacity='${art.opacity}'/>`;
    case 'dots':
      return `<g fill='${color}'>${art.points
        .map(([x, y]) => `<rect x='${x}' y='${y}' width='1' height='1'/>`)
        .join('')}</g>`;
    case 'lines':
      return `<g stroke='${color}' stroke-width='1' fill='none'>${art.segments
        .map(([x1, y1, x2, y2]) => `<line x1='${x1}' y1='${y1}' x2='${x2}' y2='${y2}'/>`)
        .join('')}</g>`;
  }
}

/**
 * The SVG source for one repeat of a pattern tile in `color`, or null for
 * `none`, `solid` and a colour that is not a plain hex. Pure and DOM-free, so
 * the preview contract is unit-tested without a browser.
 */
export function patternTileSvg(pattern: string, color: string): string | null {
  const art = tileArtFor(pattern);
  const paint = previewColor(color);
  if (art === null || paint === null) return null;
  return `<svg xmlns='http://www.w3.org/2000/svg' width='${TILE_SIZE}' height='${TILE_SIZE}' viewBox='0 0 ${TILE_SIZE} ${TILE_SIZE}'>${tileBody(art, paint)}</svg>`;
}

/** The CSS `url()` that tiles the preview, or null when the pattern has no texture. */
export function patternTileUrl(pattern: string, color: string): string | null {
  const svg = patternTileSvg(pattern, color);
  return svg === null ? null : `url("data:image/svg+xml,${encodeURIComponent(svg)}")`;
}

export interface FillPatternPreview {
  readonly backgroundColor: string;
  readonly backgroundImage: string | null;
}

/**
 * The colours and texture for one pattern preview. `color` and
 * `backgroundColor` are the dialog's draft values, already normalised by
 * `parseColor`; anything that is not a plain six-digit hex is ignored rather
 * than trusted.
 */
export function fillPatternPreview(
  pattern: string,
  color: string | null,
  backgroundColor: string | null,
): FillPatternPreview {
  const fill = previewColor(color);
  const background = previewColor(backgroundColor) ?? PATTERN_PREVIEW_BACKGROUND;
  if (pattern === 'solid') {
    return { backgroundColor: fill ?? PATTERN_PREVIEW_BACKGROUND, backgroundImage: null };
  }
  if (pattern === 'none' || pattern === '') {
    return { backgroundColor: PATTERN_PREVIEW_BACKGROUND, backgroundImage: null };
  }
  return {
    backgroundColor: background,
    backgroundImage: fill === null ? null : patternTileUrl(pattern, fill),
  };
}
