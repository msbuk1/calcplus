import type { PrintIconItem } from './types';

/**
 * Renderer-neutral paint geometry for the print plan (print-pdf §2).
 *
 * The SVG preview, the PDFKit writer and the 300 DPI canvas fallback all paint
 * the same pattern tiles and the same seven conditional-formatting icon
 * families. The geometry lives here once, in page points, so the three
 * renderers can only differ in how they execute the same shapes — not in what
 * they draw.
 *
 * The pattern tile is the grid's 8x8 CSS-pixel tile (`applyPattern` in
 * `src/grid/canvas-renderer.ts`) converted to 6 points: 8 px at 96 dpi is 6 pt.
 * The icon geometry mirrors `paintCfIcon` in `src/grid/cf-paint.ts`; sizes and
 * fractions are copied from the reference, so an icon prints at Excel's
 * 12-device-pixel size with the same proportions.
 */

/** OOXML's 8x8 pattern tile in points (8 CSS px at 96 dpi). */
export const PATTERN_TILE = 6;

/** One primitive inside the pattern tile, relative to the tile's top-left. */
export type PatternTileOp =
  | {
      readonly kind: 'fill';
      readonly x: number;
      readonly y: number;
      readonly width: number;
      readonly height: number;
      readonly opacity: number;
    }
  | {
      readonly kind: 'stroke';
      readonly x1: number;
      readonly y1: number;
      readonly x2: number;
      readonly y2: number;
      readonly opacity: number;
    };

/**
 * The tile's paint for one OOXML pattern name, in points. An unknown name
 * paints a half-opacity block so the colour stays visible instead of silently
 * disappearing (the grid makes the same choice).
 */
export function patternTileOps(pattern: string): readonly PatternTileOp[] {
  const box = (opacity: number, x = 0, y = 0, width = PATTERN_TILE, height = PATTERN_TILE) =>
    ({ kind: 'fill', x, y, width, height, opacity }) as const;
  const line = (x1: number, y1: number, x2: number, y2: number) =>
    ({ kind: 'stroke', x1, y1, x2, y2, opacity: 1 }) as const;
  switch (pattern) {
    case 'gray125':
      return [box(1, 0.75, 0.75, 0.75, 0.75), box(1, 3.75, 3.75, 0.75, 0.75)];
    case 'gray0625':
      return [box(1, 0.75, 0.75, 0.75, 0.75)];
    case 'lightGray':
      return [box(0.25)];
    case 'mediumGray':
      return [box(0.5)];
    case 'darkGray':
      return [box(0.75)];
    case 'lightHorizontal':
      return [line(0, 4.875, PATTERN_TILE, 4.875)];
    case 'darkHorizontal':
      return [line(0, 1.875, PATTERN_TILE, 1.875), line(0, 4.875, PATTERN_TILE, 4.875)];
    case 'lightVertical':
      return [line(4.875, 0, 4.875, PATTERN_TILE)];
    case 'darkVertical':
      return [line(1.875, 0, 1.875, PATTERN_TILE), line(4.875, 0, 4.875, PATTERN_TILE)];
    case 'lightDown':
      return [line(0.375, 5.625, 5.625, 0.375)];
    case 'darkDown':
      return [line(-1.875, 5.625, 5.625, -1.875), line(1.875, 8.625, 8.625, 1.875)];
    case 'lightUp':
      return [line(0.375, 0.375, 5.625, 5.625)];
    case 'darkUp':
      return [line(-1.875, 0.375, 5.625, 7.875), line(1.875, -2.625, 8.625, 4.875)];
    case 'lightGrid':
      return [line(0, 2.625, PATTERN_TILE, 2.625), line(2.625, 0, 2.625, PATTERN_TILE)];
    case 'darkGrid':
      return [
        line(0, 1.125, PATTERN_TILE, 1.125),
        line(0, 4.125, PATTERN_TILE, 4.125),
        line(1.125, 0, 1.125, PATTERN_TILE),
        line(4.125, 0, 4.125, PATTERN_TILE),
      ];
    case 'lightTrellis':
      return [line(0, 3, 3, 0), line(3, PATTERN_TILE, PATTERN_TILE, 3)];
    case 'darkTrellis':
      return [
        line(0, 1.5, 1.5, 0),
        line(3, PATTERN_TILE, PATTERN_TILE, 3),
        line(0, 4.5, 1.5, PATTERN_TILE),
        line(4.5, 0, PATTERN_TILE, 1.5),
      ];
    default:
      return [box(0.5)];
  }
}

/** One drawing primitive of a conditional-formatting icon. */
export type IconSegment =
  | { readonly t: 'move'; readonly x: number; readonly y: number }
  | { readonly t: 'line'; readonly x: number; readonly y: number }
  | { readonly t: 'close' }
  | { readonly t: 'circle'; readonly x: number; readonly y: number; readonly r: number };

/** One filled or stroked glyph of an icon, in page points. */
export interface IconShape {
  readonly paint: 'fill' | 'stroke';
  readonly color: string;
  /** Stroke width in points; ignored for fills. */
  readonly lineWidth: number;
  readonly segments: readonly IconSegment[];
}

const ICON_RED = '#F8696B';
const ICON_AMBER = '#FFEB84';
const ICON_GREEN = '#63BE7B';
const ICON_DARK = '#404040';
const ICON_OFF = '#E3E3E3';
const ICON_WHITE = '#FFFFFF';

/** The icon families' glyphs, matching `paintCfIcon` (cf-paint §5.1 step 8). */
export function iconShapes(item: PrintIconItem): readonly IconShape[] {
  const size = Math.min(item.box.width, item.box.height);
  if (size <= 0) return [];
  const half = size / 2;
  const centreX = item.box.x + item.box.width / 2;
  const centreY = item.box.y + item.box.height / 2;
  const width = Math.max(1, size / 10);
  switch (item.family) {
    case 'arrow':
      return [
        arrow(centreX, centreY, half, item.color, arrowAngle(item.iconIndex, item.iconCount)),
      ];
    case 'traffic':
      return [
        {
          paint: 'fill',
          color: item.color,
          lineWidth: width,
          segments: [circle(centreX, centreY, half * 0.78)],
        },
        {
          paint: 'stroke',
          color: ICON_DARK,
          lineWidth: width,
          segments: [circle(centreX, centreY, half * 0.78)],
        },
      ];
    case 'sign':
      return sign(centreX, centreY, half, size, item.iconIndex, item.iconCount);
    case 'flag':
      return flag(item.box.x, item.box.y, half, item.color, size);
    case 'rating':
      return rating(item.box.x, item.box.y, size, item.iconIndex + 1, item.iconCount);
    case 'quarter':
      return quarter(centreX, centreY, half, size, item.iconIndex + 1, item.iconCount);
    case 'dot':
      return [
        {
          paint: 'fill',
          color: item.color,
          lineWidth: width,
          segments: [circle(centreX, centreY, half * 0.46)],
        },
      ];
  }
}

function circle(x: number, y: number, r: number): IconSegment {
  return { t: 'circle', x, y, r };
}

function polygon(points: readonly (readonly [number, number])[]): readonly IconSegment[] {
  const segments: IconSegment[] = points.map(([x, y], index) =>
    index === 0 ? { t: 'move', x, y } : { t: 'line', x, y },
  );
  segments.push({ t: 'close' });
  return segments;
}

/** Up, sideways, down for three; a progressively tilted fan beyond that. */
function arrowAngle(index: number, count: number): number {
  if (count <= 1) return 0;
  const spread = Math.PI / 2;
  return -spread / 2 + (spread * index) / (count - 1);
}

function arrow(cx: number, cy: number, half: number, color: string, angle: number): IconShape {
  const points: [number, number][] = [
    [0, -half * 0.95],
    [half * 0.62, -half * 0.12],
    [half * 0.2, -half * 0.12],
    [half * 0.2, half * 0.9],
    [-half * 0.2, half * 0.9],
    [-half * 0.2, -half * 0.12],
    [-half * 0.62, -half * 0.12],
  ];
  const rotation = angle + Math.PI / 2;
  const cos = Math.cos(rotation);
  const sin = Math.sin(rotation);
  return {
    paint: 'fill',
    color,
    lineWidth: Math.max(1, half / 5),
    segments: polygon(
      points.map(([x, y]) => [cx + x * cos - y * sin, cy + x * sin + y * cos] as const),
    ),
  };
}

function sign(
  cx: number,
  cy: number,
  half: number,
  size: number,
  index: number,
  count: number,
): readonly IconShape[] {
  const last = count - 1;
  const lineWidth = Math.max(1, size / 9);
  if (index === 0) {
    const arm = half * 0.36;
    return [
      { paint: 'fill', color: ICON_RED, lineWidth, segments: [circle(cx, cy, half * 0.82)] },
      {
        paint: 'stroke',
        color: ICON_WHITE,
        lineWidth,
        segments: [
          { t: 'move', x: cx - arm, y: cy - arm },
          { t: 'line', x: cx + arm, y: cy + arm },
          { t: 'move', x: cx + arm, y: cy - arm },
          { t: 'line', x: cx - arm, y: cy + arm },
        ],
      },
    ];
  }
  if (index === last) {
    return [
      { paint: 'fill', color: ICON_GREEN, lineWidth, segments: [circle(cx, cy, half * 0.82)] },
      {
        paint: 'stroke',
        color: ICON_WHITE,
        lineWidth,
        segments: [
          { t: 'move', x: cx - half * 0.36, y: cy },
          { t: 'line', x: cx - half * 0.08, y: cy + half * 0.32 },
          { t: 'line', x: cx + half * 0.4, y: cy - half * 0.32 },
        ],
      },
    ];
  }
  return [
    {
      paint: 'fill',
      color: ICON_AMBER,
      lineWidth,
      segments: polygon([
        [cx, cy - half * 0.86],
        [cx + half * 0.86, cy + half * 0.72],
        [cx - half * 0.86, cy + half * 0.72],
      ]),
    },
    {
      paint: 'stroke',
      color: '#333333',
      lineWidth: Math.max(1, size / 3),
      segments: [
        { t: 'move', x: cx, y: cy - half * 0.34 },
        { t: 'line', x: cx, y: cy + half * 0.12 },
      ],
    },
    {
      paint: 'fill',
      color: '#333333',
      lineWidth,
      segments: [circle(cx, cy + half * 0.42, Math.max(0.7, size / 10))],
    },
  ];
}

function flag(
  x: number,
  y: number,
  half: number,
  color: string,
  size: number,
): readonly IconShape[] {
  return [
    {
      paint: 'stroke',
      color: ICON_DARK,
      lineWidth: Math.max(1, size / 12),
      segments: [
        { t: 'move', x: x + half * 0.4, y: y + half * 0.9 },
        { t: 'line', x: x + half * 0.4, y: y + half * 0.1 },
      ],
    },
    {
      paint: 'fill',
      color,
      lineWidth: 1,
      segments: polygon([
        [x + half * 0.4, y + half * 0.1],
        [x + half * 1.72, y + half * 0.42],
        [x + half * 0.4, y + half * 0.74],
      ]),
    },
  ];
}

function rating(
  x: number,
  y: number,
  size: number,
  filled: number,
  count: number,
): readonly IconShape[] {
  const slots = Math.max(1, count);
  const gap = Math.max(1, size / 12);
  const barWidth = (size - gap * (slots - 1)) / slots;
  const shapes: IconShape[] = [];
  for (let slot = 0; slot < slots; slot += 1) {
    const height = size * (0.35 + (0.65 * (slot + 1)) / slots);
    const left = x + slot * (barWidth + gap);
    const top = y + size - height;
    shapes.push({
      paint: 'fill',
      color: slot < filled ? ICON_DARK : ICON_OFF,
      lineWidth: 1,
      segments: polygon([
        [left, top],
        [left + Math.max(1, barWidth), top],
        [left + Math.max(1, barWidth), top + height],
        [left, top + height],
      ]),
    });
  }
  return shapes;
}

function quarter(
  cx: number,
  cy: number,
  half: number,
  size: number,
  filled: number,
  count: number,
): readonly IconShape[] {
  const radius = half * 0.86;
  const shapes: IconShape[] = [
    { paint: 'fill', color: ICON_OFF, lineWidth: 1, segments: [circle(cx, cy, radius)] },
    {
      paint: 'stroke',
      color: '#808080',
      lineWidth: Math.max(1, size / 14),
      segments: [circle(cx, cy, radius)],
    },
  ];
  const fraction = Math.min(1, Math.max(0, filled / Math.max(1, count - 1)));
  if (fraction === 0) return shapes;
  const steps = Math.max(2, Math.ceil(fraction * 72));
  const points: [number, number][] = [[cx, cy]];
  for (let step = 0; step <= steps; step += 1) {
    const angle = -Math.PI / 2 + (Math.PI * 2 * fraction * step) / steps;
    points.push([cx + radius * Math.cos(angle), cy + radius * Math.sin(angle)]);
  }
  shapes.push({ paint: 'fill', color: ICON_DARK, lineWidth: 1, segments: polygon(points) });
  return shapes;
}
