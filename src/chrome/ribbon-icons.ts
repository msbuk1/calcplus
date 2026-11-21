/**
 * The calcplus ribbon icon set (licensing §3.7).
 *
 * Every glyph is original SVG artwork drawn for calcplus on the same 16 x 16
 * grid. The drawing uses only `currentColor`, so hover, pressed, checked and
 * disabled state come from the control's colour tokens rather than the icon.
 * Keep new artwork in this file: one family, one optical size, one baseline.
 */

export const RIBBON_ICON_SIZE = 16;

/**
 * The square badge frame shared by the Function Library tiles. Insert and
 * Merge drew this frame first; every Function Library glyph reuses it so the
 * ten badges line up down the group. The interior stays open: only open
 * marks, no closed detail that fills in at 16px.
 */
export const RIBBON_BADGE_FRAME = 'M2.75 2.75h10.5v10.5H2.75z';

export interface RibbonIconCircle {
  readonly cx: number;
  readonly cy: number;
  readonly radius: number;
  /** A dot or pupil: filled with `currentColor` instead of stroked. */
  readonly filled?: boolean;
}

export interface RibbonIconDefinition {
  /** Open stroked outlines. */
  readonly paths: readonly string[];
  /** Solid shapes, drawn with `currentColor` and no outline. */
  readonly filled?: readonly string[];
  /** Interior rules drawn with a dash pattern. */
  readonly dashed?: readonly string[];
  /** Structural guides drawn with the shared stroke at reduced emphasis. */
  readonly muted?: readonly string[];
  readonly circles?: readonly RibbonIconCircle[];
}

const ICON_DRAWINGS = {
  'align-bottom': {
    paths: ['M2.5 12.5h11', 'M8 3v6.5', 'M5.5 7.5 8 10l2.5-2.5'],
  },
  'align-centre': {
    paths: ['M2.5 3.5h11', 'M4.5 6.5h7', 'M2.5 9.5h11', 'M4.5 12.5h7'],
  },
  'align-left': {
    paths: ['M2.5 3.5h11', 'M2.5 6.5h7', 'M2.5 9.5h11', 'M2.5 12.5h7'],
  },
  'align-middle': {
    paths: ['M2.5 8h11', 'M5.5 5.5 8 3l2.5 2.5', 'M5.5 10.5 8 13l2.5-2.5'],
  },
  'align-right': {
    paths: ['M2.5 3.5h11', 'M6.5 6.5h7', 'M2.5 9.5h11', 'M6.5 12.5h7'],
  },
  'align-top': {
    paths: ['M2.5 3.5h11', 'M8 13V6.5', 'M5.5 8.5 8 6l2.5 2.5'],
  },
  autosum: {
    // The framed sigma is shared by the Home › Editing split and the
    // Formulas › Function Library tile, so both tabs read the same badge.
    paths: [RIBBON_BADGE_FRAME, 'M11.1 4.6H5.4l3.1 3.6-3.1 3.6h5.7'],
  },
  bold: {
    paths: [
      'M5.25 2.75h3.75a2.75 2.75 0 0 1 0 5.5H5.25z',
      'M5.25 8.25h4.25a2.875 2.875 0 0 1 0 5.75H5.25z',
    ],
  },
  'border-all': {
    paths: [
      'M2.75 2.75h10.5v10.5H2.75z',
      'M6.25 2.75v10.5',
      'M9.75 2.75v10.5',
      'M2.75 6.25h10.5',
      'M2.75 9.75h10.5',
    ],
  },
  'border-bottom': {
    muted: ['M2.75 2.75h10.5v10.5H2.75z'],
    paths: ['M2.75 13.25h10.5'],
  },
  'border-bottom-double': {
    muted: ['M2.75 2.75h10.5v10.5H2.75z'],
    paths: ['M2.75 12.2h10.5', 'M2.75 13.45h10.5'],
  },
  'border-inside': {
    muted: ['M2.75 2.75h10.5v10.5H2.75z'],
    paths: ['M6.25 2.75v10.5', 'M9.75 2.75v10.5', 'M2.75 6.25h10.5', 'M2.75 9.75h10.5'],
  },
  'border-left': {
    muted: ['M2.75 2.75h10.5v10.5H2.75z'],
    paths: ['M2.75 2.75v10.5'],
  },
  'border-more': {
    // The More Borders entry opens Format Cells; it shares that launcher glyph.
    paths: ['M2.75 2.75h10.5v10.5H2.75z', 'M10.9 4.1 12.4 5.6 6.4 11.6H4.9v-1.5z'],
  },
  'border-none': {
    muted: ['M2.75 2.75h10.5v10.5H2.75z'],
    paths: ['M3.1 12.9 12.9 3.1'],
  },
  'border-outside': {
    paths: ['M2.75 2.75h10.5v10.5H2.75z'],
  },
  'border-right': {
    muted: ['M2.75 2.75h10.5v10.5H2.75z'],
    paths: ['M13.25 2.75v10.5'],
  },
  'border-thick-bottom': {
    paths: [],
    muted: ['M2.75 2.75h10.5v10.5H2.75z'],
    filled: ['M2.75 12h10.5v1.25H2.75z'],
  },
  'border-thick-outside': {
    paths: ['M2.25 2.25h11.5v11.5H2.25z', 'M3.5 3.5h9v9h-9z'],
  },
  'border-top': {
    muted: ['M2.75 2.75h10.5v10.5H2.75z'],
    paths: ['M2.75 2.75h10.5'],
  },
  'border-top-bottom': {
    muted: ['M2.75 2.75h10.5v10.5H2.75z'],
    paths: ['M2.75 2.75h10.5', 'M2.75 13.25h10.5'],
  },
  'border-top-double-bottom': {
    muted: ['M2.75 2.75h10.5v10.5H2.75z'],
    paths: ['M2.75 2.75h10.5', 'M2.75 12.2h10.5', 'M2.75 13.45h10.5'],
  },
  'border-top-thick-bottom': {
    paths: ['M2.75 2.75h10.5'],
    muted: ['M2.75 2.75h10.5v10.5H2.75z'],
    filled: ['M2.75 12h10.5v1.25H2.75z'],
  },
  borders: {
    paths: [
      'M2.75 2.75h10.5v10.5H2.75z',
      'M6.25 2.75v10.5',
      'M9.75 2.75v10.5',
      'M2.75 6.25h10.5',
      'M2.75 9.75h10.5',
    ],
  },
  clear: {
    paths: [
      'M3.1 11.9 8.7 6.3a1.4 1.4 0 0 1 2 0l2 2a1.4 1.4 0 0 1 0 2l-3.6 3.6H5.5z',
      'M5.7 9.3l4.9 4.9',
    ],
  },
  comma: {
    paths: [],
    filled: ['M9.4 6.4c0 2.2-1.2 3.8-3 4.8l-.6-.95c1.2-.7 1.9-1.6 2-2.7H6.2V4.3h3.2z'],
  },
  'conditional-formatting': {
    paths: ['M2.75 7h7.5v2.5h-7.5z', 'M2.75 10.75h4.5v2.5H2.75z'],
    filled: ['M2.75 3.25h10.5v2.5H2.75z'],
  },
  copy: {
    paths: [
      'M5.5 2.75h5.75a1.5 1.5 0 0 1 1.5 1.5v8',
      'M2.75 4.75h6.5a1.5 1.5 0 0 1 1.5 1.5v6.5a1.5 1.5 0 0 1-1.5 1.5h-6.5a1.5 1.5 0 0 1-1.5-1.5v-6.5a1.5 1.5 0 0 1 1.5-1.5z',
    ],
  },
  currency: {
    paths: ['M8 2.5v11', 'M11 4.5a3 3 0 0 0-5.3 1c0 2.3 4.6 1 4.6 3.6a3 3 0 0 1-5.6 1.2'],
  },
  cut: {
    paths: ['M5.3 5.4 13 12', 'M5.3 10.6 13 4'],
    circles: [
      { cx: 4, cy: 4, radius: 1.6 },
      { cx: 4, cy: 12, radius: 1.6 },
    ],
  },
  'decrease-decimal': {
    paths: [
      'M5.4 8H1.2',
      'M2.8 5.6 0.6 8 2.8 10.4',
      'M7.2 8a1.8 3.2 0 1 0 3.6 0a1.8 3.2 0 1 0-3.6 0',
      'M11.7 8a1.8 3.2 0 1 0 3.6 0a1.8 3.2 0 1 0-3.6 0',
    ],
  },
  delete: {
    paths: ['M2.75 2.75h10.5v10.5H2.75z', 'M5.75 5.75l4.5 4.5', 'M10.25 5.75l-4.5 4.5'],
  },
  fill: {
    paths: [
      'M3.1 8.9 7.9 4.1a1 1 0 0 1 1.4 0l2.6 2.6a1 1 0 0 1 0 1.4l-4.8 4.8a1 1 0 0 1-1.4 0L3.1 10.3a1 1 0 0 1 0-1.4z',
      'M5.2 4.8 3.9 3.5a1.3 1.3 0 0 1 1.85-1.85l1.85 1.85',
    ],
    filled: ['M12.6 10.8c.65.85 1 1.5 1 2a1.05 1.05 0 0 1-2.1 0c0-.5.35-1.15 1.1-2z'],
  },
  filter: {
    paths: ['M2.75 3.25h10.5l-4 4.75v4.25l-2.5 1.5V8z'],
  },
  'filter-clear': {
    paths: ['M2.25 3h8.5l-3.2 4v3.6l-2.1 1.25V7z', 'M10.75 9.5l3.5 3.5', 'M14.25 9.5l-3.5 3.5'],
  },
  'filter-reapply': {
    paths: ['M2.25 3h8.5l-3.2 4v3.6l-2.1 1.25V7z', 'M14 10.2a3 3 0 1 1-.9-2.7', 'M14 8v2.4h-2.4'],
  },
  find: {
    paths: ['M10.3 10.3 13.6 13.6'],
    circles: [{ cx: 7, cy: 7, radius: 4.2 }],
  },
  'font-colour': {
    paths: ['M3.25 12.5 8 3l4.75 9.5', 'M5.15 9.25h5.7'],
    filled: ['M2.75 13.5h10.5v1.35H2.75z'],
  },
  'font-family': {
    paths: [
      'M2.75 12 5.5 4l2.75 8',
      'M3.9 9.25h3.2',
      'M10.5 8.5c0-1.4 2.75-1.4 2.75 0v3.5',
      'M13.25 9.6c-2.4-1-3.55 1.6-2.2 2.6.9.65 2 .3 2.2-.6',
    ],
  },
  'font-size': {
    paths: [
      'M2.75 12 5.5 4l2.75 8',
      'M3.9 9.25h3.2',
      'M11.5 3v10',
      'M9.75 4.75 11.5 3l1.75 1.75',
      'M9.75 11.25 11.5 13l1.75-1.75',
    ],
  },
  'format-cells': {
    paths: ['M2.75 2.75h10.5v10.5H2.75z', 'M10.9 4.1 12.4 5.6 6.4 11.6H4.9v-1.5z'],
  },
  'format-painter': {
    paths: [
      'M5.25 10.1 11.6 3.75a1.75 1.75 0 0 1 2.5 2.5L7.75 12.6z',
      'M4.35 9.95l1.7 1.7',
      'M2.75 14.5c1.15.15 2-.25 2.55-1.15',
    ],
  },
  'freeze-panes': {
    paths: ['M2.75 2.75h10.5v10.5H2.75z', 'M6.5 2.75v10.5', 'M2.75 6.5h10.5'],
    filled: ['M2.75 2.75h3.75v3.75H2.75z'],
  },
  // The Function Library category menus. Every glyph is the same square badge
  // frame with only open interior marks; nothing fills in at 16px.
  'function-recent': {
    paths: [RIBBON_BADGE_FRAME, 'M8 5.6V8h2.1'],
  },
  'function-financial': {
    paths: [RIBBON_BADGE_FRAME],
    circles: [{ cx: 8, cy: 8, radius: 2.2 }],
  },
  'function-logical': {
    paths: [RIBBON_BADGE_FRAME, 'M8 5.9 10.1 8 8 10.1 5.9 8z'],
  },
  'function-text': {
    paths: [RIBBON_BADGE_FRAME, 'M5.4 6.2h5.2', 'M8 6.2v4.2'],
  },
  'function-date-time': {
    // A short rule and one marked day. The rule stops inside the frame on
    // purpose: spanning the full width reads as a solid slab at 16px.
    paths: [RIBBON_BADGE_FRAME, 'M5.2 6.2h5.6'],
    circles: [{ cx: 8, cy: 9.8, radius: 1, filled: true }],
  },
  'function-lookup': {
    // The shaft stops short of the head on purpose. Ending it at the tip
    // piles three strokes onto one point, which rasterises as a blob at 16px.
    // Two-stroke corners stay open, as the radical and the diamond show.
    paths: [RIBBON_BADGE_FRAME, 'M4.9 8h3.3', 'M8.2 5.8 10.4 8 8.2 10.2'],
  },
  'function-math': {
    paths: [RIBBON_BADGE_FRAME, 'M4.8 8.5 6 11.2 7.6 5.6h3.4'],
  },
  'function-more': {
    paths: [RIBBON_BADGE_FRAME],
    circles: [
      { cx: 5.2, cy: 8, radius: 0.8, filled: true },
      { cx: 8, cy: 8, radius: 0.8, filled: true },
      { cx: 10.8, cy: 8, radius: 0.8, filled: true },
    ],
  },
  'go-to': {
    paths: ['M8 2.75v1.5', 'M8 11.75v1.5', 'M2.75 8h1.5', 'M11.75 8h1.5'],
    circles: [
      { cx: 8, cy: 8, radius: 3.4 },
      { cx: 8, cy: 8, radius: 0.7, filled: true },
    ],
  },
  gridlines: {
    paths: ['M2.75 2.75h10.5v10.5H2.75z'],
    dashed: ['M2.75 6.25h10.5', 'M2.75 9.75h10.5', 'M6.25 2.75v10.5', 'M9.75 2.75v10.5'],
  },
  headings: {
    paths: ['M2.75 2.75h10.5v10.5H2.75z', 'M6.25 2.75v10.5', 'M2.75 6.25h10.5', 'M2.75 9.75h10.5'],
    filled: ['M2.75 2.75h10.5v3.5H2.75z'],
  },
  insert: {
    paths: ['M2.75 2.75h10.5v10.5H2.75z', 'M8 5.5v5', 'M5.5 8h5'],
  },
  'insert-function': {
    paths: [
      RIBBON_BADGE_FRAME,
      // An italic f and a subscript x, two open marks inside the badge.
      'M8.2 4.6H7.4c-.8 0-1.2.5-1.3 1.3L5.2 10.8M4.9 7.4h2.6',
      'M9.2 7.5 11.9 10.8M11.9 7.5 9.2 10.8',
    ],
  },
  'increase-decimal': {
    paths: [
      'M1.2 8h4.2',
      'M3.2 5.6 5.4 8 3.2 10.4',
      'M7.2 8a1.8 3.2 0 1 0 3.6 0a1.8 3.2 0 1 0-3.6 0',
      'M11.7 8a1.8 3.2 0 1 0 3.6 0a1.8 3.2 0 1 0-3.6 0',
    ],
  },
  italic: {
    paths: ['M9.5 2.75h4', 'M2.5 13.25h4', 'M9.75 2.75 6.25 13.25'],
  },
  merge: {
    paths: [
      'M2.75 2.75h10.5v10.5H2.75z',
      'M2.75 8h3.4',
      'M4.75 6.35 6.4 8 4.75 9.65',
      'M13.25 8h-3.4',
      'M11.25 6.35 9.6 8 11.25 9.65',
    ],
  },
  'number-format': {
    paths: ['M5.5 3.25 4.5 12.75', 'M10.5 3.25 9.5 12.75', 'M3 6.25h10', 'M2.5 9.75h10'],
  },
  'page-setup': {
    paths: [
      'M3.5 2.5h6.5l2.5 2.5v8.5h-9z',
      'M10 2.5v2.5h2.5',
      'M5.5 6.75h5',
      'M5.5 9.25h5',
      'M5.5 11.75h3',
    ],
  },
  'print-area': {
    paths: ['M3 2.5h7l3 3v8H3z', 'M10 2.5v3h3'],
    dashed: ['M5 6.75h5.5v4.5H5z'],
  },
  paste: {
    paths: [
      'M5.75 2.5H4.5A1.5 1.5 0 0 0 3 4v9a1.5 1.5 0 0 0 1.5 1.5h7A1.5 1.5 0 0 0 13 13V4a1.5 1.5 0 0 0-1.5-1.5h-1.25',
      'M6 1.75h4a1 1 0 0 1 1 1v1.5a1 1 0 0 1-1 1H6a1 1 0 0 1-1-1v-1.5a1 1 0 0 1 1-1z',
      'M6 7.25h4',
      'M6 9.75h4',
      'M6 12.25h2.5',
    ],
  },
  percent: {
    paths: ['M12.5 3.5 3.5 12.5'],
    circles: [
      { cx: 4.5, cy: 4.5, radius: 1.6 },
      { cx: 11.5, cy: 11.5, radius: 1.6 },
    ],
  },
  print: {
    paths: ['M2.25 6h11.5v5.25H2.25z', 'M4.75 2.5h6.5v3.5h-6.5z', 'M4.75 9h6.5v4.5h-6.5z'],
    circles: [{ cx: 12, cy: 7.4, radius: 0.55, filled: true }],
  },
  redo: {
    paths: ['M10.25 4.25 13 7l-2.75 2.75', 'M13 7H6.9A3.9 3.9 0 0 0 3 10.9v.6'],
  },
  replace: {
    paths: [
      'M2.5 5.25h9.25',
      'M9.5 3 12 5.25 9.5 7.5',
      'M13.5 10.75H4.25',
      'M6.5 8.5 4 10.75 6.5 13',
    ],
  },
  'show-comments': {
    paths: [
      'M3 3.25h10a1.5 1.5 0 0 1 1.5 1.5v5a1.5 1.5 0 0 1-1.5 1.5H8.5L5.25 14v-2.75H3a1.5 1.5 0 0 1-1.5-1.5v-5a1.5 1.5 0 0 1 1.5-1.5z',
      'M4.5 6h7',
      'M4.5 8h4.5',
    ],
  },
  save: {
    paths: [
      'M2.75 3.25h8.25l2.25 2.25v7.25a.75.75 0 0 1-.75.75H3.5a.75.75 0 0 1-.75-.75z',
      'M5.25 2.75v3.5h4.5v-3.5',
      'M5 9.25h6v4.25H5z',
    ],
  },
  sort: {
    paths: [
      'M2.5 7.5 4.5 2.5l2 5',
      'M3.15 6h2.7',
      'M2.5 9.5h4l-4 4h4',
      'M11.75 3.25v9.5',
      'M9.75 10.5 11.75 12.5l2-2',
    ],
  },
  strikethrough: {
    paths: [
      'M11.4 4.4a3.2 3.2 0 0 0-5.6 1.2c0 2.4 5.3 1.1 5.3 3.9a3.4 3.4 0 0 1-6.2 1.4',
      'M2.5 8h11',
    ],
  },
  underline: {
    paths: ['M4.5 2.5v5a3.5 3.5 0 0 0 7 0v-5', 'M3.5 14h9'],
  },
  undo: {
    paths: ['M5.75 4.25 3 7l2.75 2.75', 'M3 7h6.1A3.9 3.9 0 0 1 13 10.9v.6'],
  },
  'wrap-text': {
    paths: ['M2.5 3.5h11', 'M2.5 7.5h7.5a2.5 2.5 0 0 1 0 5H7', 'M8.5 10.5 7 12l1.5 1.5'],
  },
  zoom: {
    paths: ['M10 10l3.5 3.5', 'M7 5.2v3.6', 'M5.2 7h3.6'],
    circles: [{ cx: 7, cy: 7, radius: 4 }],
  },
  'zoom-100': {
    paths: [
      'M9.5 9.5l3.75 3.75',
      'M4.75 6v3',
      'M6.8 6.2a1.1 1.4 0 1 0 0 2.8 1.1 1.4 0 0 0 0-2.8z',
      'M9.4 6.2a1.1 1.4 0 1 0 0 2.8 1.1 1.4 0 0 0 0-2.8z',
    ],
    circles: [{ cx: 6.75, cy: 6.75, radius: 4.25 }],
  },
} as const satisfies Readonly<Record<string, RibbonIconDefinition>>;

export type RibbonIconName = keyof typeof ICON_DRAWINGS;

/** The single source of truth for every icon name used by the chrome map. */
export const RIBBON_ICONS: Readonly<Record<RibbonIconName, RibbonIconDefinition>> = ICON_DRAWINGS;

export const RIBBON_ICON_NAMES: readonly RibbonIconName[] = Object.freeze(
  (Object.keys(ICON_DRAWINGS) as RibbonIconName[]).sort(),
);

const RIBBON_ICON_SET: ReadonlySet<string> = new Set(RIBBON_ICON_NAMES);

export function isRibbonIconName(value: string): value is RibbonIconName {
  return RIBBON_ICON_SET.has(value);
}
