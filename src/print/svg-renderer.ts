import type {
  PrintBox,
  PrintDataBarItem,
  PrintFillItem,
  PrintIconFamily,
  PrintIconItem,
  PrintItem,
  PrintLineItem,
  PrintPage,
  PrintPaper,
  PrintPlan,
  PrintTextItem,
} from './types';

/**
 * The SVG preview renderer (print-pdf §2): one SVG element per page, in page
 * points, for the preview window and `window.print()`.
 *
 * It renders exactly what the `PrintPlan` carries and nothing else. There is no
 * layout here: every coordinate, font size and stroke width is already in page
 * points, and the item array is painted in its own order (the planner's §5.1
 * order). The output is a plain string, so the preview, the print stack and the
 * tests all read the same bytes.
 *
 * Two measured facts from the M7 spike shape the print path
 * (`.scratch/calcplus/research/26-print-pdf-fidelity.md` §1):
 *
 * - SVG fill and stroke print even with "Background graphics" off, while a CSS
 *   background is suppressed. Every fill, pattern and border here is SVG
 *   shape paint, never a CSS background.
 * - Chromium's real A4 page box is 594.96 × 841.92pt, not the nominal
 *   595.28 × 841.89. Content sized to nominal A4 overflows by 0.32pt and
 *   spawns an empty extra page. `printPageBox` returns the box a print page
 *   element may occupy: always inside the browser's page, so N planned pages
 *   print as exactly N pages. The SVG itself keeps the plan's exact geometry
 *   and is clipped by that element.
 *
 * Security: every string that comes from the sheet (cell text, font family,
 * colours) is escaped or validated before it reaches the markup, so a hostile
 * workbook cannot inject elements or attributes.
 */

/** Points per CSS pixel at 96 dpi. */
const PT_PER_PIXEL = 0.75;

/** Chromium's measured A4 page box in points (spike §2.3). */
export const CHROMIUM_A4_PAGE = Object.freeze({ width: 594.96, height: 841.92 });

/** Nominal A4 in points, the plan's own paper box. */
const NOMINAL_A4 = { width: 595.28, height: 841.89 };

/**
 * The gap between a print page element and the browser's page box. A whole
 * point keeps sub-pixel rounding from ever spilling onto an empty page; it
 * falls inside the page margin, so no sheet content is lost.
 */
export const PRINT_PAGE_SAFETY = 1;

/** The OOXML pattern tile is 8 CSS pixels, i.e. 6 points. */
const PATTERN_TILE = 8 * PT_PER_PIXEL;

const ICON_RED = '#F8696B';
const ICON_AMBER = '#FFEB84';
const ICON_GREEN = '#63BE7B';
const ICON_DARK = '#404040';
const ICON_OFF = '#E3E3E3';
const ICON_OUTLINE = '#808080';

/** A print page element's size: never larger than the browser's page box. */
export interface PrintPageBox {
  readonly width: number;
  readonly height: number;
}

/** One rendered page, ready for the preview or the print stack. */
export interface RenderedPrintPage {
  /** 0-based position in the plan. */
  readonly index: number;
  /** 1-based page number. */
  readonly number: number;
  /** The announced label: "Page 1 of 3, Sheet1". */
  readonly label: string;
  /** The page box in points, exactly as the plan carries it. */
  readonly box: PrintBox;
  /** The box the print page element occupies (inside the browser's page). */
  readonly printBox: PrintPageBox;
  /** The `<svg>` markup for the page. */
  readonly svg: string;
}

/** Everything the preview window and the print-only stack need. */
export interface RenderedPrintDocument {
  readonly sheetName: string;
  readonly paper: PrintPaper;
  readonly pageCount: number;
  readonly pages: readonly RenderedPrintPage[];
}

/**
 * The box a printed page element may occupy. For A4 it is Chromium's measured
 * real page box, not the nominal one; every paper keeps `PRINT_PAGE_SAFETY` of
 * slack so a page can never overflow into a blank one.
 */
export function printPageBox(paper: PrintPaper): PrintPageBox {
  const real = chromiumPageBox(paper) ?? { width: paper.width, height: paper.height };
  return {
    width: Math.max(1, round(real.width - PRINT_PAGE_SAFETY)),
    height: Math.max(1, round(real.height - PRINT_PAGE_SAFETY)),
  };
}

/** The announced page label. Exported so the dialog and the SVG never differ. */
export function printPageLabel(number: number, pageCount: number, sheetName: string): string {
  const sheet = sheetName.trim() === '' ? 'Sheet' : sheetName.trim();
  return `Page ${number} of ${pageCount}, ${sheet}`;
}

/** Render one plan into one SVG per page. */
export function renderPrintPlan(plan: PrintPlan): RenderedPrintDocument {
  const pageCount = plan.pages.length;
  const printBox = printPageBox(plan.paper);
  const pages: RenderedPrintPage[] = plan.pages.map((page) => {
    const label = printPageLabel(page.index + 1, pageCount, plan.sheetName);
    return {
      index: page.index,
      number: page.index + 1,
      label,
      box: page.box,
      printBox,
      svg: renderPrintPage(page, label),
    };
  });
  return { sheetName: plan.sheetName, paper: plan.paper, pageCount, pages };
}

/** Render one planned page as a standalone `<svg>` element. */
export function renderPrintPage(page: PrintPage, label: string): string {
  const prefix = `print-${page.index}`;
  const definitions = new Definitions(prefix);
  const clips = new ClipRegistry(prefix);
  const body: string[] = [];
  for (const item of page.items) body.push(renderItem(item, definitions, clips));

  const { width, height } = page.box;
  const root =
    `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${num(width)} ${num(height)}"` +
    ` width="${num(width)}pt" height="${num(height)}pt" preserveAspectRatio="xMinYMin meet"` +
    ` role="img" aria-label="${escapeAttribute(label)}" data-print-page="${page.index + 1}">`;
  const defs = definitions.render(clips);
  return `${root}${defs}${body.join('')}</svg>`;
}

// ---------------------------------------------------------------------------
// Items
// ---------------------------------------------------------------------------

function renderItem(item: PrintItem, defs: Definitions, clips: ClipRegistry): string {
  switch (item.kind) {
    case 'fill':
      return renderFill(item, defs);
    case 'line':
      return renderLine(item);
    case 'text':
      return renderText(item, clips);
    case 'dataBar':
      return renderDataBar(item, defs);
    case 'icon':
      return renderIcon(item);
    case 'image':
      // v1 reserves the anchor and paints nothing (print-pdf §5).
      return '';
  }
}

function renderFill(item: PrintFillItem, defs: Definitions): string {
  if (item.pattern === '' || item.pattern === 'solid') {
    const fill = safeColor(item.color, null);
    if (fill === null) return '';
    return `<rect ${rectAttributes(item.box)} fill="${fill}" data-kind="fill"/>`;
  }
  if (item.color === null && item.background === null) return '';
  const id = defs.pattern(
    item.pattern,
    safeColor(item.color, null),
    safeColor(item.background, null),
  );
  return `<rect ${rectAttributes(item.box)} fill="url(#${id})" data-kind="fill"/>`;
}

function renderLine(item: PrintLineItem): string {
  const stroke = safeColor(item.color, '#000000');
  const width = Number.isFinite(item.width) && item.width > 0 ? item.width : 0.25;
  const attributes = [
    `x1="${num(item.x1)}"`,
    `y1="${num(item.y1)}"`,
    `x2="${num(item.x2)}"`,
    `y2="${num(item.y2)}"`,
    `stroke="${stroke}"`,
    `stroke-width="${num(width)}"`,
    'stroke-linecap="butt"',
  ];
  if (item.dash !== null && item.dash.length > 0) {
    attributes.push(`stroke-dasharray="${item.dash.map(num).join(' ')}"`);
  }
  return `<line ${attributes.join(' ')} data-kind="line"/>`;
}

function renderText(item: PrintTextItem, clips: ClipRegistry): string {
  const text = item.text;
  if (text === '') return '';
  const font = item.font;
  const size = Number.isFinite(font.size) && font.size > 0 ? font.size : 1;
  const attributes = [
    `x="${num(item.x)}"`,
    `y="${num(item.y)}"`,
    `text-anchor="${anchorOf(item.anchor)}"`,
    `font-family="${escapeAttribute(fontFamily(font.family))}"`,
    `font-size="${num(size)}"`,
    `fill="${safeColor(font.color, '#000000')}"`,
  ];
  if (font.bold) attributes.push('font-weight="bold"');
  if (font.italic) attributes.push('font-style="italic"');
  if (Number.isFinite(item.rotation) && item.rotation !== 0) {
    attributes.push(`transform="rotate(${num(item.rotation)} ${num(item.x)} ${num(item.y)})"`);
  }
  const clipId = clips.ensure(item.clip);
  return (
    `<g clip-path="url(#${clipId})"><text ${attributes.join(' ')} xml:space="preserve"` +
    ` data-kind="text">${escapeText(text)}</text></g>`
  );
}

/**
 * §5.1 step 3: the data bar is a rect (solid or a left-to-right gradient), an
 * optional 1pt border and an optional 1pt zero axis. Text paints after it.
 */
function renderDataBar(item: PrintDataBarItem, defs: Definitions): string {
  const fill =
    item.gradient === null
      ? safeColor(item.color, null)
      : `url(#${defs.gradient(
          safeColor(item.gradient.from, '#000000') ?? '#000000',
          safeColor(item.gradient.to, '#000000') ?? '#000000',
        )})`;
  if (fill === null) return '';
  const parts: string[] = [`<rect ${rectAttributes(item.box)} fill="${fill}"/>`];
  if (item.borderColor !== null) {
    parts.push(
      `<rect ${rectAttributes(item.box)} fill="none" stroke="${safeColor(item.borderColor, '#000000')}" stroke-width="1"/>`,
    );
  }
  if (item.axisX !== null) {
    parts.push(
      `<line x1="${num(item.axisX)}" y1="${num(item.box.y)}" x2="${num(item.axisX)}"` +
        ` y2="${num(item.box.y + item.box.height)}" stroke="${safeColor(item.axisColor, '#808080')}"` +
        ' stroke-width="1"/>',
    );
  }
  return `<g data-kind="dataBar">${parts.join('')}</g>`;
}

/**
 * §5.1 step 8: one glyph per icon family. The geometry is the grid's own
 * `paintCfIcon` (`src/grid/cf-paint.ts`), which is the authoritative, shipped
 * artwork; this module mirrors its maths rather than its pixels. The index is
 * clamped to the set exactly as the grid clamps it before drawing.
 */
function renderIcon(item: PrintIconItem): string {
  const box = item.box;
  if (box.width <= 0 || box.height <= 0) return '';
  const size = Math.min(box.width, box.height);
  const count = Math.max(1, Math.round(item.iconCount));
  const index = clamp(Math.round(item.iconIndex), 0, count - 1);
  const glyph = iconGlyph(item.family, item, box, size, index, count);
  if (glyph === '') return '';
  return (
    `<g data-kind="icon" data-icon-family="${item.family}" data-icon-index="${item.iconIndex}"` +
    ` data-icon-count="${item.iconCount}">${glyph}</g>`
  );
}

function iconGlyph(
  family: PrintIconFamily,
  item: PrintIconItem,
  box: PrintBox,
  size: number,
  index: number,
  count: number,
): string {
  switch (family) {
    case 'arrow':
      return arrowGlyph(box, size, safeColor(item.color, ICON_DARK), index, count);
    case 'traffic':
      return trafficGlyph(box, size, safeColor(item.color, ICON_DARK));
    case 'sign':
      return signGlyph(box, size, index, count);
    case 'flag':
      return flagGlyph(box, size, safeColor(item.color, ICON_DARK));
    case 'rating':
      return ratingGlyph(box, size, index, count);
    case 'quarter':
      return quarterGlyph(box, size, index, count);
    case 'dot':
      return dotGlyph(box, size, safeColor(item.color, ICON_DARK));
  }
}

/**
 * `paintArrow` + `arrowAngle`: the path points up and the canvas rotates it by
 * `arrowAngle + 90°`, where `arrowAngle` is
 * `-45° + 90° * index / (count - 1)` (0° for a one-icon set). The clockwise
 * rotation from the up-facing path is therefore `45° + 90° * index / (count - 1)`
 * and a flat `90°` for one icon: up-right, sideways, down-right for three.
 */
function arrowGlyph(
  box: PrintBox,
  size: number,
  color: string,
  index: number,
  count: number,
): string {
  const cx = box.x + box.width / 2;
  const cy = box.y + box.height / 2;
  const half = size / 2;
  const points = [
    [cx, cy - half * 0.95],
    [cx + half * 0.62, cy - half * 0.12],
    [cx + half * 0.2, cy - half * 0.12],
    [cx + half * 0.2, cy + half * 0.9],
    [cx - half * 0.2, cy + half * 0.9],
    [cx - half * 0.2, cy - half * 0.12],
    [cx - half * 0.62, cy - half * 0.12],
  ];
  const path = `M ${points.map(([x, y]) => `${num(x!)} ${num(y!)}`).join(' L ')} Z`;
  const arrowAngle = count <= 1 ? 0 : -45 + (90 * index) / (count - 1);
  const rotation = arrowAngle + 90;
  return `<path d="${path}" fill="${color}" transform="rotate(${num(rotation)} ${num(cx)} ${num(cy)})"/>`;
}

function trafficGlyph(box: PrintBox, size: number, color: string): string {
  const cx = box.x + box.width / 2;
  const cy = box.y + box.height / 2;
  const stroke = Math.max(0.5, size / 10);
  return (
    `<circle cx="${num(cx)}" cy="${num(cy)}" r="${num(size * 0.39)}" fill="${color}"` +
    ` stroke="${ICON_DARK}" stroke-width="${num(stroke)}"/>`
  );
}

/**
 * Index 0 is a red disc with a cross, the last index a green disc with a tick,
 * every middle index an amber triangle with a bar and a dot.
 */
function signGlyph(box: PrintBox, size: number, index: number, count: number): string {
  const cx = box.x + box.width / 2;
  const cy = box.y + box.height / 2;
  const half = size / 2;
  const last = Math.max(0, count - 1);
  const stroke = Math.max(0.5, size / 9);
  if (index <= 0) {
    const arm = half * 0.36;
    return (
      `<circle cx="${num(cx)}" cy="${num(cy)}" r="${num(half * 0.82)}" fill="${ICON_RED}"/>` +
      `<path d="M ${num(cx - arm)} ${num(cy - arm)} L ${num(cx + arm)} ${num(cy + arm)}` +
      ` M ${num(cx + arm)} ${num(cy - arm)} L ${num(cx - arm)} ${num(cy + arm)}"` +
      ` stroke="#FFFFFF" stroke-width="${num(stroke)}" fill="none"/>`
    );
  }
  if (index >= last) {
    return (
      `<circle cx="${num(cx)}" cy="${num(cy)}" r="${num(half * 0.82)}" fill="${ICON_GREEN}"/>` +
      `<path d="M ${num(cx - half * 0.36)} ${num(cy)} L ${num(cx - half * 0.08)} ${num(cy + half * 0.32)}` +
      ` L ${num(cx + half * 0.4)} ${num(cy - half * 0.32)}"` +
      ` stroke="#FFFFFF" stroke-width="${num(stroke)}" fill="none"/>`
    );
  }
  const triangle =
    `M ${num(cx)} ${num(cy - half * 0.86)} L ${num(cx + half * 0.86)} ${num(cy + half * 0.72)}` +
    ` L ${num(cx - half * 0.86)} ${num(cy + half * 0.72)} Z`;
  return (
    `<path d="${triangle}" fill="${ICON_AMBER}"/>` +
    `<path d="M ${num(cx)} ${num(cy - half * 0.34)} L ${num(cx)} ${num(cy + half * 0.12)}"` +
    ` stroke="#333333" stroke-width="${num(Math.max(0.5, size / 3))}" fill="none"/>` +
    `<circle cx="${num(cx)}" cy="${num(cy + half * 0.42)}" r="${num(Math.max(0.35, size / 10))}"` +
    ' fill="#333333"/>'
  );
}

function flagGlyph(box: PrintBox, size: number, color: string): string {
  const half = size / 2;
  const poleX = box.x + half * 0.4;
  const top = box.y + half * 0.1;
  const stroke = Math.max(0.5, size / 12);
  return (
    `<path d="M ${num(poleX)} ${num(box.y + half * 0.9)} L ${num(poleX)} ${num(top)}"` +
    ` stroke="${ICON_DARK}" stroke-width="${num(stroke)}" fill="none"/>` +
    `<path d="M ${num(poleX)} ${num(top)} L ${num(box.x + half * 1.72)} ${num(box.y + half * 0.42)}` +
    ` L ${num(poleX)} ${num(box.y + half * 0.74)} Z" fill="${color}"/>`
  );
}

/** `iconCount` bars; the first `iconIndex + 1` are dark, the rest pale. */
function ratingGlyph(box: PrintBox, size: number, index: number, count: number): string {
  const slots = Math.max(1, count);
  const gap = Math.max(0.5, size / 12);
  const barWidth = Math.max(0.5, (size - gap * (slots - 1)) / slots);
  const filled = clamp(index + 1, 0, slots);
  const parts: string[] = [];
  for (let slot = 0; slot < slots; slot += 1) {
    const height = size * (0.35 + (0.65 * (slot + 1)) / slots);
    parts.push(
      `<rect x="${num(box.x + slot * (barWidth + gap))}"` +
        ` y="${num(box.y + size - height)}" width="${num(barWidth)}" height="${num(height)}"` +
        ` fill="${slot < filled ? ICON_DARK : ICON_OFF}"/>`,
    );
  }
  return parts.join('');
}

/**
 * `paintQuarter`: a grey disc (0.86·half) with a dark pie wedge starting at
 * 12 o'clock. The canvas is given `index + 1` as its `filled` count and draws
 * `clamp((index + 1) / max(1, count - 1))` of a turn; a fraction of 1 is a full
 * dark disc, which an SVG arc cannot express, so it is drawn as a circle.
 */
function quarterGlyph(box: PrintBox, size: number, index: number, count: number): string {
  const cx = box.x + box.width / 2;
  const cy = box.y + box.height / 2;
  const radius = size * 0.43;
  const fraction = clamp((index + 1) / Math.max(1, count - 1), 0, 1);
  const disc =
    `<circle cx="${num(cx)}" cy="${num(cy)}" r="${num(radius)}" fill="${ICON_OFF}"` +
    ` stroke="${ICON_OUTLINE}" stroke-width="${num(Math.max(0.5, size / 14))}"/>`;
  if (fraction >= 1) {
    return `${disc}<circle cx="${num(cx)}" cy="${num(cy)}" r="${num(radius)}" fill="${ICON_DARK}"/>`;
  }
  const start = -Math.PI / 2;
  const end = start + Math.PI * 2 * fraction;
  const endX = cx + Math.cos(end) * radius;
  const endY = cy + Math.sin(end) * radius;
  const largeArc = fraction > 0.5 ? 1 : 0;
  return (
    disc +
    `<path d="M ${num(cx)} ${num(cy)} L ${num(cx)} ${num(cy - radius)}` +
    ` A ${num(radius)} ${num(radius)} 0 ${largeArc} 1 ${num(endX)} ${num(endY)} Z"` +
    ` fill="${ICON_DARK}"/>`
  );
}

function dotGlyph(box: PrintBox, size: number, color: string): string {
  const cx = box.x + box.width / 2;
  const cy = box.y + box.height / 2;
  return `<circle cx="${num(cx)}" cy="${num(cy)}" r="${num(size * 0.23)}" fill="${color}"/>`;
}

// ---------------------------------------------------------------------------
// Definitions: patterns, gradients, clip paths
// ---------------------------------------------------------------------------

type TileArt =
  | { readonly kind: 'wash'; readonly opacity: number }
  | { readonly kind: 'dots'; readonly points: readonly (readonly [number, number])[] }
  | {
      readonly kind: 'lines';
      readonly segments: readonly (readonly [number, number, number, number])[];
    };

/**
 * The tile art, in an eight-unit square. It mirrors `PATTERN_ART` in
 * `src/chrome/fill-patterns.ts` and `applyPattern` in the grid's canvas
 * renderer, so the preview dialog, the grid and print show the same texture.
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

/** An imported pattern the table does not name paints like a half wash. */
const UNKNOWN_PATTERN_ART: TileArt = { kind: 'wash', opacity: 0.5 };

/** The page's `<defs>`: patterns, gradients and clip paths, built on demand. */
class Definitions {
  private readonly prefix: string;
  private readonly entries: string[] = [];
  private readonly ids = new Map<string, string>();
  private readonly gradients = new Map<string, string>();

  constructor(prefix: string) {
    this.prefix = prefix;
  }

  pattern(pattern: string, color: string | null, background: string | null): string {
    const key = `pattern:${pattern}:${color ?? ''}:${background ?? ''}`;
    const existing = this.ids.get(key);
    if (existing !== undefined) return existing;
    const id = `${this.prefix}-pattern-${this.ids.size}`;
    this.ids.set(key, id);
    const art = PATTERN_ART[pattern] ?? UNKNOWN_PATTERN_ART;
    const tile = PATTERN_TILE;
    const backgroundRect =
      background === null
        ? ''
        : `<rect x="0" y="0" width="${num(tile)}" height="${num(tile)}" fill="${background}"/>`;
    const paint = color === null ? '' : tileBody(art, color, tile);
    this.entries.push(
      `<pattern id="${id}" patternUnits="userSpaceOnUse" width="${num(tile)}" height="${num(tile)}">` +
        `${backgroundRect}${paint}</pattern>`,
    );
    return id;
  }

  gradient(from: string, to: string): string {
    const key = `${from}:${to}`;
    const existing = this.gradients.get(key);
    if (existing !== undefined) return existing;
    const id = `${this.prefix}-gradient-${this.gradients.size}`;
    this.gradients.set(key, id);
    this.entries.push(
      `<linearGradient id="${id}" x1="0" y1="0" x2="1" y2="0">` +
        `<stop offset="0" stop-color="${from}"/>` +
        `<stop offset="1" stop-color="${to}"/></linearGradient>`,
    );
    return id;
  }

  render(clips: ClipRegistry): string {
    const clipDefs = clips.render();
    if (this.entries.length === 0 && clipDefs === '') return '';
    return `<defs>${this.entries.join('')}${clipDefs}</defs>`;
  }
}

/** One `<clipPath>` per distinct clip box; a text clip is a hard clip. */
class ClipRegistry {
  private readonly prefix: string;
  private readonly ids = new Map<string, string>();
  private readonly boxes: { readonly id: string; readonly box: PrintBox }[] = [];

  constructor(prefix: string) {
    this.prefix = prefix;
  }

  ensure(box: PrintBox): string {
    const key = `${num(box.x)}:${num(box.y)}:${num(box.width)}:${num(box.height)}`;
    const existing = this.ids.get(key);
    if (existing !== undefined) return existing;
    const id = `${this.prefix}-clip-${this.ids.size}`;
    this.ids.set(key, id);
    this.boxes.push({ id, box });
    return id;
  }

  render(): string {
    return this.boxes
      .map(({ id, box }) => `<clipPath id="${id}"><rect ${rectAttributes(box)}/></clipPath>`)
      .join('');
  }
}

function tileBody(art: TileArt, color: string, tile: number): string {
  const scale = (value: number): number => (value / 8) * tile;
  switch (art.kind) {
    case 'wash':
      return `<rect x="0" y="0" width="${num(tile)}" height="${num(tile)}" fill="${color}" fill-opacity="${art.opacity}"/>`;
    case 'dots':
      return `<g fill="${color}">${art.points
        .map(
          ([x, y]) =>
            `<rect x="${num(scale(x))}" y="${num(scale(y))}" width="${num(tile / 8)}" height="${num(tile / 8)}"/>`,
        )
        .join('')}</g>`;
    case 'lines':
      return `<g stroke="${color}" stroke-width="${num(tile / 8)}" fill="none">${art.segments
        .map(
          ([x1, y1, x2, y2]) =>
            `<line x1="${num(scale(x1))}" y1="${num(scale(y1))}" x2="${num(scale(x2))}" y2="${num(scale(y2))}"/>`,
        )
        .join('')}</g>`;
  }
}

// ---------------------------------------------------------------------------
// Small helpers
// ---------------------------------------------------------------------------

function chromiumPageBox(paper: PrintPaper): PrintPageBox | null {
  const portrait =
    Math.abs(paper.width - NOMINAL_A4.width) <= 0.05 &&
    Math.abs(paper.height - NOMINAL_A4.height) <= 0.05;
  if (portrait) return CHROMIUM_A4_PAGE;
  const landscape =
    Math.abs(paper.width - NOMINAL_A4.height) <= 0.05 &&
    Math.abs(paper.height - NOMINAL_A4.width) <= 0.05;
  if (landscape) return { width: CHROMIUM_A4_PAGE.height, height: CHROMIUM_A4_PAGE.width };
  return null;
}

function rectAttributes(box: PrintBox): string {
  return (
    `x="${num(box.x)}" y="${num(box.y)}" width="${num(Math.max(0, box.width))}"` +
    ` height="${num(Math.max(0, box.height))}"`
  );
}

/** Two decimals are all the plan carries; never emit an exponent. */
function num(value: number): string {
  if (!Number.isFinite(value)) return '0';
  const rounded = Math.round(value * 100) / 100;
  return String(Object.is(rounded, -0) ? 0 : rounded);
}

function anchorOf(anchor: PrintTextItem['anchor']): string {
  return anchor === 'middle' ? 'middle' : anchor === 'end' ? 'end' : 'start';
}

const HEX_COLOR = /^#(?:[0-9a-f]{3}|[0-9a-f]{6})$/i;

/** Only a plain hex colour passes; anything else falls back. */
function safeColor(value: string | null, fallback: string): string;
function safeColor(value: string | null, fallback: null): string | null;
function safeColor(value: string | null, fallback: string | null): string | null {
  if (value === null) return fallback;
  const trimmed = value.trim();
  return HEX_COLOR.test(trimmed) ? trimmed.toUpperCase() : fallback;
}

const SERIF_FAMILIES = /^(times( new roman)?|georgia|cambria|garamond|serif)$/i;
const MONO_FAMILIES = /^(courier( new)?|consolas|monaco|monospace)$/i;

/** A safe CSS font list: the sheet's family, then a generic fallback. */
function fontFamily(name: string): string {
  // Letters, digits, spaces and the punctuation real family names use. Quoted
  // names, semicolons, parentheses and braces never reach the attribute.
  const clean = name.replace(/[^\p{L}\p{N} ._-]/gu, '').trim();
  const generic = MONO_FAMILIES.test(clean)
    ? 'monospace'
    : SERIF_FAMILIES.test(clean)
      ? 'serif'
      : 'sans-serif';
  return clean === '' ? generic : `${clean}, ${generic}`;
}

function escapeText(value: string): string {
  return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function escapeAttribute(value: string): string {
  return escapeText(value).replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function round(value: number): number {
  return Math.round(value * 100) / 100;
}

function clamp(value: number, low: number, high: number): number {
  if (!Number.isFinite(value)) return low;
  return Math.min(high, Math.max(low, value));
}
