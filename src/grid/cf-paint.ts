import type {
  CfBorderEdgePatch,
  CfBorderPatch,
  CfCellEvaluation,
  CfColorScaleOverlay,
  CfDataBarOverlay,
  CfIconSetId,
  CfOverlay,
  ProjectedStyle,
} from '@view/types';
import type { CellRect } from './panes';

/**
 * Conditional-formatting painting (conditional formatting §5).
 *
 * The evaluation engine hands the renderer paint-ready overlays; this module
 * turns one cell's overlays into a *plan* — an ordered list of paint ops — and
 * paints the three pieces the canvas renderer has no other reason to know
 * about: the fill, the data bar and the icon.
 *
 * §5.1 order, per cell, is the order of the ops here:
 *
 *   1. manual fill        (the renderer's own pass, before this plan)
 *   2. CF fill / colour-scale fill
 *   3. data bar
 *   4. manual border      (the renderer's own pass, before the CF border op)
 *   5. CF border
 *   6. cell text, shifted right when an icon is shown
 *   7. CF font and CF number format (carried by the text op)
 *   8. icon
 *   9. selection overlay  (the renderer's own pass, after every cell)
 *
 * Everything here is derived: the plan never mutates a cell, its style or its
 * value, and a merged range is painted from its anchor's plan (§5.2). Bar
 * widths, icon sizes and CF fonts scale with zoom; thresholds never do, because
 * the threshold maths belongs to the engine (§5.4).
 */

/** Below this zoom the icons are suppressed; fills, bars and scales still paint. */
export const CF_ICON_MIN_ZOOM = 0.25;
/** Excel's default icon box is about 12 device pixels at 100%. */
export const CF_ICON_BASE_SIZE = 12;

const ICON_RED = '#F8696B';
const ICON_AMBER = '#FFEB84';
const ICON_GREEN = '#63BE7B';
const ICON_GREY_DARK = '#7F7F7F';
const ICON_GREY_MID = '#A6A6A6';
const ICON_GREY_LIGHT = '#D9D9D9';

export type CfPaintOpKind = 'fill' | 'dataBar' | 'text' | 'border' | 'icon';

/** The fill under everything else: a CF fill or a colour-scale blend. */
export interface CfFillPaint {
  readonly kind: 'fill';
  readonly rect: CellRect;
  readonly color: string;
}

/** The bar itself, already in canvas coordinates. */
export interface CfDataBarPaint {
  readonly kind: 'dataBar';
  readonly rect: CellRect;
  readonly color: string;
  readonly fill: 'gradient' | 'solid';
  readonly border: boolean;
  readonly negative: boolean;
  /** The zero axis, or `null` when the range does not straddle zero. */
  readonly axisX: number | null;
  readonly axisColor: string;
}

/** How the cell text is painted: shifted for an icon, restyled by a CF dxf. */
export interface CfTextPaint {
  readonly kind: 'text';
  readonly hidden: boolean;
  readonly shift: number;
  readonly style: ProjectedStyle;
  readonly text: string;
  /**
   * `true` when the projection's width-aware layout still applies (`*` fill,
   * `_` skip, the `#` overflow marker). `false` when a CF number format
   * produced its own text, which paints exactly as formatted.
   */
  readonly layoutText: boolean;
}

/** The CF border, painted over the manual one. Only the edges the dxf sets. */
export interface CfBorderPaint {
  readonly kind: 'border';
  readonly border: ProjectedStyle['border'];
}

export interface CfIconBox {
  readonly x: number;
  readonly y: number;
  readonly size: number;
}

export interface CfIconPaint {
  readonly kind: 'icon';
  readonly iconSet: CfIconSetId;
  readonly iconIndex: number;
  readonly iconCount: number;
  readonly box: CfIconBox;
}

export type CfPaintOp = CfFillPaint | CfDataBarPaint | CfTextPaint | CfBorderPaint | CfIconPaint;

export interface CfPaintInput {
  readonly evaluation: CfCellEvaluation;
  /** The cell's rectangle, or the whole merged rectangle at its anchor (§5.2). */
  readonly rect: CellRect;
  readonly zoom: number;
  /** The cell's manual style; the plan never replaces it, only overlays it. */
  readonly style: ProjectedStyle;
  /** The projection's display text. */
  readonly displayText: string;
  /**
   * The text a CF number format produces, when the rule sets one. `null` or
   * omitted keeps the projected display text and its width-aware layout.
   */
  readonly cfDisplayText?: string | null;
}

/**
 * Plan one cell's conditional formatting, in §5.1 order. An unpainted cell
 * produces no ops at all, so the renderer's fast path stays fast.
 */
export function cfPaintOps(input: CfPaintInput): readonly CfPaintOp[] {
  const { evaluation, rect, zoom, style, displayText } = input;
  if (!evaluation.painted) return [];
  const ops: CfPaintOp[] = [];
  const fillColor = cfFillColor(evaluation.overlay, evaluation.colorScale);
  if (fillColor !== null) ops.push({ kind: 'fill', rect, color: fillColor });

  const bar = evaluation.dataBar;
  if (bar !== null) ops.push(dataBarPaint(bar, rect));

  const paintedStyle = cfTextStyle(style, evaluation.overlay);
  const icon = evaluation.iconSet;
  const iconSuppressed = icon !== null && iconsSuppressed(zoom);
  ops.push({
    kind: 'text',
    // An icon-only rule and a hidden data-bar value both suppress the text.
    hidden:
      (icon !== null && (icon.showValue === false || icon.iconOnly)) ||
      (bar !== null && !bar.showValue),
    shift: icon !== null && !iconSuppressed ? cfIconShift(rect, zoom) : 0,
    style: paintedStyle,
    text: input.cfDisplayText ?? displayText,
    layoutText: input.cfDisplayText === undefined || input.cfDisplayText === null,
  });

  const border = cfBorderStyle(evaluation.overlay.border);
  if (border !== null) ops.push({ kind: 'border', border });

  if (icon !== null && !iconSuppressed) {
    ops.push({
      kind: 'icon',
      iconSet: icon.iconSet,
      iconIndex: icon.iconIndex,
      iconCount: icon.iconCount,
      box: cfIconBox(rect, zoom),
    });
  }

  return ops;
}

/** The fill colour for a cell: a CF fill first, then a colour-scale blend. */
export function cfFillColor(
  overlay: CfOverlay,
  colorScale: CfColorScaleOverlay | null,
): string | null {
  const fill = overlay.fill;
  if (fill !== undefined && fill !== null) {
    // A dxf fill is a pattern fill in OOXML. Conditional fills in practice are
    // solid; a pattern's foreground colour painted solid is the closest
    // approximation the plan can make without a pattern tile (§11).
    const color = fill.color ?? fill.backgroundColor ?? null;
    if (color !== null) return color;
  }
  return colorScale === null ? null : colorScale.color;
}

/**
 * The data bar in canvas coordinates. The overlay's fractions are of the
 * cell — or merged — width, so a merged bar scales with the merged rectangle
 * (§5.2) and a zoomed cell's bar scales with the geometry (§5.4).
 */
export function cfDataBarRect(bar: CfDataBarOverlay, rect: CellRect): CellRect {
  const { start, end } = mirroredBar(bar);
  const left = rect.x + Math.min(start, end) * rect.width;
  const width = Math.max(0, Math.abs(end - start) * rect.width);
  // Excel insets the bar a pixel inside the row so the gridline still reads.
  return { x: left, y: rect.y + 1, width, height: Math.max(1, rect.height - 2) };
}

/** The zero axis, mirrored with the bar for a right-to-left rule. */
export function cfDataBarAxisX(bar: CfDataBarOverlay, rect: CellRect): number {
  return rect.x + mirroredBar(bar).axis * rect.width;
}

function mirroredBar(bar: CfDataBarOverlay): { start: number; end: number; axis: number } {
  if (bar.direction !== 'rightToLeft') {
    return { start: bar.start, end: bar.end, axis: bar.axis };
  }
  return { start: 1 - bar.end, end: 1 - bar.start, axis: 1 - bar.axis };
}

function dataBarPaint(bar: CfDataBarOverlay, rect: CellRect): CfDataBarPaint {
  // The axis is the zero line: always drawn when the range straddles zero (a
  // mixed-sign range is exactly where it carries meaning), or when the rule
  // names an axis colour.
  const mixed = bar.min < 0 && bar.max > 0;
  return {
    kind: 'dataBar',
    rect: cfDataBarRect(bar, rect),
    color: bar.color,
    fill: bar.fill,
    border: bar.border,
    negative: bar.negative,
    axisX: bar.axisColor !== null || mixed ? cfDataBarAxisX(bar, rect) : null,
    axisColor: bar.axisColor ?? '#808080',
  };
}

/** Whether icons are suppressed at this zoom (§5.4: below 25%). */
export function iconsSuppressed(zoom: number): boolean {
  return zoom < CF_ICON_MIN_ZOOM;
}

/** The icon's square, at the left of the cell or merged rectangle. */
export function cfIconBox(rect: CellRect, zoom: number): CfIconBox {
  const size = Math.max(4, Math.min(Math.max(4, rect.height - 2), CF_ICON_BASE_SIZE * zoom));
  return { x: rect.x + 1, y: rect.y + (rect.height - size) / 2, size };
}

/** How far the text shifts right to clear the icon. */
export function cfIconShift(rect: CellRect, zoom: number): number {
  return cfIconBox(rect, zoom).size + Math.max(2, 2 * zoom);
}

/**
 * The style the text paints with: the manual style with only the properties the
 * CF dxf sets laid over it (§3.1). Manual formatting shows through otherwise.
 */
export function cfTextStyle(style: ProjectedStyle, overlay: CfOverlay): ProjectedStyle {
  const font = overlay.font;
  const numberFormat = overlay.numberFormat;
  if (font === null && numberFormat === null) return style;
  const patched = font === null ? style.font : { ...style.font, ...definedFont(font) };
  return {
    ...style,
    font: patched,
    numberFormat: numberFormat?.code ?? style.numberFormat,
  };
}

function definedFont(font: NonNullable<CfOverlay['font']>): Partial<ProjectedStyle['font']> {
  const patch: Record<string, unknown> = {};
  if (font.name !== undefined) patch.name = font.name;
  if (font.size !== undefined) patch.size = font.size;
  if (font.bold !== undefined) patch.bold = font.bold;
  if (font.italic !== undefined) patch.italic = font.italic;
  if (font.underline !== undefined) patch.underline = font.underline;
  if (font.underlineStyle !== undefined) patch.underlineStyle = font.underlineStyle;
  if (font.strike !== undefined) patch.strike = font.strike;
  if (font.color !== undefined) patch.color = font.color;
  return patch as Partial<ProjectedStyle['font']>;
}

const NO_EDGE = Object.freeze({ style: 'none', color: null, weight: null });

/**
 * The CF border as a paint-ready border: only the edges the dxf sets, so the
 * manual border stays visible everywhere else and CF draws over it where it
 * does set an edge (§3.1, §5.1).
 */
export function cfBorderStyle(border: CfBorderPatch | null): ProjectedStyle['border'] | null {
  if (border === null) return null;
  const left = edgeOf(border.left);
  const right = edgeOf(border.right);
  const top = edgeOf(border.top);
  const bottom = edgeOf(border.bottom);
  const diagonal = edgeOf(border.diagonal);
  const diagonalUp = border.diagonalUp === true && diagonal.style !== 'none';
  const diagonalDown = border.diagonalDown === true && diagonal.style !== 'none';
  if (
    left.style === 'none' &&
    right.style === 'none' &&
    top.style === 'none' &&
    bottom.style === 'none' &&
    diagonal.style === 'none'
  ) {
    return null;
  }
  return { left, right, top, bottom, diagonal, diagonalUp, diagonalDown };
}

function edgeOf(edge: CfBorderEdgePatch | undefined): ProjectedStyle['border']['left'] {
  if (edge === undefined) return NO_EDGE;
  return {
    style: edge.style ?? 'thin',
    color: edge.color ?? null,
    weight: edge.weight ?? null,
  };
}

// ---------------------------------------------------------------------------
// Painting
// ---------------------------------------------------------------------------

/** Paint a CF fill. It covers the gridlines drawn underneath it (§5.1). */
export function paintCfFill(context: CanvasRenderingContext2D, op: CfFillPaint): void {
  const { x, y, width, height } = op.rect;
  if (width <= 0 || height <= 0) return;
  context.save();
  context.fillStyle = op.color;
  context.fillRect(x, y, width, height);
  context.restore();
}

/**
 * Paint a data bar: a gradient or solid bar, its optional border, and the zero
 * axis. The bar sits *behind* the text, which the renderer paints afterwards.
 */
export function paintCfDataBar(context: CanvasRenderingContext2D, op: CfDataBarPaint): void {
  const { x, y, width, height } = op.rect;
  if (width <= 0 || height <= 0) return;
  context.save();
  if (op.fill === 'gradient' && width > 1) {
    // Excel fades a gradient bar from the value's tip towards the axis. The
    // exact ramp is Excel's; a near-identical one avoids copying its art.
    const light = lighten(op.color, 0.45);
    const gradient = context.createLinearGradient(x, y, x + width, y);
    if (op.negative) {
      gradient.addColorStop(0, op.color);
      gradient.addColorStop(1, light);
    } else {
      gradient.addColorStop(0, light);
      gradient.addColorStop(1, op.color);
    }
    context.fillStyle = gradient;
  } else {
    context.fillStyle = op.color;
  }
  context.fillRect(x, y, width, height);
  if (op.border) {
    context.strokeStyle = darken(op.color, 0.25);
    context.lineWidth = 1;
    context.strokeRect(x + 0.5, y + 0.5, Math.max(0, width - 1), Math.max(0, height - 1));
  }
  if (op.axisX !== null) {
    const axis = snapToPixel(op.axisX, 1);
    context.strokeStyle = op.axisColor;
    context.lineWidth = 1;
    context.beginPath();
    context.moveTo(axis, y);
    context.lineTo(axis, y + height);
    context.stroke();
  }
  context.restore();
}

/**
 * Paint one icon set glyph. The glyphs are canvas paths, not Microsoft art:
 * near-identical families (arrows, lights, signs, flags, ratings, quarters and
 * dots) at Excel's size, coloured by the set (§3.3, §11).
 */
export function paintCfIcon(context: CanvasRenderingContext2D, op: CfIconPaint): void {
  const style = ICON_STYLES[op.iconSet];
  if (style === undefined) return;
  const index = Math.min(op.iconCount - 1, Math.max(0, op.iconIndex));
  const color = style.colors[Math.min(style.colors.length - 1, index)] ?? ICON_GREY_DARK;
  context.save();
  context.lineWidth = Math.max(1, op.box.size / 10);
  switch (style.family) {
    case 'arrow':
      paintArrow(context, op.box, color, arrowAngle(index, op.iconCount));
      break;
    case 'traffic':
      paintTrafficLight(context, op.box, color);
      break;
    case 'sign':
      paintSign(context, op.box, index, op.iconCount);
      break;
    case 'flag':
      paintFlag(context, op.box, color);
      break;
    case 'rating':
      paintRating(context, op.box, index + 1, op.iconCount);
      break;
    case 'quarter':
      paintQuarter(context, op.box, index + 1, op.iconCount);
      break;
    case 'dot':
      paintDot(context, op.box, color);
      break;
  }
  context.restore();
}

type IconFamily = 'arrow' | 'traffic' | 'sign' | 'flag' | 'rating' | 'quarter' | 'dot';

const ICON_STYLES: Readonly<
  Record<CfIconSetId, { family: IconFamily; colors: readonly string[] }>
> = Object.freeze({
  '3Arrows': { family: 'arrow', colors: [ICON_RED, ICON_AMBER, ICON_GREEN] },
  '3ArrowsGray': {
    family: 'arrow',
    colors: [ICON_GREY_LIGHT, ICON_GREY_MID, ICON_GREY_DARK],
  },
  '4Arrows': { family: 'arrow', colors: [ICON_RED, '#FFB628', ICON_AMBER, ICON_GREEN] },
  '4ArrowsGray': {
    family: 'arrow',
    colors: [ICON_GREY_LIGHT, '#B0B0B0', ICON_GREY_MID, ICON_GREY_DARK],
  },
  '5Arrows': {
    family: 'arrow',
    colors: [ICON_RED, '#F9966B', ICON_AMBER, ICON_GREEN, '#2E9E4F'],
  },
  '5ArrowsGray': {
    family: 'arrow',
    colors: [ICON_GREY_LIGHT, '#C4C4C4', ICON_GREY_MID, '#8C8C8C', ICON_GREY_DARK],
  },
  '3TrafficLights1': { family: 'traffic', colors: [ICON_RED, ICON_AMBER, ICON_GREEN] },
  '3TrafficLights2': { family: 'traffic', colors: [ICON_RED, ICON_AMBER, ICON_GREEN] },
  '4TrafficLights': {
    family: 'traffic',
    colors: ['#000000', ICON_RED, ICON_AMBER, ICON_GREEN],
  },
  '3Flags': { family: 'flag', colors: [ICON_RED, ICON_AMBER, ICON_GREEN] },
  '3Signs': { family: 'sign', colors: [ICON_RED, ICON_AMBER, ICON_GREEN] },
  '3Symbols': { family: 'sign', colors: [ICON_GREEN, ICON_AMBER, ICON_RED] },
  '3Symbols2': { family: 'sign', colors: [ICON_GREEN, ICON_AMBER, ICON_RED] },
  '4Rating': {
    family: 'rating',
    colors: [ICON_GREY_LIGHT, ICON_GREY_MID, ICON_GREY_DARK, '#000000'],
  },
  '5Rating': {
    family: 'rating',
    colors: [ICON_GREY_LIGHT, ICON_GREY_LIGHT, ICON_GREY_MID, ICON_GREY_DARK, '#000000'],
  },
  '5Quarters': {
    family: 'quarter',
    colors: [ICON_GREY_LIGHT, ICON_GREY_LIGHT, ICON_GREY_MID, ICON_GREY_DARK, '#000000'],
  },
  '4RedToBlack': {
    family: 'dot',
    colors: ['#000000', '#4D4D4D', '#A60000', ICON_RED],
  },
});

/** Up, sideways, down for three; a progressively tilted fan beyond that. */
function arrowAngle(index: number, count: number): number {
  if (count <= 1) return 0;
  const spread = Math.PI / 2; // 90 degrees from straight up to straight down
  return -spread / 2 + (spread * index) / (count - 1);
}

function paintArrow(
  context: CanvasRenderingContext2D,
  box: CfIconBox,
  color: string,
  angle: number,
): void {
  const half = box.size / 2;
  context.translate(box.x + half, box.y + half);
  context.rotate(angle + Math.PI / 2);
  context.fillStyle = color;
  context.beginPath();
  context.moveTo(0, -half * 0.95);
  context.lineTo(half * 0.62, -half * 0.12);
  context.lineTo(half * 0.2, -half * 0.12);
  context.lineTo(half * 0.2, half * 0.9);
  context.lineTo(-half * 0.2, half * 0.9);
  context.lineTo(-half * 0.2, -half * 0.12);
  context.lineTo(-half * 0.62, -half * 0.12);
  context.closePath();
  context.fill();
}

function paintTrafficLight(context: CanvasRenderingContext2D, box: CfIconBox, color: string): void {
  const half = box.size / 2;
  context.fillStyle = color;
  context.beginPath();
  context.arc(box.x + half, box.y + half, half * 0.78, 0, Math.PI * 2);
  context.fill();
  context.strokeStyle = '#404040';
  context.stroke();
}

function paintSign(
  context: CanvasRenderingContext2D,
  box: CfIconBox,
  index: number,
  count: number,
): void {
  const half = box.size / 2;
  const centerX = box.x + half;
  const centerY = box.y + half;
  const last = count - 1;
  if (index === 0) {
    // Red disc with a cross.
    context.fillStyle = ICON_RED;
    context.beginPath();
    context.arc(centerX, centerY, half * 0.82, 0, Math.PI * 2);
    context.fill();
    context.strokeStyle = '#ffffff';
    context.lineWidth = Math.max(1, box.size / 9);
    const arm = half * 0.36;
    context.beginPath();
    context.moveTo(centerX - arm, centerY - arm);
    context.lineTo(centerX + arm, centerY + arm);
    context.moveTo(centerX + arm, centerY - arm);
    context.lineTo(centerX - arm, centerY + arm);
    context.stroke();
    return;
  }
  if (index === last) {
    // Green disc with a tick.
    context.fillStyle = ICON_GREEN;
    context.beginPath();
    context.arc(centerX, centerY, half * 0.82, 0, Math.PI * 2);
    context.fill();
    context.strokeStyle = '#ffffff';
    context.lineWidth = Math.max(1, box.size / 9);
    context.beginPath();
    context.moveTo(centerX - half * 0.36, centerY);
    context.lineTo(centerX - half * 0.08, centerY + half * 0.32);
    context.lineTo(centerX + half * 0.4, centerY - half * 0.32);
    context.stroke();
    return;
  }
  // Amber triangle with a bar and a dot.
  context.fillStyle = ICON_AMBER;
  context.beginPath();
  context.moveTo(centerX, centerY - half * 0.86);
  context.lineTo(centerX + half * 0.86, centerY + half * 0.72);
  context.lineTo(centerX - half * 0.86, centerY + half * 0.72);
  context.closePath();
  context.fill();
  context.strokeStyle = '#333333';
  context.lineWidth = Math.max(1, box.size / 3);
  context.beginPath();
  context.moveTo(centerX, centerY - half * 0.34);
  context.lineTo(centerX, centerY + half * 0.12);
  context.stroke();
  context.beginPath();
  context.arc(centerX, centerY + half * 0.42, Math.max(0.7, box.size / 10), 0, Math.PI * 2);
  context.fillStyle = '#333333';
  context.fill();
}

function paintFlag(context: CanvasRenderingContext2D, box: CfIconBox, color: string): void {
  const half = box.size / 2;
  context.strokeStyle = '#404040';
  context.lineWidth = Math.max(1, box.size / 12);
  context.beginPath();
  context.moveTo(box.x + half * 0.4, box.y + half * 0.9);
  context.lineTo(box.x + half * 0.4, box.y + half * 0.1);
  context.stroke();
  context.fillStyle = color;
  context.beginPath();
  context.moveTo(box.x + half * 0.4, box.y + half * 0.1);
  context.lineTo(box.x + half * 1.72, box.y + half * 0.42);
  context.lineTo(box.x + half * 0.4, box.y + half * 0.74);
  context.closePath();
  context.fill();
}

const RATING_OFF = '#E3E3E3';

function paintRating(
  context: CanvasRenderingContext2D,
  box: CfIconBox,
  filled: number,
  count: number,
): void {
  const slots = Math.max(1, count);
  const gap = Math.max(1, box.size / 12);
  const barWidth = (box.size - gap * (slots - 1)) / slots;
  for (let slot = 0; slot < slots; slot += 1) {
    // A taller bar the higher the rating; low bars stay grey.
    const height = box.size * (0.35 + (0.65 * (slot + 1)) / slots);
    context.fillStyle = slot < filled ? '#404040' : RATING_OFF;
    context.fillRect(
      box.x + slot * (barWidth + gap),
      box.y + box.size - height,
      Math.max(1, barWidth),
      height,
    );
  }
}

function paintQuarter(
  context: CanvasRenderingContext2D,
  box: CfIconBox,
  filled: number,
  count: number,
): void {
  const half = box.size / 2;
  const centerX = box.x + half;
  const centerY = box.y + half;
  const radius = half * 0.86;
  context.fillStyle = RATING_OFF;
  context.beginPath();
  context.arc(centerX, centerY, radius, 0, Math.PI * 2);
  context.fill();
  context.strokeStyle = '#808080';
  context.lineWidth = Math.max(1, box.size / 14);
  context.stroke();
  const fraction = Math.min(1, Math.max(0, filled / Math.max(1, count - 1)));
  if (fraction === 0) return;
  context.fillStyle = '#404040';
  context.beginPath();
  context.moveTo(centerX, centerY);
  context.arc(centerX, centerY, radius, -Math.PI / 2, -Math.PI / 2 + Math.PI * 2 * fraction);
  context.closePath();
  context.fill();
}

function paintDot(context: CanvasRenderingContext2D, box: CfIconBox, color: string): void {
  const half = box.size / 2;
  context.fillStyle = color;
  context.beginPath();
  context.arc(box.x + half, box.y + half, half * 0.46, 0, Math.PI * 2);
  context.fill();
  context.strokeStyle = color;
  context.stroke();
}

// ---------------------------------------------------------------------------
// Small colour helpers
// ---------------------------------------------------------------------------

function parseColor(color: string): { red: number; green: number; blue: number } | null {
  const hex = color.trim().replace(/^#/, '');
  const normalised =
    hex.length === 8 ? hex.slice(2) : hex.length === 6 ? hex : hex.length === 3 ? hex : null;
  if (normalised === null) return null;
  const expanded =
    normalised.length === 3
      ? normalised
          .split('')
          .map((part) => part + part)
          .join('')
      : normalised;
  return {
    red: Number.parseInt(expanded.slice(0, 2), 16),
    green: Number.parseInt(expanded.slice(2, 4), 16),
    blue: Number.parseInt(expanded.slice(4, 6), 16),
  };
}

function formatColor(red: number, green: number, blue: number): string {
  const channel = (value: number): string =>
    Math.round(Math.min(255, Math.max(0, value)))
      .toString(16)
      .padStart(2, '0')
      .toUpperCase();
  return `#${channel(red)}${channel(green)}${channel(blue)}`;
}

/** Mix a colour towards white; `amount` 0 is the colour, 1 is white. */
export function lighten(color: string, amount: number): string {
  const rgb = parseColor(color);
  if (rgb === null) return color;
  return formatColor(
    rgb.red + (255 - rgb.red) * amount,
    rgb.green + (255 - rgb.green) * amount,
    rgb.blue + (255 - rgb.blue) * amount,
  );
}

/** Mix a colour towards black; `amount` 0 is the colour, 1 is black. */
export function darken(color: string, amount: number): string {
  const rgb = parseColor(color);
  if (rgb === null) return color;
  return formatColor(rgb.red * (1 - amount), rgb.green * (1 - amount), rgb.blue * (1 - amount));
}

function snapToPixel(value: number, lineWidth: number): number {
  return lineWidth % 2 === 1 ? Math.round(value) + 0.5 : Math.round(value);
}
