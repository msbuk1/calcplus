import type { ProjectedCell, ProjectedStyle } from '@view/types';

/**
 * Text layout helpers shared by the canvas renderer and autofit.
 *
 * Everything here is pure and takes a measure callback, so the sizing maths
 * can be checked without a canvas. The renderer supplies the real
 * `measureText` function.
 */

export const TEXT_PADDING = 4;
export const INDENT_WIDTH = 8;
/** Excel's line box is about 1.25 times the font size. */
export const LINE_SPACING = 1.25;
export const MIN_FONT_PIXELS = 6;
export const DEFAULT_TEXT_COLOR = '#202020';

export type TextMeasure = (text: string, font: string) => number;
export type HorizontalTextAlignment = 'left' | 'center' | 'right';
export type VerticalTextAlignment = 'top' | 'center' | 'bottom';

export type TextRotation =
  | { readonly kind: 'none' }
  | { readonly kind: 'angle'; readonly radians: number }
  | { readonly kind: 'stacked' };

/**
 * A font size in points, rendered at 96 dpi at the given display zoom and
 * rounded to a quarter pixel.
 */
export function fontSizePixels(style: ProjectedStyle, zoom = 1): number {
  const raw = ((Math.max(1, style.font.size) * 96) / 72) * zoom;
  return Math.max(0.5, Math.round(raw * 4) / 4);
}

/** The CSS font shorthand for a projected style. */
export function fontFor(style: ProjectedStyle, sizePixels = fontSizePixels(style)): string {
  const italic = style.font.italic ? 'italic ' : '';
  const bold = style.font.bold ? 'bold ' : '';
  const name = style.font.name.replaceAll('"', '\\"');
  const family = name === '' ? 'sans-serif' : `"${name}", system-ui, sans-serif`;
  return `${italic}${bold}${formatNumber(sizePixels)}px ${family}`;
}

/** The colour the cell text is painted in: number-format colour, then font. */
export function textColorFor(style: ProjectedStyle): string {
  return style.numberFormatColor ?? style.font.color ?? DEFAULT_TEXT_COLOR;
}

/**
 * Horizontal alignment. Excel aligns by value type when the style says
 * `general`: numbers and errors right, booleans centred and text left.
 */
export function horizontalFor(
  style: ProjectedStyle,
  value: ProjectedCell['value'],
): HorizontalTextAlignment {
  const horizontal = style.alignment.horizontal;
  if (horizontal === 'left' || horizontal === 'fill' || horizontal === 'justify') return 'left';
  if (
    horizontal === 'center' ||
    horizontal === 'centerContinuous' ||
    horizontal === 'distributed'
  ) {
    return 'center';
  }
  if (horizontal === 'right') return 'right';
  switch (value.type) {
    case 'number':
    case 'error':
      return 'right';
    case 'boolean':
      return 'center';
    default:
      return 'left';
  }
}

/** Vertical alignment. Justified and distributed text starts at the top. */
export function verticalFor(style: ProjectedStyle): VerticalTextAlignment {
  const vertical = style.alignment.vertical;
  if (vertical === 'center') return 'center';
  if (vertical === 'top' || vertical === 'justify' || vertical === 'distributed') return 'top';
  return 'bottom';
}

/**
 * Excel's text rotation codes: 1-90 counterclockwise, 91-180 clockwise
 * (91 is one degree clockwise), 255 stacked vertically.
 */
export function rotationFor(style: ProjectedStyle): TextRotation {
  const rotation = style.alignment.textRotation;
  if (rotation === 255) return { kind: 'stacked' };
  if (rotation > 0 && rotation <= 90) {
    return { kind: 'angle', radians: (-rotation * Math.PI) / 180 };
  }
  if (rotation > 90 && rotation <= 180) {
    return { kind: 'angle', radians: ((rotation - 90) * Math.PI) / 180 };
  }
  return { kind: 'none' };
}

/** Wrap a display string into lines no wider than `maxWidth`. */
export function wrapTextLines(
  text: string,
  maxWidth: number,
  measure: (text: string) => number,
): string[] {
  if (text === '') return [''];
  const lines: string[] = [];
  for (const paragraph of text.split(/\r?\n/)) {
    if (paragraph === '') {
      lines.push('');
      continue;
    }
    const words = paragraph.split(/(\s+)/).filter((part) => part !== '');
    let line = '';
    for (const word of words) {
      const candidate = line + word;
      if (line !== '' && measure(candidate) > maxWidth) {
        lines.push(line.trimEnd());
        line = word.trimStart();
        if (line === '') continue;
      } else {
        line = candidate;
      }
      // A single word wider than the cell breaks by character, as Excel does
      // when the column is narrower than the longest word.
      if (measure(line) > maxWidth) {
        const broken = breakLongToken(line, maxWidth, measure);
        for (let index = 0; index < broken.length - 1; index += 1) {
          lines.push(broken[index] ?? '');
        }
        line = broken[broken.length - 1] ?? '';
      }
    }
    lines.push(line.trimEnd());
  }
  return lines.length === 0 ? [''] : lines;
}

/** The largest font size at or below the style's size that fits `maxWidth`. */
export function fitFontSize(
  style: ProjectedStyle,
  text: string,
  maxWidth: number,
  measure: TextMeasure,
  minimum = MIN_FONT_PIXELS,
  zoom = 1,
): number {
  const original = fontSizePixels(style, zoom);
  if (measure(text, fontFor(style, original)) <= maxWidth) return original;
  const minimumSize = Math.max(0.5, minimum * zoom);
  let low = Math.min(minimumSize, original);
  let high = original;
  let best = low;
  while (high - low > 0.25) {
    const middle = (low + high) / 2;
    if (measure(text, fontFor(style, middle)) <= maxWidth) {
      best = middle;
      low = middle;
    } else {
      high = middle;
    }
  }
  return Math.max(minimumSize, best);
}

/** The height of one line box for a size in pixels. */
export function lineHeightFor(sizePixels: number): number {
  return Math.max(1, sizePixels * LINE_SPACING);
}

/** The horizontal padding a cell adds before its text, at the display zoom. */
export function horizontalPadding(zoom = 1): number {
  return TEXT_PADDING * zoom;
}

/** The indentation the aligned side adds, in pixels at the display zoom. */
export function horizontalIndent(style: ProjectedStyle, zoom = 1): number {
  return Math.max(0, style.alignment.indent) * INDENT_WIDTH * zoom;
}

/**
 * The left and right insets for a cell's text. Excel indents only the side the
 * text is aligned to; centring uses the ordinary padding on both sides.
 */
export function textInsets(
  style: ProjectedStyle,
  horizontal: HorizontalTextAlignment,
  zoom = 1,
): { readonly left: number; readonly right: number } {
  const padding = TEXT_PADDING * zoom;
  const indent = horizontalIndent(style, zoom);
  if (horizontal === 'left') return { left: padding + indent, right: padding };
  if (horizontal === 'right') return { left: padding, right: padding + indent };
  return { left: padding, right: padding };
}

function breakLongToken(
  token: string,
  maxWidth: number,
  measure: (text: string) => number,
): string[] {
  const lines: string[] = [];
  let current = '';
  for (const character of token) {
    const candidate = current + character;
    if (current !== '' && measure(candidate) > maxWidth) {
      lines.push(current);
      current = character;
    } else {
      current = candidate;
    }
  }
  lines.push(current);
  return lines;
}

function formatNumber(value: number): string {
  return Number.isInteger(value) ? String(value) : String(Math.round(value * 100) / 100);
}
