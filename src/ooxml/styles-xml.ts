import type {
  AlignmentStyle,
  BorderEdgeStyle,
  BorderStyle,
  FillStyle,
  FontStyle,
  NamedStyleRecord,
  StyleRecord,
} from '@model/types';
import { createDefaultStyle } from '@model/workbook';
import {
  DEFAULT_ALIGNMENT,
  DEFAULT_BORDER,
  DEFAULT_FILL,
  DEFAULT_FONT,
  builtinNumberFormat,
  normalizeHorizontalAlignment,
  normalizeVerticalAlignment,
  resolveColor,
} from '@model/style';
import { decodeXml, parseXml, openTagStart } from './xml';

export interface StylesXmlInfo {
  readonly styles: StyleRecord[];
  readonly namedStyles: NamedStyleRecord[];
}

interface OpenElement {
  readonly local: string;
  readonly start: number;
}

interface FontAccumulator {
  name: string | null;
  size: number | null;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  color: ReturnType<typeof resolveColor>;
}

interface FillAccumulator {
  pattern: string | null;
  color: ReturnType<typeof resolveColor>;
}

type BorderEdgeKey = 'left' | 'right' | 'top' | 'bottom' | 'diagonal';

interface BorderAccumulator {
  left: BorderEdgeStyle;
  right: BorderEdgeStyle;
  top: BorderEdgeStyle;
  bottom: BorderEdgeStyle;
  diagonal: BorderEdgeStyle;
  diagonalUp: boolean;
  diagonalDown: boolean;
}

const NO_EDGE: BorderEdgeStyle = { style: 'none', color: null };

/**
 * Parse the workbook style sheet into the shared, append-only style table:
 * font, fill, border, alignment and the number-format reference (built-in id
 * plus authored code) for every cell style.
 */
export function parseStylesXml(bytes: Uint8Array, partName: string): StylesXmlInfo {
  const xml = decodeXml(bytes);
  const styles: StyleRecord[] = [];
  const namedStyles: NamedStyleRecord[] = [];
  const fonts: FontAccumulator[] = [];
  const fills: FillAccumulator[] = [];
  const borders: BorderAccumulator[] = [];
  const customFormats = new Map<number, string>();
  const stack: OpenElement[] = [];
  let inCellXfs = false;
  let inCellStyleXfs = false;
  let cellXfsDepth = -1;
  let cellStyleXfsDepth = -1;
  let inFonts = false;
  let inFills = false;
  let inBorders = false;
  let font: FontAccumulator | null = null;
  let fill: FillAccumulator | null = null;
  let border: BorderAccumulator | null = null;
  let borderEdge: BorderEdgeKey | null = null;

  parseXml(
    xml,
    {
      onOpenTag(tag, position) {
        const start = openTagStart(xml, position);
        stack.push({ local: tag.local, start });
        const depth = stack.length;
        const parent = stack.length >= 2 ? stack[stack.length - 2]!.local : null;
        if (tag.local === 'cellXfs') {
          inCellXfs = true;
          cellXfsDepth = depth;
        } else if (tag.local === 'cellStyleXfs') {
          inCellStyleXfs = true;
          cellStyleXfsDepth = depth;
        } else if (tag.local === 'fonts') {
          inFonts = true;
        } else if (tag.local === 'fills') {
          inFills = true;
        } else if (tag.local === 'borders') {
          inBorders = true;
        } else if (tag.local === 'numFmt') {
          const id = numberAttribute(tag, 'numFmtId', Number.NaN);
          const code = stringAttribute(tag, 'formatCode');
          if (Number.isInteger(id) && code !== null) customFormats.set(id, code);
        } else if (inFonts && tag.local === 'font' && parent === 'fonts') {
          font = {
            name: null,
            size: null,
            bold: false,
            italic: false,
            underline: false,
            strike: false,
            color: null,
          };
        } else if (font !== null) {
          applyFontChild(tag, font);
        } else if (inFills && tag.local === 'fill' && parent === 'fills') {
          fill = { pattern: null, color: null };
        } else if (fill !== null) {
          applyFillChild(tag, fill);
        } else if (inBorders && tag.local === 'border' && parent === 'borders') {
          border = {
            left: NO_EDGE,
            right: NO_EDGE,
            top: NO_EDGE,
            bottom: NO_EDGE,
            diagonal: NO_EDGE,
            diagonalUp: isTruthy(stringAttribute(tag, 'diagonalUp')),
            diagonalDown: isTruthy(stringAttribute(tag, 'diagonalDown')),
          };
          borderEdge = null;
        } else if (border !== null) {
          borderEdge = applyBorderChild(tag, border, borderEdge);
        } else if (inCellXfs && tag.local === 'xf') {
          styles.push({
            id: styles.length,
            numFmtId: numberAttribute(tag, 'numFmtId', 0),
            fontId: numberAttribute(tag, 'fontId', 0),
            fillId: numberAttribute(tag, 'fillId', 0),
            borderId: numberAttribute(tag, 'borderId', 0),
            xfId: numberAttribute(tag, 'xfId', 0),
            rawXml: '',
            font: DEFAULT_FONT,
            fill: DEFAULT_FILL,
            border: DEFAULT_BORDER,
            alignment: DEFAULT_ALIGNMENT,
            numFmtCode: null,
          });
        } else if (inCellXfs && tag.local === 'alignment') {
          const last = styles[styles.length - 1];
          if (last !== undefined) {
            styles[styles.length - 1] = { ...last, alignment: readAlignment(tag) };
          }
        } else if (inCellStyleXfs && tag.local === 'xf') {
          namedStyles.push({
            name: '',
            xfId: namedStyles.length,
            builtinId: null,
          });
        }
      },
      onCloseTag(tag, _start, end) {
        const depth = stack.length;
        const open = stack.pop();
        if (font !== null && tag.local === 'font') {
          fonts.push(font);
          font = null;
        } else if (fill !== null && tag.local === 'fill') {
          fills.push(fill);
          fill = null;
        } else if (border !== null && tag.local === 'border') {
          borders.push(border);
          border = null;
          borderEdge = null;
        }
        if (
          borderEdge !== null &&
          (tag.local === 'left' ||
            tag.local === 'right' ||
            tag.local === 'top' ||
            tag.local === 'bottom' ||
            tag.local === 'diagonal')
        ) {
          borderEdge = null;
        }
        if (inCellXfs && tag.local === 'xf' && open !== undefined) {
          const index = styles.length - 1;
          const last = styles[index];
          if (last !== undefined) {
            styles[index] = { ...last, rawXml: xml.slice(open.start, end) };
          }
        }
        if (tag.local === 'cellXfs' && depth === cellXfsDepth) {
          inCellXfs = false;
          cellXfsDepth = -1;
        }
        if (tag.local === 'cellStyleXfs' && depth === cellStyleXfsDepth) {
          inCellStyleXfs = false;
          cellStyleXfsDepth = -1;
        }
        if (tag.local === 'fonts') inFonts = false;
        if (tag.local === 'fills') inFills = false;
        if (tag.local === 'borders') inBorders = false;
      },
    },
    partName,
  );

  if (styles.length === 0) styles.push(createDefaultStyle(0));
  return {
    styles: styles.map((style, index) => ({
      ...style,
      id: index,
      font: resolveFont(fonts[style.fontId]),
      fill: resolveFill(fills[style.fillId]),
      border: resolveBorder(borders[style.borderId]),
      numFmtCode: resolveNumberFormat(style.numFmtId, customFormats),
    })),
    namedStyles,
  };
}

function applyFontChild(
  tag: { local: string; attributes: Record<string, { local: string; value: string }> },
  font: FontAccumulator,
): void {
  switch (tag.local) {
    case 'b':
    case 'bCs':
      font.bold = true;
      break;
    case 'i':
    case 'iCs':
      font.italic = true;
      break;
    case 'u':
      font.underline = stringAttribute(tag, 'val') !== 'none';
      break;
    case 'strike':
    case 'strikeCs':
      font.strike = true;
      break;
    case 'sz': {
      const size = Number(stringAttribute(tag, 'val'));
      if (Number.isFinite(size) && size > 0) font.size = size;
      break;
    }
    case 'name': {
      const name = stringAttribute(tag, 'val');
      if (name !== null && name !== '') font.name = name;
      break;
    }
    case 'color': {
      const color = resolveColor({
        rgb: stringAttribute(tag, 'rgb'),
        indexed: stringAttribute(tag, 'indexed'),
        theme: stringAttribute(tag, 'theme'),
        tint: stringAttribute(tag, 'tint'),
        auto: stringAttribute(tag, 'auto'),
      });
      if (color !== null) font.color = color;
      break;
    }
    default:
      break;
  }
}

function applyFillChild(
  tag: { local: string; attributes: Record<string, { local: string; value: string }> },
  fill: FillAccumulator,
): void {
  switch (tag.local) {
    case 'patternFill': {
      const pattern = stringAttribute(tag, 'patternType');
      if (pattern !== null) fill.pattern = pattern;
      break;
    }
    case 'fgColor': {
      const color = resolveColor({
        rgb: stringAttribute(tag, 'rgb'),
        indexed: stringAttribute(tag, 'indexed'),
        theme: stringAttribute(tag, 'theme'),
        tint: stringAttribute(tag, 'tint'),
        auto: stringAttribute(tag, 'auto'),
      });
      if (color !== null) fill.color = color;
      break;
    }
    default:
      break;
  }
}

function applyBorderChild(
  tag: { local: string; attributes: Record<string, { local: string; value: string }> },
  border: BorderAccumulator,
  currentEdge: BorderEdgeKey | null,
): BorderEdgeKey | null {
  switch (tag.local) {
    case 'left':
    case 'right':
    case 'top':
    case 'bottom':
    case 'diagonal': {
      const edge: BorderEdgeStyle = {
        style: stringAttribute(tag, 'style') ?? 'none',
        color: null,
      };
      border[tag.local] = edge;
      return tag.local;
    }
    case 'diagonalUp':
      border.diagonalUp = true;
      return currentEdge;
    case 'diagonalDown':
      border.diagonalDown = true;
      return currentEdge;
    case 'color':
      if (currentEdge !== null) {
        const edge = border[currentEdge];
        border[currentEdge] = {
          ...edge,
          color: resolveColor({
            rgb: stringAttribute(tag, 'rgb'),
            indexed: stringAttribute(tag, 'indexed'),
            theme: stringAttribute(tag, 'theme'),
            tint: stringAttribute(tag, 'tint'),
            auto: stringAttribute(tag, 'auto'),
          }),
        };
      }
      return currentEdge;
    default:
      return currentEdge;
  }
}

function resolveBorder(accumulator: BorderAccumulator | undefined): BorderStyle {
  return accumulator ?? DEFAULT_BORDER;
}

function resolveFont(accumulator: FontAccumulator | undefined): FontStyle {
  if (accumulator === undefined) return DEFAULT_FONT;
  return {
    name: accumulator.name ?? DEFAULT_FONT.name,
    size: accumulator.size ?? DEFAULT_FONT.size,
    bold: accumulator.bold,
    italic: accumulator.italic,
    underline: accumulator.underline,
    strike: accumulator.strike,
    color: accumulator.color,
  };
}

function resolveFill(accumulator: FillAccumulator | undefined): FillStyle {
  if (accumulator === undefined) return DEFAULT_FILL;
  const pattern = accumulator.pattern ?? 'none';
  if (pattern === 'none' || pattern === 'gray125') {
    return { pattern, color: null };
  }
  return { pattern, color: accumulator.color };
}

function resolveNumberFormat(
  numFmtId: number,
  customFormats: ReadonlyMap<number, string>,
): string | null {
  const custom = customFormats.get(numFmtId);
  if (custom !== undefined) return custom;
  const builtin = builtinNumberFormat(numFmtId);
  if (builtin === undefined || builtin === 'General') return null;
  return builtin;
}

function readAlignment(tag: {
  attributes: Record<string, { local: string; value: string }>;
}): AlignmentStyle {
  const indent = Number(stringAttribute(tag, 'indent') ?? '0');
  const rotation = Number(stringAttribute(tag, 'textRotation') ?? '0');
  return {
    horizontal: normalizeHorizontalAlignment(stringAttribute(tag, 'horizontal')),
    vertical: normalizeVerticalAlignment(stringAttribute(tag, 'vertical')),
    wrapText: stringAttribute(tag, 'wrapText') === '1',
    indent: Number.isFinite(indent) ? Math.max(0, indent) : 0,
    textRotation: Number.isFinite(rotation) ? rotation : 0,
  };
}

function numberAttribute(
  tag: { attributes: Record<string, { local: string; value: string }> },
  local: string,
  fallback: number,
): number {
  const value = stringAttribute(tag, local);
  if (value === null) return fallback;
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function stringAttribute(
  tag: { attributes: Record<string, { local: string; value: string }> },
  local: string,
): string | null {
  for (const attribute of Object.values(tag.attributes)) {
    if (attribute.local === local) return attribute.value;
  }
  return null;
}

function isTruthy(value: string | null): boolean {
  return value === '1' || value === 'true' || value === 'on';
}
