import type {
  AlignmentStyle,
  BorderEdgeStyle,
  BorderStyle,
  FillStyle,
  FontStyle,
  NamedStyleRecord,
  StyleRecord,
  UnderlineStyle,
} from '@model/types';
import { createDefaultStyle } from '@model/workbook';
import {
  DEFAULT_ALIGNMENT,
  DEFAULT_BORDER,
  DEFAULT_FILL,
  DEFAULT_FONT,
  borderWeight,
  builtinNumberFormat,
  normalizeHorizontalAlignment,
  normalizeVerticalAlignment,
  resolveColor,
  underlineFromValue,
  type ColorAttributes,
} from '@model/style';
import { decodeXml, parseXml, openTagStart } from './xml';

export interface StylesXmlInfo {
  readonly styles: StyleRecord[];
  readonly namedStyles: NamedStyleRecord[];
  /** The number of `<xf>` entries actually present in `<cellXfs>`. */
  readonly sourceStyleCount: number;
  /** The source font table, by `fontId`. */
  readonly fonts: readonly FontStyle[];
  /** The source fill table, by `fillId`. */
  readonly fills: readonly FillStyle[];
  /** The source border table, by `borderId`. */
  readonly borders: readonly BorderStyle[];
  /** Authored `<numFmt>` entries, by id. */
  readonly customFormats: ReadonlyMap<number, string>;
  /** The workbook's `<colors><indexedColors>` palette, when declared. */
  readonly indexedPalette: readonly string[] | null;
}

interface OpenElement {
  readonly local: string;
  readonly start: number;
}

interface RawFont {
  name: string | null;
  size: number | null;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  underlineStyle: UnderlineStyle;
  strike: boolean;
  color: ColorAttributes | null;
}

interface RawFill {
  pattern: string | null;
  foreground: ColorAttributes | null;
  background: ColorAttributes | null;
}

type BorderEdgeKey = 'left' | 'right' | 'top' | 'bottom' | 'diagonal';

interface RawBorderEdge {
  style: string;
  color: ColorAttributes | null;
}

interface RawBorder {
  left: RawBorderEdge;
  right: RawBorderEdge;
  top: RawBorderEdge;
  bottom: RawBorderEdge;
  diagonal: RawBorderEdge;
  diagonalUp: boolean;
  diagonalDown: boolean;
}

const NO_RAW_EDGE: () => RawBorderEdge = () => ({ style: 'none', color: null });

/**
 * Parse the workbook style sheet into the shared, append-only style table:
 * font, fill, border, alignment and the number-format reference (built-in id
 * plus authored code) for every cell style.
 */
export function parseStylesXml(bytes: Uint8Array, partName: string): StylesXmlInfo {
  const xml = decodeXml(bytes);
  const indexedPalette = parseIndexedPalette(xml, partName);
  const styles: StyleRecord[] = [];
  const namedStyles: NamedStyleRecord[] = [];
  const fonts: RawFont[] = [];
  const fills: RawFill[] = [];
  const borders: RawBorder[] = [];
  const customFormats = new Map<number, string>();
  const stack: OpenElement[] = [];
  let inCellXfs = false;
  let inCellStyleXfs = false;
  let cellXfsDepth = -1;
  let cellStyleXfsDepth = -1;
  let inFonts = false;
  let inFills = false;
  let inBorders = false;
  let font: RawFont | null = null;
  let fill: RawFill | null = null;
  let border: RawBorder | null = null;
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
        } else if (tag.local === 'numFmt' && parent === 'numFmts') {
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
            underlineStyle: 'none',
            strike: false,
            color: null,
          };
        } else if (font !== null) {
          applyFontChild(tag, font);
        } else if (inFills && tag.local === 'fill' && parent === 'fills') {
          fill = { pattern: null, foreground: null, background: null };
        } else if (fill !== null) {
          applyFillChild(tag, fill);
        } else if (inBorders && tag.local === 'border' && parent === 'borders') {
          border = {
            left: NO_RAW_EDGE(),
            right: NO_RAW_EDGE(),
            top: NO_RAW_EDGE(),
            bottom: NO_RAW_EDGE(),
            diagonal: NO_RAW_EDGE(),
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
            numFmtKind: 'general',
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

  if (styles.length === 0) {
    const palette = indexedPalette;
    return {
      styles: [createDefaultStyle(0)],
      namedStyles,
      sourceStyleCount: 0,
      fonts:
        fonts.length === 0
          ? [resolveFont(undefined, palette)]
          : fonts.map((value) => resolveFont(value, palette)),
      fills:
        fills.length === 0
          ? [resolveFill(undefined, palette)]
          : fills.map((value) => resolveFill(value, palette)),
      borders:
        borders.length === 0
          ? [resolveBorder(undefined, palette)]
          : borders.map((value) => resolveBorder(value, palette)),
      customFormats,
      indexedPalette,
    };
  }
  return {
    styles: styles.map((style, index) => {
      const numberFormat = resolveNumberFormat(style.numFmtId, customFormats);
      return {
        ...style,
        id: index,
        font: resolveFont(fonts[style.fontId], indexedPalette),
        fill: resolveFill(fills[style.fillId], indexedPalette),
        border: resolveBorder(borders[style.borderId], indexedPalette),
        numFmtCode: numberFormat.numFmtCode,
        numFmtKind: numberFormat.numFmtKind,
      };
    }),
    namedStyles,
    sourceStyleCount: styles.length,
    fonts: fonts.map((value) => resolveFont(value, indexedPalette)),
    fills: fills.map((value) => resolveFill(value, indexedPalette)),
    borders: borders.map((value) => resolveBorder(value, indexedPalette)),
    customFormats,
    indexedPalette,
  };
}

/** Parse the effective legacy indexed palette from `<colors><indexedColors>`. */
function parseIndexedPalette(xml: string, partName: string): readonly string[] | null {
  const palette: string[] = [];
  let inIndexedColors = false;
  parseXml(
    xml,
    {
      onOpenTag(tag) {
        if (tag.local === 'indexedColors') {
          inIndexedColors = true;
          return;
        }
        if (!inIndexedColors) return;
        if (tag.local === 'rgbColor') {
          const rgb = stringAttribute(tag, 'rgb');
          const hex = rgb === null ? null : normalizeRgb(rgb);
          if (hex !== null) palette.push(hex);
        }
      },
      onCloseTag(tag) {
        if (tag.local === 'indexedColors') inIndexedColors = false;
      },
    },
    partName,
  );
  return palette.length === 0 ? null : palette;
}

function normalizeRgb(value: string): string | null {
  const trimmed = value.trim();
  if (/^[0-9a-f]{8}$/i.test(trimmed)) return trimmed.slice(2).toUpperCase();
  if (/^[0-9a-f]{6}$/i.test(trimmed)) return trimmed.toUpperCase();
  return null;
}

function applyFontChild(
  tag: { local: string; attributes: Record<string, { local: string; value: string }> },
  font: RawFont,
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
    case 'u': {
      const underline = underlineFromValue(stringAttribute(tag, 'val'));
      font.underline = underline.underline;
      font.underlineStyle = underline.underlineStyle;
      break;
    }
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
    case 'color':
      font.color = readColorAttributes(tag);
      break;
    default:
      break;
  }
}

function applyFillChild(
  tag: { local: string; attributes: Record<string, { local: string; value: string }> },
  fill: RawFill,
): void {
  switch (tag.local) {
    case 'patternFill': {
      const pattern = stringAttribute(tag, 'patternType');
      if (pattern !== null) fill.pattern = pattern;
      break;
    }
    case 'fgColor':
      fill.foreground = readColorAttributes(tag);
      break;
    case 'bgColor':
      fill.background = readColorAttributes(tag);
      break;
    default:
      break;
  }
}

function applyBorderChild(
  tag: { local: string; attributes: Record<string, { local: string; value: string }> },
  border: RawBorder,
  currentEdge: BorderEdgeKey | null,
): BorderEdgeKey | null {
  switch (tag.local) {
    case 'left':
    case 'right':
    case 'top':
    case 'bottom':
    case 'diagonal': {
      border[tag.local] = {
        style: stringAttribute(tag, 'style') ?? 'none',
        color: null,
      };
      return tag.local;
    }
    case 'color':
      if (currentEdge !== null) {
        border[currentEdge] = {
          ...border[currentEdge],
          color: readColorAttributes(tag),
        };
      }
      return currentEdge;
    default:
      return currentEdge;
  }
}

function resolveBorder(
  accumulator: RawBorder | undefined,
  palette: readonly string[] | null,
): BorderStyle {
  if (accumulator === undefined) return DEFAULT_BORDER;
  return {
    left: resolveBorderEdge(accumulator.left, palette),
    right: resolveBorderEdge(accumulator.right, palette),
    top: resolveBorderEdge(accumulator.top, palette),
    bottom: resolveBorderEdge(accumulator.bottom, palette),
    diagonal: resolveBorderEdge(accumulator.diagonal, palette),
    diagonalUp: accumulator.diagonalUp,
    diagonalDown: accumulator.diagonalDown,
  };
}

function resolveBorderEdge(
  edge: RawBorderEdge,
  palette: readonly string[] | null,
): BorderEdgeStyle {
  return {
    style: edge.style,
    color: edge.color === null ? null : resolveColor(edge.color, palette),
    weight: borderWeight(edge.style),
  };
}

function resolveFont(
  accumulator: RawFont | undefined,
  palette: readonly string[] | null,
): FontStyle {
  if (accumulator === undefined) return DEFAULT_FONT;
  return {
    name: accumulator.name ?? DEFAULT_FONT.name,
    size: accumulator.size ?? DEFAULT_FONT.size,
    bold: accumulator.bold,
    italic: accumulator.italic,
    underline: accumulator.underline,
    underlineStyle: accumulator.underlineStyle,
    strike: accumulator.strike,
    color: accumulator.color === null ? null : resolveColor(accumulator.color, palette),
  };
}

function resolveFill(
  accumulator: RawFill | undefined,
  palette: readonly string[] | null,
): FillStyle {
  if (accumulator === undefined) return DEFAULT_FILL;
  const pattern = accumulator.pattern ?? 'none';
  if (pattern === 'none') {
    return { pattern, color: null, backgroundColor: null };
  }
  return {
    pattern,
    color: accumulator.foreground === null ? null : resolveColor(accumulator.foreground, palette),
    backgroundColor:
      accumulator.background === null ? null : resolveColor(accumulator.background, palette),
  };
}

function resolveNumberFormat(
  numFmtId: number,
  customFormats: ReadonlyMap<number, string>,
): Pick<StyleRecord, 'numFmtCode' | 'numFmtKind'> {
  const custom = customFormats.get(numFmtId);
  if (custom !== undefined) return { numFmtCode: custom, numFmtKind: 'custom' };
  const builtin = builtinNumberFormat(numFmtId);
  if (builtin === null) return { numFmtCode: null, numFmtKind: 'unknown' };
  if (builtin === 'General') return { numFmtCode: null, numFmtKind: 'general' };
  return { numFmtCode: builtin, numFmtKind: 'builtin' };
}

function readAlignment(tag: {
  attributes: Record<string, { local: string; value: string }>;
}): AlignmentStyle {
  const indent = Number(stringAttribute(tag, 'indent') ?? '0');
  const rotation = Number(stringAttribute(tag, 'textRotation') ?? '0');
  return {
    horizontal: normalizeHorizontalAlignment(stringAttribute(tag, 'horizontal')),
    vertical: normalizeVerticalAlignment(stringAttribute(tag, 'vertical')),
    wrapText: isTruthy(stringAttribute(tag, 'wrapText')),
    indent: Number.isFinite(indent) ? Math.max(0, indent) : 0,
    textRotation: Number.isFinite(rotation) ? rotation : 0,
    shrinkToFit: isTruthy(stringAttribute(tag, 'shrinkToFit')),
  };
}

function readColorAttributes(tag: {
  attributes: Record<string, { local: string; value: string }>;
}): ColorAttributes {
  return {
    rgb: stringAttribute(tag, 'rgb'),
    indexed: stringAttribute(tag, 'indexed'),
    theme: stringAttribute(tag, 'theme'),
    tint: stringAttribute(tag, 'tint'),
    auto: stringAttribute(tag, 'auto'),
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
