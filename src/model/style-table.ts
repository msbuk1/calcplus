import type {
  AlignmentStyle,
  BorderStyle,
  FillStyle,
  FontStyle,
  NumberFormatKind,
  StyleRecord,
  Workbook,
} from './types';
import { borderWeight, builtinNumberFormat, builtinNumberFormatId } from './style';
import { appendStyle, type StyleRecordInput } from './workbook';

/**
 * The shared, append-only style table (workbook-model §2).
 *
 * A formatting command creates or reuses a style record. Existing ids are
 * never renumbered, so `.xlsx` style references survive a round-trip. A record
 * that cannot be found is appended with the next contiguous id.
 */

export interface StylePatch {
  readonly font?: Partial<FontStyle>;
  readonly fill?: Partial<FillStyle>;
  readonly border?: Partial<BorderStyle>;
  readonly alignment?: Partial<AlignmentStyle>;
  readonly numFmtCode?: string | null;
  readonly numFmtId?: number;
}

/** Build the next style record from a base record and a formatting patch. */
export function patchedStyleInput(base: StyleRecord, patch: StylePatch): StyleRecordInput {
  const numberFormat = patchedNumberFormat(base, patch);
  return {
    numFmtId: numberFormat.numFmtId,
    // The source sub-record ids stay as they are until the styles writer can
    // append font/fill/border records. Display and undo read the resolved
    // parts below, which are authoritative for a new record.
    fontId: base.fontId,
    fillId: base.fillId,
    borderId: base.borderId,
    xfId: base.xfId,
    rawXml: '',
    font: patchedFont(base.font, patch.font),
    fill: patchedFill(base.fill, patch.fill),
    border: mergeBorder(base.border, patch.border),
    alignment: { ...base.alignment, ...patch.alignment },
    numFmtCode: numberFormat.numFmtCode,
    numFmtKind: numberFormat.numFmtKind,
  };
}

function patchedFont(base: FontStyle, patch: Partial<FontStyle> | undefined): FontStyle {
  const font = { ...base, ...patch };
  if (patch?.color !== undefined) {
    font.color = patch.color === null ? null : patch.color.toUpperCase();
  }
  if (patch?.underlineStyle !== undefined) {
    return { ...font, underline: patch.underlineStyle !== 'none' };
  }
  if (patch?.underline !== undefined) {
    return {
      ...font,
      underlineStyle: patch.underline ? (base.underline ? base.underlineStyle : 'single') : 'none',
    };
  }
  return font;
}

function patchedFill(base: FillStyle, patch: Partial<FillStyle> | undefined): FillStyle {
  const fill = { ...base, ...patch };
  if (patch?.color !== undefined) {
    fill.color = patch.color === null ? null : patch.color.toUpperCase();
  }
  if (patch?.backgroundColor !== undefined) {
    fill.backgroundColor =
      patch.backgroundColor === null ? null : patch.backgroundColor.toUpperCase();
  }
  return fill;
}

function patchedNumberFormat(
  base: StyleRecord,
  patch: StylePatch,
): {
  readonly numFmtId: number;
  readonly numFmtCode: string | null;
  readonly numFmtKind: NumberFormatKind;
} {
  if (patch.numFmtCode === undefined) {
    if (patch.numFmtId === undefined) {
      return {
        numFmtId: base.numFmtId,
        numFmtCode: base.numFmtCode,
        numFmtKind: base.numFmtKind,
      };
    }
    if (patch.numFmtId === 0) {
      return { numFmtId: 0, numFmtCode: null, numFmtKind: 'general' };
    }
    const builtin = builtinNumberFormat(patch.numFmtId);
    if (builtin !== null) {
      return {
        numFmtId: patch.numFmtId,
        numFmtCode: builtin === 'General' ? null : builtin,
        numFmtKind: 'builtin',
      };
    }
    return {
      numFmtId: patch.numFmtId,
      numFmtCode: null,
      numFmtKind: 'unknown',
    };
  }

  const requested = patch.numFmtCode;
  if (requested === null || requested.trim() === '' || requested.trim() === 'General') {
    return { numFmtId: 0, numFmtCode: null, numFmtKind: 'general' };
  }
  const builtinId = builtinNumberFormatId(requested);
  if (builtinId !== null) {
    const builtinCode = builtinNumberFormat(builtinId);
    return {
      numFmtId: builtinId,
      numFmtCode: builtinCode === 'General' ? null : (builtinCode ?? requested),
      numFmtKind: 'builtin',
    };
  }
  // A custom id is allocated by the styles writer at save time. Until then the
  // record is identified by its authored code, which is what round-trips.
  return { numFmtId: 0, numFmtCode: requested, numFmtKind: 'custom' };
}

/**
 * Find an identical style or append a new one. Identity covers the visual
 * parts and the number-format reference; the source `rawXml` and sub-record
 * ids are transport details, not appearance.
 */
export function findOrAppendStyle(workbook: Workbook, input: StyleRecordInput): number {
  const signature = styleSignature(input);
  for (const existing of workbook.styles) {
    if (styleSignature(existing) === signature) return existing.id;
  }
  return appendStyle(workbook, input);
}

export function styleSignature(style: StyleRecordInput | StyleRecord): string {
  const font = style.font;
  const fill = style.fill;
  const border = style.border;
  const alignment = style.alignment;
  return JSON.stringify([
    numberFormatSignature(style),
    font === undefined
      ? null
      : [
          font.name,
          font.size,
          font.bold,
          font.italic,
          font.underline,
          font.underlineStyle ?? (font.underline ? 'single' : 'none'),
          font.strike,
          font.color === null ? null : font.color.toUpperCase(),
        ],
    fill === undefined
      ? null
      : [
          fill.pattern,
          fill.color === null ? null : fill.color.toUpperCase(),
          fill.backgroundColor === null || fill.backgroundColor === undefined
            ? null
            : fill.backgroundColor.toUpperCase(),
        ],
    border === undefined
      ? null
      : [
          edgeSignature(border.left),
          edgeSignature(border.right),
          edgeSignature(border.top),
          edgeSignature(border.bottom),
          edgeSignature(border.diagonal),
          border.diagonalUp,
          border.diagonalDown,
        ],
    alignment === undefined
      ? null
      : [
          alignment.horizontal,
          alignment.vertical,
          alignment.wrapText,
          alignment.indent,
          alignment.textRotation,
          alignment.shrinkToFit ?? false,
        ],
  ]);
}

function numberFormatSignature(style: StyleRecordInput | StyleRecord): readonly unknown[] {
  const kind = style.numFmtKind ?? (style.numFmtCode === null ? 'general' : 'custom');
  switch (kind) {
    case 'general':
      return ['general'];
    case 'builtin':
      return ['builtin', style.numFmtId];
    case 'custom':
      return ['custom', style.numFmtCode ?? ''];
    case 'unknown':
      return ['unknown', style.numFmtId];
  }
}

function edgeSignature(edge: {
  readonly style: string;
  readonly color: string | null;
  readonly weight?: string | null;
}): string {
  const color = edge.color === null ? '' : edge.color.toUpperCase();
  return `${edge.style}\u0000${color}\u0000${edge.weight ?? ''}`;
}

function mergeBorder(base: BorderStyle, patch: Partial<BorderStyle> | undefined): BorderStyle {
  if (patch === undefined) return base;
  return {
    left: normalizedEdge(patch.left ?? base.left),
    right: normalizedEdge(patch.right ?? base.right),
    top: normalizedEdge(patch.top ?? base.top),
    bottom: normalizedEdge(patch.bottom ?? base.bottom),
    diagonal: normalizedEdge(patch.diagonal ?? base.diagonal),
    diagonalUp: patch.diagonalUp ?? base.diagonalUp,
    diagonalDown: patch.diagonalDown ?? base.diagonalDown,
  };
}

function normalizedEdge(edge: BorderStyle['left']): BorderStyle['left'] {
  const normalized = edge.color === null ? edge : { ...edge, color: edge.color.toUpperCase() };
  return normalized.weight === undefined || normalized.weight === null
    ? { ...normalized, weight: borderWeight(normalized.style) }
    : normalized;
}
