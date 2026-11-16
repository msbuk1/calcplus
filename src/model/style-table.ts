import type {
  AlignmentStyle,
  BorderStyle,
  FillStyle,
  FontStyle,
  StyleRecord,
  Workbook,
} from './types';
import { builtinNumberFormatId } from './style';
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
  const numFmtCode = patch.numFmtCode === undefined ? base.numFmtCode : patch.numFmtCode;
  const numFmtId =
    patch.numFmtId !== undefined
      ? patch.numFmtId
      : patch.numFmtCode !== undefined
        ? (builtinNumberFormatId(numFmtCode) ?? 0)
        : base.numFmtId;
  return {
    numFmtId,
    // The source sub-record ids stay as they are until the styles writer can
    // append font/fill/border records. Display and undo read the resolved
    // parts below, which are authoritative for a new record.
    fontId: base.fontId,
    fillId: base.fillId,
    borderId: base.borderId,
    xfId: base.xfId,
    rawXml: '',
    font: { ...base.font, ...patch.font },
    fill: { ...base.fill, ...patch.fill },
    border: mergeBorder(base.border, patch.border),
    alignment: { ...base.alignment, ...patch.alignment },
    numFmtCode,
  };
}

/**
 * Find an identical style or append a new one. Identity covers the visual
 * parts and the number-format reference; the source `rawXml` and `xfId` are
 * transport details, not appearance.
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
    // The authored code is what the user sees. Two storage ids with the same
    // code are interchangeable for display and for an append-only round-trip.
    style.numFmtCode ?? `id:${style.numFmtId}`,
    font === undefined
      ? null
      : [font.name, font.size, font.bold, font.italic, font.underline, font.strike, font.color],
    fill === undefined ? null : [fill.pattern, fill.color],
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
        ],
  ]);
}

function edgeSignature(edge: { readonly style: string; readonly color: string | null }): string {
  return `${edge.style}\u0000${edge.color ?? ''}`;
}

function mergeBorder(base: BorderStyle, patch: Partial<BorderStyle> | undefined): BorderStyle {
  if (patch === undefined) return base;
  return {
    left: patch.left ?? base.left,
    right: patch.right ?? base.right,
    top: patch.top ?? base.top,
    bottom: patch.bottom ?? base.bottom,
    diagonal: patch.diagonal ?? base.diagonal,
    diagonalUp: patch.diagonalUp ?? base.diagonalUp,
    diagonalDown: patch.diagonalDown ?? base.diagonalDown,
  };
}
