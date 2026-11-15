import type { CellValue } from './value';
import { cloneCellValue } from './value';

/** Stable identifiers. The string value is opaque to consumers. */
export type WorkbookId = string;
export type SheetId = string;
export type StyleId = number;

export interface CellAddress {
  readonly row: number;
  readonly col: number;
}

export interface GridRange {
  readonly startRow: number;
  readonly startCol: number;
  readonly endRow: number;
  readonly endCol: number;
}

export interface RichTextPayload {
  /** The original `<si>` XML. Opaque to the model; displayed as plain text. */
  readonly xml: string;
  readonly plainText: string;
}

export type FormulaStorageKind = 'shared-master' | 'shared-follower' | 'array';

/**
 * Storage metadata for formulas that are not plain authored formulas.
 * File-compat §6.1: shared groups and legacy array formulas are preserved
 * verbatim and are read-only in M1.
 */
export interface FormulaSource {
  readonly kind: FormulaStorageKind;
  /** `si` attribute for a shared formula. */
  readonly sharedIndex?: string;
  /** `ref` attribute for a shared master or a legacy array formula. */
  readonly range?: string;
  /** Exact original `<f>` element. Never rewritten when the cell is unchanged. */
  readonly xml: string;
}

export interface Cell {
  value: CellValue;
  formula: string | null;
  formulaDerived: boolean;
  savedResult: CellValue | null;
  styleId: StyleId;
  richText: RichTextPayload | null;
  /** Present only for shared groups and legacy array formulas. */
  readonly formulaSource?: FormulaSource;
}

export interface DefinedName {
  readonly name: string;
  readonly formula: string;
  /** `null` for workbook scope, otherwise the owning sheet. */
  readonly sheetId: SheetId | null;
}

export interface RowDimension {
  height?: number;
  hidden?: boolean;
  styleId?: StyleId;
}

export interface ColumnDimension {
  width?: number;
  hidden?: boolean;
  styleId?: StyleId;
}

export interface MergedRange {
  readonly range: GridRange;
}

export interface CommentRecord {
  readonly text: string;
  readonly author?: string;
  readonly threaded: boolean;
}

export interface HyperlinkRecord {
  readonly target: string;
  readonly location?: string;
  readonly tooltip?: string;
}

export interface ValidationRecord {
  readonly range: GridRange;
  readonly rawXml: string;
}

export interface ConditionalFormatRecord {
  readonly range: GridRange;
  readonly rawXml: string;
}

export interface SheetView {
  freezeRows: number;
  freezeCols: number;
  showGridLines: boolean;
  showHeadings: boolean;
  zoom: number;
}

export interface OpaquePart {
  readonly path: string;
  readonly bytes: Uint8Array;
  /** Position on a sheet, once M5/M6 range transforms land. */
  readonly anchor: null;
}

export interface Sheet {
  readonly id: SheetId;
  name: string;
  /** Part path inside the package, used only by the OOXML layer. */
  readonly partPath: string;
  /** Sparse map keyed by `row:col`, per workbook-model §1. */
  readonly cells: Map<string, Cell>;
  readonly rowDimensions: Map<number, RowDimension>;
  readonly columnDimensions: Map<number, ColumnDimension>;
  readonly merges: MergedRange[];
  readonly names: DefinedName[];
  readonly comments: Map<string, CommentRecord>;
  readonly hyperlinks: Map<string, HyperlinkRecord>;
  readonly validation: ValidationRecord[];
  readonly conditionalFormats: ConditionalFormatRecord[];
  protection: { readonly rawXml: string } | null;
  readonly view: SheetView;
  readonly opaqueParts: Map<string, OpaquePart>;
}

/** A font, resolved for display. Workbook-model §2: the style table holds the font. */
export interface FontStyle {
  readonly name: string;
  /** Size in points. */
  readonly size: number;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly underline: boolean;
  readonly strike: boolean;
  /** CSS colour, or `null` for the automatic colour. */
  readonly color: string | null;
}

/** A fill, resolved for display. Only solid fills are painted in M2. */
export interface FillStyle {
  /** The OOXML pattern type, for example `none`, `gray125` or `solid`. */
  readonly pattern: string;
  /** CSS colour of the foreground, or `null` when the fill paints nothing. */
  readonly color: string | null;
}

export type HorizontalAlignment =
  'general' | 'left' | 'center' | 'right' | 'fill' | 'justify' | 'centerContinuous' | 'distributed';

export type VerticalAlignment = 'top' | 'center' | 'bottom' | 'justify' | 'distributed';

/** Cell alignment, resolved for display. */
export interface AlignmentStyle {
  readonly horizontal: HorizontalAlignment;
  readonly vertical: VerticalAlignment;
  readonly wrapText: boolean;
  readonly indent: number;
  /** Degrees, clockwise, 0-180. */
  readonly textRotation: number;
}

export interface StyleRecord {
  readonly id: StyleId;
  readonly numFmtId: number;
  readonly fontId: number;
  readonly fillId: number;
  readonly borderId: number;
  readonly xfId: number;
  /** Original `<xf>` XML. The append-only table never renumbers existing ids. */
  readonly rawXml: string;
  /** The font, fill and alignment this style points at, resolved for display. */
  readonly font: FontStyle;
  readonly fill: FillStyle;
  readonly alignment: AlignmentStyle;
  /** The authored number-format code, or `null` when the style uses the general format. */
  readonly numFmtCode: string | null;
}

export interface NamedStyleRecord {
  readonly name: string;
  readonly xfId: number;
  readonly builtinId: number | null;
}

export interface ThemeRecord {
  /** The original theme part, or `null` when the workbook has no theme. */
  readonly path: string | null;
  readonly bytes: Uint8Array | null;
}

export interface WorkbookProtectionRecord {
  readonly rawXml: string;
}

export interface Workbook {
  id: WorkbookId;
  /** Tab order. */
  readonly sheets: Sheet[];
  readonly names: DefinedName[];
  /** Append-only style table. Numeric style ids match OOXML `s` indexes. */
  readonly styles: StyleRecord[];
  readonly theme: ThemeRecord;
  readonly namedStyles: NamedStyleRecord[];
  readonly protection: WorkbookProtectionRecord | null;
  readonly opaqueParts: Map<string, OpaquePart>;
  /** Original bytes of every part the model reads (workbook, styles, shared strings, worksheets). */
  readonly sourceParts: Map<string, Uint8Array>;
  date1904: boolean;
}

export function cellKey(row: number, col: number): string {
  return `${row}:${col}`;
}

export function parseCellKey(key: string): CellAddress {
  const separator = key.indexOf(':');
  if (separator < 0) throw new Error(`Invalid cell key: ${key}`);
  const row = Number(key.slice(0, separator));
  const col = Number(key.slice(separator + 1));
  if (!Number.isInteger(row) || !Number.isInteger(col) || row < 0 || col < 0) {
    throw new Error(`Invalid cell key: ${key}`);
  }
  return { row, col };
}

export function createCell(value: CellValue, styleId = 0): Cell {
  return {
    value,
    formula: null,
    formulaDerived: false,
    savedResult: null,
    styleId,
    richText: null,
  };
}

export function cloneCell(cell: Cell | null): Cell | null {
  if (cell === null) return null;
  return {
    value: cloneCellValue(cell.value),
    formula: cell.formula,
    formulaDerived: cell.formulaDerived,
    savedResult: cell.savedResult === null ? null : cloneCellValue(cell.savedResult),
    styleId: cell.styleId,
    richText: cell.richText === null ? null : { ...cell.richText },
    ...(cell.formulaSource === undefined ? {} : { formulaSource: { ...cell.formulaSource } }),
  };
}

export function formulaSourceEquals(
  left: FormulaSource | undefined,
  right: FormulaSource | undefined,
): boolean {
  if (left === undefined || right === undefined) return left === right;
  return (
    left.kind === right.kind &&
    left.sharedIndex === right.sharedIndex &&
    left.range === right.range &&
    left.xml === right.xml
  );
}

export function cellStateEquals(left: Cell | null, right: Cell | null): boolean {
  if (left === null || right === null) return left === right;
  return (
    cellValueEqualsForModel(left.value, right.value) &&
    left.formula === right.formula &&
    left.formulaDerived === right.formulaDerived &&
    nullableCellValueEquals(left.savedResult, right.savedResult) &&
    left.styleId === right.styleId &&
    richTextEquals(left.richText, right.richText) &&
    formulaSourceEquals(left.formulaSource, right.formulaSource)
  );
}

/**
 * The formula identity used to decide whether a save changes formula storage.
 * A shared follower counts as a formula even though its authored text is empty.
 */
export function cellFormulaKey(cell: Cell | null): string | null {
  if (cell === null) return null;
  if (cell.formula !== null) {
    return `${cell.formulaSource?.kind ?? 'normal'}\u0000${cell.formula}`;
  }
  if (cell.formulaSource?.kind === 'shared-follower') return 'shared-follower';
  return null;
}

function cellValueEqualsForModel(left: CellValue, right: CellValue): boolean {
  if (left.type !== right.type) return false;
  switch (left.type) {
    case 'number':
      return right.type === 'number' && Object.is(left.value, right.value);
    case 'text':
      return right.type === 'text' && left.value === right.value;
    case 'boolean':
      return right.type === 'boolean' && left.value === right.value;
    case 'error':
      return right.type === 'error' && left.value === right.value;
    case 'blank':
      return true;
  }
}

function nullableCellValueEquals(left: CellValue | null, right: CellValue | null): boolean {
  if (left === null || right === null) return left === right;
  return cellValueEqualsForModel(left, right);
}

function richTextEquals(left: Cell['richText'], right: Cell['richText']): boolean {
  if (left === null || right === null) return left === right;
  return left.xml === right.xml && left.plainText === right.plainText;
}

export function cellIsEmpty(cell: Cell): boolean {
  return (
    cell.value.type === 'blank' &&
    cell.formula === null &&
    cell.savedResult === null &&
    cell.richText === null &&
    cell.formulaSource === undefined
  );
}

export function getCell(sheet: Sheet, row: number, col: number): Cell | null {
  return sheet.cells.get(cellKey(row, col)) ?? null;
}

/** Model mutation for the command/import layers only. Surfaces dispatch commands. */
export function setCellSnapshot(sheet: Sheet, row: number, col: number, cell: Cell | null): void {
  const key = cellKey(row, col);
  if (cell === null) {
    sheet.cells.delete(key);
    return;
  }
  sheet.cells.set(key, cell);
}

export function getSheetById(workbook: Workbook, sheetId: SheetId): Sheet {
  const sheet = workbook.sheets.find((candidate) => candidate.id === sheetId);
  if (sheet === undefined) throw new Error(`Unknown sheet: ${sheetId}`);
  return sheet;
}

export function columnToLetters(column: number): string {
  if (!Number.isInteger(column) || column < 0) throw new RangeError('Column index must be >= 0');
  let value = column;
  let result = '';
  while (value >= 0) {
    const digit = value % 26;
    result = String.fromCharCode(65 + digit) + result;
    value = Math.floor(value / 26) - 1;
  }
  return result;
}

export function lettersToColumn(letters: string): number {
  if (!/^[A-Z]+$/i.test(letters)) throw new Error(`Invalid column letters: ${letters}`);
  let value = 0;
  for (const character of letters.toUpperCase()) {
    value = value * 26 + (character.charCodeAt(0) - 64);
  }
  return value - 1;
}

export function formatCellRef(row: number, col: number): string {
  return `${columnToLetters(col)}${row + 1}`;
}

export function parseCellRef(ref: string): CellAddress {
  const match = /^([A-Z]+)(\d+)$/i.exec(ref.trim());
  if (match === null) throw new Error(`Invalid cell reference: ${ref}`);
  const col = lettersToColumn(match[1]!);
  const row = Number(match[2]) - 1;
  if (!Number.isSafeInteger(row) || row < 0)
    throw new Error(`Invalid row in cell reference: ${ref}`);
  return { row, col };
}

export function parseRangeRef(ref: string): GridRange {
  const [startText, endText = startText] = ref.split(':');
  const start = parseCellRef(startText!);
  const end = parseCellRef(endText!);
  return {
    startRow: Math.min(start.row, end.row),
    startCol: Math.min(start.col, end.col),
    endRow: Math.max(start.row, end.row),
    endCol: Math.max(start.col, end.col),
  };
}

export function rangeContains(range: GridRange, row: number, col: number): boolean {
  return (
    row >= range.startRow && row <= range.endRow && col >= range.startCol && col <= range.endCol
  );
}

export function singleCellRange(row: number, col: number): GridRange {
  return { startRow: row, startCol: col, endRow: row, endCol: col };
}
