import type {
  AutoFilter,
  Cell,
  CellAddress,
  ColumnDimension,
  ColumnFilter,
  ColumnOffset,
  CommentRecord,
  ConditionalFormatRecord,
  GridRange,
  HyperlinkRecord,
  RowDimension,
  Sheet,
  SheetId,
  ValidationRecord,
  Workbook,
} from './types';
import { cellKey, formatCellRef, parseCellKey } from './types';
import { evaluateAutoFilter } from './autofilter';
import {
  MAX_COLUMN,
  MAX_ROW,
  findAreaReferences,
  formatAreaReference,
  parseWholeArea,
  type AreaReference,
} from './references';

/**
 * The single structural-edit range transform (workbook-model §8).
 *
 * A structural edit is a row, column or cell insert/delete. The module plans
 * every range-bearing consequence without mutating the workbook:
 *
 * - authored formulas, defined names and CF/validation rule text (relative and
 *   absolute references, across sheets; `#REF!` when a reference collapses);
 * - merged ranges (shift, expand, clip; a merge that collapses is dropped);
 * - AutoFilter ranges, criteria column offsets and the cleared-filter outcome
 *   (sort-autofilter §3.5);
 * - conditional-format and data-validation ranges (CF §7);
 * - comment and hyperlink cell keys;
 * - the row and column dimension maps (hidden flags, heights and widths move
 *   or renumber with their positions).
 *
 * A cell insert/delete carries a `band`: the perpendicular run of rows (for a
 * column edit) or columns (for a row edit) the shift applies to. Dimension
 * records never move for a banded edit because the sheet's rows and columns
 * themselves do not.
 */

export type StructureAxis = 'row' | 'column';
export type StructureKind = 'insert' | 'delete';

export interface StructuralEdit {
  readonly kind: StructureKind;
  readonly sheetId: SheetId;
  readonly axis: StructureAxis;
  /** Zero-based first row or column. */
  readonly index: number;
  /** Number of rows or columns inserted or deleted. Always at least 1. */
  readonly count: number;
  /**
   * Present for a cell insert/delete: the inclusive span on the perpendicular
   * axis the shift applies to. A row edit shifts only cells inside these
   * columns; a column edit shifts only cells inside these rows. Absent means
   * the whole sheet.
   */
  readonly band?: readonly [number, number];
}

export interface FormulaRewrite {
  readonly sheetId: SheetId;
  readonly row: number;
  readonly col: number;
  readonly before: string;
  readonly after: string;
}

export interface MergeRewrite {
  readonly sheetId: SheetId;
  readonly before: GridRange;
  readonly after: GridRange | null;
}

export interface NameRewrite {
  readonly scope: 'workbook' | 'sheet';
  readonly sheetId: SheetId | null;
  readonly name: string;
  readonly before: string;
  readonly after: string;
}

/**
 * One cell the structural edit moves, removes or rewrites. `from` is `null`
 * only for a cell a caller materialises; `to` is `null` when the edit deletes
 * the cell. `before` and `after` share the same object when nothing about the
 * cell changed but its address, which keeps a large move compact.
 */
export interface CellMove {
  readonly sheetId: SheetId;
  readonly from: CellAddress | null;
  readonly to: CellAddress | null;
  readonly before: Cell | null;
  readonly after: Cell | null;
}

export interface DimensionRenumber<T> {
  readonly before: readonly (readonly [number, T])[];
  readonly after: readonly (readonly [number, T])[];
}

export interface AttachmentRenumber<T> {
  readonly before: readonly (readonly [string, T])[];
  readonly after: readonly (readonly [string, T])[];
}

export interface ConditionalFormatRenumber {
  readonly before: readonly ConditionalFormatRecord[];
  readonly after: readonly ConditionalFormatRecord[];
}

export interface ValidationRenumber {
  readonly before: readonly ValidationRecord[];
  readonly after: readonly ValidationRecord[];
}

export interface AutoFilterRenumber {
  readonly before: AutoFilter | null;
  readonly after: AutoFilter | null;
  /**
   * Rows the previous filter hid and that survive the edit. They become
   * visible when the filter is cleared, so the command removes their hidden
   * flag (sort-autofilter §3.5: "filter-hidden rows are restored").
   */
  readonly restoredHiddenRows: readonly number[];
}

export type TransformBlockCode = 'merge-split' | 'grouped-formula' | 'filter-split';

export interface TransformBlock {
  readonly code: TransformBlockCode;
  readonly message: string;
}

/** Excel's message for an edit that would change part of a merged cell. */
export const MERGED_EDIT_BLOCK_MESSAGE = 'Cannot change part of a merged cell.';
/** A shared or legacy array group cannot follow a structural edit in this milestone. */
export const GROUPED_FORMULA_BLOCK_MESSAGE =
  'Shared and array formulas cannot be moved by a structural edit. The edit was blocked.';
/** A cell edit inside a filter range would split the filtered rectangle. */
export const FILTER_SPLIT_BLOCK_MESSAGE =
  'This edit would split the AutoFilter range. Clear the filter and try again.';

export interface RangeTransformPlan {
  readonly edit: StructuralEdit;
  readonly formulaRewrites: readonly FormulaRewrite[];
  readonly mergeRewrites: readonly MergeRewrite[];
  readonly nameRewrites: readonly NameRewrite[];
  /** Every cell the edit moves, deletes or rewrites, across every sheet. */
  readonly cellMoves: readonly CellMove[];
  readonly rowDimensions: DimensionRenumber<RowDimension>;
  readonly columnDimensions: DimensionRenumber<ColumnDimension>;
  readonly comments: AttachmentRenumber<CommentRecord>;
  readonly hyperlinks: AttachmentRenumber<HyperlinkRecord>;
  readonly validation: ValidationRenumber;
  readonly conditionalFormats: ConditionalFormatRenumber;
  readonly autoFilter: AutoFilterRenumber;
  /** Formulas in shared/array groups, which stay with their group for now. */
  readonly groupedFormulaCount: number;
  /** Why the edit must be refused, or `null` when it can proceed. */
  readonly block: TransformBlock | null;
  readonly changed: boolean;
}

export interface RangeTransformOptions {
  /** Name lookup is case-insensitive, as Excel sheet names are. */
  readonly workbook: Workbook;
}

// ---------------------------------------------------------------------------
// Ranges and cells
// ---------------------------------------------------------------------------

/** Shift or shrink a range for one axis. `null` means every cell was deleted. */
export function transformRange(range: GridRange, edit: StructuralEdit): GridRange | null {
  if (edit.count <= 0) return range;
  if (!rangeIntersectsBand(range, edit)) return range;
  if (edit.axis === 'row') {
    const rows = transformSpan(range.startRow, range.endRow, edit.index, edit.count, edit.kind);
    if (rows === null) return null;
    if (rows[0] === range.startRow && rows[1] === range.endRow) return range;
    return { ...range, startRow: rows[0], endRow: rows[1] };
  }
  const columns = transformSpan(range.startCol, range.endCol, edit.index, edit.count, edit.kind);
  if (columns === null) return null;
  if (columns[0] === range.startCol && columns[1] === range.endCol) return range;
  return { ...range, startCol: columns[0], endCol: columns[1] };
}

/**
 * The address a cell lands at after the edit, or `null` when it is deleted.
 * A cell outside a banded edit's perpendicular span does not move.
 */
export function transformCellAddress(
  row: number,
  col: number,
  edit: StructuralEdit,
): CellAddress | null {
  if (edit.count <= 0) return { row, col };
  const band = edit.band;
  if (band !== undefined) {
    const perpendicular = edit.axis === 'row' ? col : row;
    if (perpendicular < band[0] || perpendicular > band[1]) return { row, col };
  }
  if (edit.axis === 'row') {
    const next = transformCoordinate(row, edit);
    return next === null ? null : { row: next, col };
  }
  const next = transformCoordinate(col, edit);
  return next === null ? null : { row, col: next };
}

function transformCoordinate(value: number, edit: StructuralEdit): number | null {
  const deleteEnd = edit.index + edit.count - 1;
  if (edit.kind === 'insert') return value >= edit.index ? value + edit.count : value;
  if (value < edit.index) return value;
  if (value > deleteEnd) return value - edit.count;
  return null;
}

function transformSpan(
  start: number,
  end: number,
  index: number,
  count: number,
  kind: StructureKind,
): readonly [number, number] | null {
  const deleteEnd = index + count - 1;
  if (kind === 'insert') {
    return [start >= index ? start + count : start, end >= index ? end + count : end];
  }
  if (end < index) return [start, end];
  if (start > deleteEnd) return [start - count, end - count];
  if (start >= index && end <= deleteEnd) return null;
  if (start < index && end > deleteEnd) return [start, end - count];
  if (start < index && end <= deleteEnd) return [start, index - 1];
  return [index, end - count];
}

function rangeIntersectsBand(range: GridRange, edit: StructuralEdit): boolean {
  const band = edit.band;
  if (band === undefined) return true;
  const start = edit.axis === 'row' ? range.startCol : range.startRow;
  const end = edit.axis === 'row' ? range.endCol : range.endRow;
  return end >= band[0] && start <= band[1];
}

// ---------------------------------------------------------------------------
// Formulas and defined names
// ---------------------------------------------------------------------------

/**
 * Rewrite the A1 references in one authored formula. Unsupported constructs
 * (R1C1, 3D ranges, structured/external references, union and intersection)
 * are left byte-for-byte as authored; the calculation layer already marks
 * those cells stale (formula semantics §4 and §9).
 */
export function transformFormulaReferences(
  formula: string,
  edit: StructuralEdit,
  options: RangeTransformOptions,
  hostSheetId: SheetId | null,
): string {
  if (formula.length === 0 || hasPreservingConstruct(formula)) return formula;
  const references = findAreaReferences(formula);
  if (references.length === 0) return formula;
  for (let index = 1; index < references.length; index += 1) {
    const previous = references[index - 1]!;
    const current = references[index]!;
    if (formula.slice(previous.end, current.start).trim() === '') {
      // Excel's intersection operator is v2; keep the authored text.
      return formula;
    }
  }

  let output = '';
  let cursor = 0;
  for (const reference of references) {
    output += formula.slice(cursor, reference.start);
    const transformed = transformReference(reference, edit, options, hostSheetId);
    output += transformed ?? formula.slice(reference.start, reference.end);
    cursor = reference.end;
  }
  return output + formula.slice(cursor);
}

/** Rewrite a defined name's formula. A sheet-scoped name is relative to its sheet. */
export function transformDefinedNameFormula(
  formula: string,
  edit: StructuralEdit,
  options: RangeTransformOptions,
  ownerSheetId: SheetId | null,
): string {
  return transformFormulaReferences(formula, edit, options, ownerSheetId);
}

function transformReference(
  reference: AreaReference,
  edit: StructuralEdit,
  options: RangeTransformOptions,
  hostSheetId: SheetId | null,
): string | null {
  const targetSheetId =
    reference.sheetName === null
      ? hostSheetId
      : findSheetIdByName(options.workbook, reference.sheetName);
  if (targetSheetId !== edit.sheetId) return null;
  if (edit.band !== undefined && reference.kind !== 'cells') {
    // A whole-row or whole-column reference covers the whole perpendicular
    // span, so a banded shift never changes its authored extent.
    return null;
  }
  const transformed = transformArea(reference, edit);
  if (transformed === null) return '#REF!';
  if (areasEqual(transformed, reference)) return null;
  return formatAreaReference(transformed);
}

function transformArea(reference: AreaReference, edit: StructuralEdit): AreaReference | null {
  if (reference.kind === 'columns') {
    if (edit.axis === 'row' || edit.band !== undefined) return reference;
    const columns = transformSpan(
      reference.startCol,
      reference.endCol,
      edit.index,
      edit.count,
      edit.kind,
    );
    if (columns === null) return null;
    return {
      ...reference,
      startCol: columns[0],
      endCol: columns[1],
      single: false,
    };
  }
  if (reference.kind === 'rows') {
    if (edit.axis === 'column' || edit.band !== undefined) return reference;
    const rows = transformSpan(
      reference.startRow,
      reference.endRow,
      edit.index,
      edit.count,
      edit.kind,
    );
    if (rows === null) return null;
    return {
      ...reference,
      startRow: rows[0],
      endRow: rows[1],
      single: false,
    };
  }
  if (!areaIntersectsBand(reference, edit)) return reference;
  const rows =
    edit.axis === 'row'
      ? transformSpan(reference.startRow, reference.endRow, edit.index, edit.count, edit.kind)
      : ([reference.startRow, reference.endRow] as const);
  if (rows === null) return null;
  const columns =
    edit.axis === 'column'
      ? transformSpan(reference.startCol, reference.endCol, edit.index, edit.count, edit.kind)
      : ([reference.startCol, reference.endCol] as const);
  if (columns === null) return null;
  return {
    ...reference,
    startRow: rows[0],
    endRow: rows[1],
    startCol: columns[0],
    endCol: columns[1],
  };
}

function areaIntersectsBand(reference: AreaReference, edit: StructuralEdit): boolean {
  const band = edit.band;
  if (band === undefined) return true;
  const start = edit.axis === 'row' ? reference.startCol : reference.startRow;
  const end = edit.axis === 'row' ? reference.endCol : reference.endRow;
  return end >= band[0] && start <= band[1];
}

// ---------------------------------------------------------------------------
// Merged ranges
// ---------------------------------------------------------------------------

export function transformMergeRange(range: GridRange, edit: StructuralEdit): GridRange | null {
  // A banded edit that touches only part of a merge cannot be represented and
  // is blocked by `planRangeTransform`; the range is left as it is here.
  if (mergeBlocksBandedEdit(range, edit)) return range;
  const transformed = transformRange(range, edit);
  if (transformed === null) return null;
  // A merge that collapses to one cell is no longer a merged range.
  if (transformed.startRow === transformed.endRow && transformed.startCol === transformed.endCol) {
    return null;
  }
  return transformed;
}

/**
 * Excel refuses to change part of a merged cell. A banded cell edit may only
 * move a merge when the selection contains the whole merged range: partly
 * overlapping it on either axis is blocked.
 */
function mergeBlocksBandedEdit(range: GridRange, edit: StructuralEdit): boolean {
  const band = edit.band;
  if (band === undefined) return false;
  const perpendicularStart = edit.axis === 'row' ? range.startCol : range.startRow;
  const perpendicularEnd = edit.axis === 'row' ? range.endCol : range.endRow;
  if (perpendicularEnd < band[0] || perpendicularStart > band[1]) return false;
  if (!(perpendicularStart >= band[0] && perpendicularEnd <= band[1])) return true;
  const start = edit.axis === 'row' ? range.startRow : range.startCol;
  const end = edit.axis === 'row' ? range.endRow : range.endCol;
  const deleteEnd = edit.index + edit.count - 1;
  if (end < edit.index || start > deleteEnd) return false;
  return !(start >= edit.index && end <= deleteEnd);
}

// ---------------------------------------------------------------------------
// Planning
// ---------------------------------------------------------------------------

/** Plan every reference rewrite without mutating the workbook. */
export function planRangeTransform(
  workbook: Workbook,
  edit: StructuralEdit,
  options: RangeTransformOptions = { workbook },
): RangeTransformPlan {
  const targetSheet = workbook.sheets.find((sheet) => sheet.id === edit.sheetId);
  if (targetSheet === undefined) {
    throw new Error(`Unknown sheet for structural edit: ${edit.sheetId}`);
  }
  if (!Number.isInteger(edit.index) || edit.index < 0) {
    throw new RangeError('The structural edit needs a non-negative integer index');
  }
  if (!Number.isInteger(edit.count) || edit.count < 1) {
    throw new RangeError('The structural edit needs a positive integer count');
  }
  const limit = edit.axis === 'row' ? MAX_ROW : MAX_COLUMN;
  if (edit.index + edit.count > limit) {
    throw new RangeError('The structural edit is outside the grid');
  }
  if (
    edit.band !== undefined &&
    (!Number.isInteger(edit.band[0]) ||
      !Number.isInteger(edit.band[1]) ||
      edit.band[0] < 0 ||
      edit.band[1] < edit.band[0] ||
      edit.band[1] >= (edit.axis === 'row' ? MAX_COLUMN : MAX_ROW))
  ) {
    throw new RangeError('The structural edit band is outside the grid');
  }

  const formulaRewrites: FormulaRewrite[] = [];
  const formulaAfterByKey = new Map<string, string>();
  const cellMoves: CellMove[] = [];
  let groupedFormulaCount = 0;
  let block: TransformBlock | null = null;

  for (const sheet of workbook.sheets) {
    const isTarget = sheet.id === edit.sheetId;
    for (const [key, cell] of sheet.cells) {
      const { row, col } = parseCellKey(key);
      if (cell.formulaSource !== undefined) {
        groupedFormulaCount += 1;
        if (isTarget) {
          const moved = transformCellAddress(row, col, edit);
          if (moved === null || moved.row !== row || moved.col !== col) {
            block ??= { code: 'grouped-formula', message: GROUPED_FORMULA_BLOCK_MESSAGE };
          }
        }
        continue;
      }
      let afterFormula: string | undefined;
      if (cell.formula !== null) {
        const after = transformFormulaReferences(cell.formula, edit, options, sheet.id);
        if (after !== cell.formula) {
          formulaRewrites.push({ sheetId: sheet.id, row, col, before: cell.formula, after });
          formulaAfterByKey.set(`${sheet.id}\u0000${key}`, after);
          afterFormula = after;
        }
      }
      if (!isTarget) {
        if (afterFormula !== undefined) {
          cellMoves.push({
            sheetId: sheet.id,
            from: { row, col },
            to: { row, col },
            before: cell,
            after: { ...cell, formula: afterFormula, stale: true },
          });
        }
        continue;
      }
      const to = transformCellAddress(row, col, edit);
      if (to === null) {
        cellMoves.push({
          sheetId: sheet.id,
          from: { row, col },
          to: null,
          before: cell,
          after: null,
        });
        continue;
      }
      if (to.row === row && to.col === col && afterFormula === undefined) continue;
      cellMoves.push({
        sheetId: sheet.id,
        from: { row, col },
        to,
        before: cell,
        after: afterFormula === undefined ? cell : { ...cell, formula: afterFormula, stale: true },
      });
    }
  }

  const mergeRewrites: MergeRewrite[] = [];
  for (const merge of targetSheet.merges) {
    if (mergeBlocksBandedEdit(merge.range, edit)) {
      block ??= { code: 'merge-split', message: MERGED_EDIT_BLOCK_MESSAGE };
    }
    mergeRewrites.push({
      sheetId: targetSheet.id,
      before: merge.range,
      after: transformMergeRange(merge.range, edit),
    });
  }

  const nameRewrites: NameRewrite[] = [];
  for (const name of workbook.names) {
    const after = transformDefinedNameFormula(name.formula, edit, options, null);
    if (after !== name.formula) {
      nameRewrites.push({
        scope: 'workbook',
        sheetId: null,
        name: name.name,
        before: name.formula,
        after,
      });
    }
  }
  for (const sheet of workbook.sheets) {
    for (const name of sheet.names) {
      const after = transformDefinedNameFormula(name.formula, edit, options, sheet.id);
      if (after !== name.formula) {
        nameRewrites.push({
          scope: 'sheet',
          sheetId: sheet.id,
          name: name.name,
          before: name.formula,
          after,
        });
      }
    }
  }

  const rowDimensions = dimensionRenumber(targetSheet.rowDimensions, edit, 'row');
  const columnDimensions = dimensionRenumber(targetSheet.columnDimensions, edit, 'column');
  const comments = attachmentRenumber(targetSheet.comments, edit);
  const hyperlinks = attachmentRenumber(targetSheet.hyperlinks, edit);

  const validationAfter = transformRecords(targetSheet.validation, (record) =>
    transformValidationRecord(record, edit, options, targetSheet.id),
  );
  const conditionalFormatsAfter = transformRecords(targetSheet.conditionalFormats, (record) =>
    transformConditionalFormatRecord(record, edit, options, targetSheet.id),
  );

  const autoFilter = planAutoFilterTransform(targetSheet, edit);
  if (autoFilter.blocked) {
    block ??= { code: 'filter-split', message: FILTER_SPLIT_BLOCK_MESSAGE };
  }

  const changed =
    formulaRewrites.length > 0 ||
    cellMoves.length > 0 ||
    mergeRewrites.some((rewrite) => rewrite.after !== rewrite.before) ||
    nameRewrites.length > 0 ||
    dimensionsChanged(rowDimensions) ||
    dimensionsChanged(columnDimensions) ||
    attachmentChanged(comments) ||
    attachmentChanged(hyperlinks) ||
    validationAfter.changed ||
    conditionalFormatsAfter.changed ||
    autoFilter.changed;

  return {
    edit,
    formulaRewrites,
    mergeRewrites,
    nameRewrites,
    cellMoves,
    rowDimensions,
    columnDimensions,
    comments,
    hyperlinks,
    validation: { before: targetSheet.validation, after: validationAfter.records },
    conditionalFormats: {
      before: targetSheet.conditionalFormats,
      after: conditionalFormatsAfter.records,
    },
    autoFilter: {
      before: targetSheet.autoFilter,
      after: autoFilter.after,
      restoredHiddenRows: autoFilter.restoredHiddenRows,
    },
    groupedFormulaCount,
    block,
    changed,
  };
}

// ---------------------------------------------------------------------------
// Dimensions, attachments, filters and CF/validation records
// ---------------------------------------------------------------------------

/** Renumber a row or column dimension map. A banded edit moves neither. */
export function transformDimensionMap<T>(
  map: ReadonlyMap<number, T>,
  edit: StructuralEdit,
  axis: StructureAxis,
): Map<number, T> {
  const result = new Map<number, T>();
  const moves = edit.band === undefined && edit.axis === axis;
  for (const [index, value] of map) {
    const next = moves ? transformCoordinate(index, edit) : index;
    if (next !== null) result.set(next, value);
  }
  return result;
}

/** Move a comment or hyperlink map with the cells it is attached to. */
export function transformAttachmentMap<T>(
  map: ReadonlyMap<string, T>,
  edit: StructuralEdit,
): Map<string, T> {
  const result = new Map<string, T>();
  for (const [key, value] of map) {
    const address = parseCellKey(key);
    const next = transformCellAddress(address.row, address.col, edit);
    if (next !== null) result.set(cellKey(next.row, next.col), value);
  }
  return result;
}

export interface AutoFilterTransformResult {
  readonly after: AutoFilter | null;
  readonly restoredHiddenRows: readonly number[];
  readonly changed: boolean;
  readonly blocked: boolean;
}

/**
 * AutoFilter §3.5: the range shifts or expands, criteria column offsets move
 * with their columns, and a header-row deletion or a range that collapses
 * below two rows clears the filter. A cell edit that intersects only part of
 * the filter rectangle is refused rather than left half-transformed.
 */
export function planAutoFilterTransform(
  sheet: Sheet,
  edit: StructuralEdit,
): AutoFilterTransformResult {
  const before = sheet.autoFilter;
  if (before === null) {
    return { after: null, restoredHiddenRows: [], changed: false, blocked: false };
  }
  if (edit.band !== undefined) {
    const intersects = filterIntersectsBand(before, edit);
    if (!intersects) {
      return { after: before, restoredHiddenRows: [], changed: false, blocked: false };
    }
    if (!filterCoversBand(before, edit)) {
      return { after: before, restoredHiddenRows: [], changed: false, blocked: true };
    }
  }

  let clears =
    edit.axis === 'row' &&
    edit.kind === 'delete' &&
    before.range.startRow >= edit.index &&
    before.range.startRow <= edit.index + edit.count - 1;
  const range = transformRange(before.range, edit);
  if (range === null) clears = true;
  else if (edit.axis === 'row' && range.endRow - range.startRow < 1) clears = true;
  else if (edit.axis === 'column' && range.endCol - range.startCol < 0) clears = true;

  if (clears || range === null) {
    return {
      after: null,
      restoredHiddenRows: transformHiddenRows(evaluateAutoFilter(sheet, before).filterHidden, edit),
      changed: true,
      blocked: false,
    };
  }

  const columns = new Map<ColumnOffset, ColumnFilter>();
  const startCol = range.startCol;
  const width = range.endCol - range.startCol;
  for (const [offset, columnFilter] of before.columns) {
    if (edit.axis === 'row') {
      columns.set(offset, columnFilter);
      continue;
    }
    const absolute = before.range.startCol + offset;
    const next = transformCoordinate(absolute, edit);
    if (next === null) continue;
    const nextOffset = next - startCol;
    if (nextOffset < 0 || nextOffset > width) continue;
    columns.set(nextOffset, columnFilter);
  }

  const after: AutoFilter = {
    sheetId: before.sheetId,
    range,
    columns,
    ...(before.raw === undefined ? {} : { raw: before.raw }),
  };
  return {
    after,
    restoredHiddenRows: [],
    changed: !sameGridRange(before.range, range) || !sameColumnMap(before.columns, columns),
    blocked: false,
  };
}

function filterIntersectsBand(filter: AutoFilter, edit: StructuralEdit): boolean {
  const band = edit.band;
  if (band === undefined) return true;
  const start = edit.axis === 'row' ? filter.range.startCol : filter.range.startRow;
  const end = edit.axis === 'row' ? filter.range.endCol : filter.range.endRow;
  return end >= band[0] && start <= band[1];
}

function filterCoversBand(filter: AutoFilter, edit: StructuralEdit): boolean {
  const band = edit.band;
  if (band === undefined) return true;
  const start = edit.axis === 'row' ? filter.range.startCol : filter.range.startRow;
  const end = edit.axis === 'row' ? filter.range.endCol : filter.range.endRow;
  return band[0] <= start && band[1] >= end;
}

function transformHiddenRows(rows: Iterable<number>, edit: StructuralEdit): number[] {
  if (edit.axis === 'column') return [...new Set(rows)].sort((left, right) => left - right);
  const result: number[] = [];
  for (const row of rows) {
    const next = transformCoordinate(row, { ...edit, band: undefined });
    if (next !== null) result.push(next);
  }
  return [...new Set(result)].sort((left, right) => left - right);
}

/** Transform one conditional-format record (CF §7). A collapsed range drops the rule. */
export function transformConditionalFormatRecord(
  record: ConditionalFormatRecord,
  edit: StructuralEdit,
  options: RangeTransformOptions,
  sheetId: SheetId,
): ConditionalFormatRecord | null {
  const range = transformRange(record.range, edit);
  if (range === null) return null;
  const rawXml = transformRangeBearingXml(record.rawXml, edit, options, sheetId);
  if (sameGridRange(range, record.range) && rawXml === record.rawXml) return record;
  return { range, rawXml };
}

/** Transform one data-validation record. A collapsed range drops the rule. */
export function transformValidationRecord(
  record: ValidationRecord,
  edit: StructuralEdit,
  options: RangeTransformOptions,
  sheetId: SheetId,
): ValidationRecord | null {
  const range = transformRange(record.range, edit);
  if (range === null) return null;
  const rawXml = transformRangeBearingXml(record.rawXml, edit, options, sheetId);
  if (sameGridRange(range, record.range) && rawXml === record.rawXml) return record;
  return { range, rawXml };
}

/**
 * Rewrite the A1 ranges inside an opaque CF or validation fragment: every
 * space-separated `sqref` area and the authored text of each `<formula>`
 * child. Attributes and children the model does not interpret stay verbatim.
 */
export function transformRangeBearingXml(
  xml: string,
  edit: StructuralEdit,
  options: RangeTransformOptions,
  sheetId: SheetId,
): string {
  let output = xml.replace(
    /(\bsqref\s*=\s*)(["'])([^"']*)\2/g,
    (_whole: string, prefix: string, quote: string, value: string) => {
      const rewritten = transformSqref(value, edit, options, sheetId);
      if (rewritten === null) return '';
      return `${prefix}${quote}${rewritten}${quote}`;
    },
  );
  output = output.replace(
    /<((?:[\w.-]+:)?)formula\d*\b([^>]*)>([\s\S]*?)<\/\1formula\d*>/g,
    (whole: string, prefix: string, attributes: string, text: string) => {
      if (text.trim() === '') return whole;
      const after = transformFormulaReferences(decodeEntities(text), edit, options, sheetId);
      if (after === decodeEntities(text)) return whole;
      const tag = /<((?:[\w.-]+:)?formula\d*)/.exec(whole)?.[1] ?? `${prefix}formula`;
      return `<${tag}${attributes}>${escapeXmlText(after)}</${tag}>`;
    },
  );
  return output;
}

/** `null` when every area in the `sqref` collapsed. */
function transformSqref(
  value: string,
  edit: StructuralEdit,
  options: RangeTransformOptions,
  sheetId: SheetId,
): string | null {
  const areas = splitSqref(value);
  if (areas.length === 0) return value;
  const kept: string[] = [];
  let changed = false;
  for (const areaText of areas) {
    const area = parseWholeArea(areaText);
    if (area === null) {
      kept.push(areaText);
      continue;
    }
    const targetId =
      area.sheetName === null ? sheetId : findSheetIdByName(options.workbook, area.sheetName);
    if (targetId !== edit.sheetId) {
      kept.push(areaText);
      continue;
    }
    const transformed = transformArea(area, edit);
    if (transformed === null) {
      changed = true;
      continue;
    }
    const text = formatAreaReference({ ...transformed, prefix: area.prefix, single: area.single });
    if (text !== areaText) changed = true;
    kept.push(text);
  }
  if (kept.length === 0) return null;
  // An unchanged area keeps the authored text byte-for-byte.
  return changed ? kept.join(' ') : value;
}

/** Split on spaces but not inside a quoted sheet name. */
function splitSqref(value: string): string[] {
  const areas: string[] = [];
  let current = '';
  let quoted = false;
  for (let index = 0; index < value.length; index += 1) {
    const character = value[index] ?? '';
    if (character === "'") quoted = !quoted;
    if (!quoted && /\s/.test(character)) {
      if (current.length > 0) areas.push(current);
      current = '';
      continue;
    }
    current += character;
  }
  if (current.length > 0) areas.push(current);
  return areas;
}

function dimensionRenumber<T>(
  map: ReadonlyMap<number, T>,
  edit: StructuralEdit,
  axis: StructureAxis,
) {
  const before = [...map.entries()]
    .sort((left, right) => left[0] - right[0])
    .map(([index, value]) => [index, value] as const);
  const after = [...transformDimensionMap(map, edit, axis).entries()]
    .sort((left, right) => left[0] - right[0])
    .map(([index, value]) => [index, value] as const);
  return { before, after };
}

function attachmentRenumber<T>(map: ReadonlyMap<string, T>, edit: StructuralEdit) {
  const before = [...map.entries()].map(([key, value]) => [key, value] as const);
  const after = [...transformAttachmentMap(map, edit).entries()].map(
    ([key, value]) => [key, value] as const,
  );
  return { before, after };
}

function transformRecords<T>(
  records: readonly T[],
  transform: (record: T) => T | null,
): { records: T[]; changed: boolean } {
  const result: T[] = [];
  let changed = false;
  for (const record of records) {
    const next = transform(record);
    if (next === null) {
      changed = true;
      continue;
    }
    if (next !== record) changed = true;
    result.push(next);
  }
  return { records: result, changed };
}

function dimensionsChanged<T>(renumber: DimensionRenumber<T>): boolean {
  if (renumber.before.length !== renumber.after.length) return true;
  for (let index = 0; index < renumber.before.length; index += 1) {
    if (renumber.before[index]?.[0] !== renumber.after[index]?.[0]) return true;
  }
  return false;
}

function attachmentChanged<T>(renumber: AttachmentRenumber<T>): boolean {
  if (renumber.before.length !== renumber.after.length) return true;
  for (let index = 0; index < renumber.before.length; index += 1) {
    if (renumber.before[index]?.[0] !== renumber.after[index]?.[0]) return true;
  }
  return false;
}

// ---------------------------------------------------------------------------
// Applying a plan (the M3C prototype and tests)
// ---------------------------------------------------------------------------

/**
 * Apply the formula, merge, name and record rewrites to the model. The M3C
 * prototype uses this to produce rewritten OOXML bytes; the M5 command builds
 * one transaction from the plan instead, because a structural edit also moves
 * cells and dimensions.
 */
export function applyPlanToWorkbook(workbook: Workbook, plan: RangeTransformPlan): void {
  for (const rewrite of plan.formulaRewrites) {
    const sheet = workbook.sheets.find((candidate) => candidate.id === rewrite.sheetId);
    if (sheet === undefined) continue;
    const cell = sheet.cells.get(cellKey(rewrite.row, rewrite.col));
    if (cell === undefined) continue;
    sheet.cells.set(cellKey(rewrite.row, rewrite.col), {
      ...cell,
      formula: rewrite.after,
      stale: true,
    });
  }
  for (const rewrite of plan.mergeRewrites) {
    const sheet = workbook.sheets.find((candidate) => candidate.id === rewrite.sheetId);
    if (sheet === undefined) continue;
    const index = sheet.merges.findIndex((merge) => rangeEquals(merge.range, rewrite.before));
    if (index < 0) continue;
    if (rewrite.after === null) sheet.merges.splice(index, 1);
    else sheet.merges[index] = { range: rewrite.after };
  }
  for (const rewrite of plan.nameRewrites) {
    if (rewrite.scope === 'workbook') {
      const index = workbook.names.findIndex((name) => name.name === rewrite.name);
      if (index >= 0) workbook.names[index] = { ...workbook.names[index]!, formula: rewrite.after };
      continue;
    }
    const sheet = workbook.sheets.find((candidate) => candidate.id === rewrite.sheetId);
    const index = sheet?.names.findIndex((name) => name.name === rewrite.name) ?? -1;
    if (sheet !== undefined && index >= 0 && sheet.names[index] !== undefined) {
      sheet.names[index] = { ...sheet.names[index]!, formula: rewrite.after };
    }
  }
  const sheet = workbook.sheets.find((candidate) => candidate.id === plan.edit.sheetId);
  if (sheet !== undefined) {
    sheet.conditionalFormats.splice(
      0,
      sheet.conditionalFormats.length,
      ...plan.conditionalFormats.after,
    );
    sheet.validation.splice(0, sheet.validation.length, ...plan.validation.after);
  }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function hasPreservingConstruct(formula: string): boolean {
  const stripped = formula.replace(/"(?:[^"]|"")*"/g, '""').replace(/'(?:[^']|'')*'/g, "''");
  if (/[[\]@]/.test(stripped)) return true;
  if (/_xlfn\.|_xlws\./i.test(stripped)) return true;
  // R1C1 notation is v2 and is preserved as stale authored text.
  if (/\bR\d+C\d+\b/i.test(stripped)) return true;
  // 3D ranges: Sheet1:Sheet3!A1 or 'Sheet 1':'Sheet 3'!A1.
  if (/(?:''|[A-Za-z_][\w.]*)\s*:\s*(?:''|[A-Za-z_][\w.]*)\s*!/.test(stripped)) return true;
  return hasUnionComma(formula);
}

/**
 * A comma is Excel's union operator when it is not a function argument
 * separator. Grouping parentheses do not make a comma a function separator:
 * `(A1:A2,C1:C2)` is a union and must stay stale.
 */
function hasUnionComma(formula: string): boolean {
  const functionParens: boolean[] = [];
  let index = 0;
  while (index < formula.length) {
    const character = formula[index] ?? '';
    if (character === '"') {
      index = skipQuoted(formula, index, '"');
      continue;
    }
    if (character === "'") {
      index = skipQuoted(formula, index, "'");
      continue;
    }
    if (character === '(') {
      functionParens.push(isFunctionParen(formula, index));
      index += 1;
      continue;
    }
    if (character === ')') {
      functionParens.pop();
      index += 1;
      continue;
    }
    if (character === ',') {
      const inFunction = functionParens[functionParens.length - 1];
      if (inFunction !== true) return true;
    }
    index += 1;
  }
  return false;
}

function isFunctionParen(formula: string, openParen: number): boolean {
  let index = openParen - 1;
  while (index >= 0 && /\s/.test(formula[index] ?? '')) index -= 1;
  if (index < 0) return false;
  return isIdentifierPart(formula[index] ?? '');
}

function isIdentifierPart(character: string): boolean {
  return /[A-Za-z0-9_.\u00C0-\u02AF]/.test(character);
}

function skipQuoted(text: string, start: number, quote: string): number {
  let index = start + 1;
  while (index < text.length) {
    if (text[index] === quote) {
      if (text[index + 1] === quote) {
        index += 2;
        continue;
      }
      return index + 1;
    }
    index += 1;
  }
  return text.length;
}

function findSheetIdByName(workbook: Workbook, name: string): SheetId | null {
  const upper = name.toUpperCase();
  return workbook.sheets.find((sheet) => sheet.name.toUpperCase() === upper)?.id ?? null;
}

function rangeEquals(left: GridRange, right: GridRange): boolean {
  return sameGridRange(left, right);
}

function sameGridRange(left: GridRange, right: GridRange): boolean {
  return (
    left.startRow === right.startRow &&
    left.startCol === right.startCol &&
    left.endRow === right.endRow &&
    left.endCol === right.endCol
  );
}

function sameColumnMap(
  left: ReadonlyMap<number, unknown>,
  right: ReadonlyMap<number, unknown>,
): boolean {
  if (left.size !== right.size) return false;
  for (const [offset, value] of left) {
    if (right.get(offset) !== value) return false;
  }
  return true;
}

function areasEqual(left: AreaReference, right: AreaReference): boolean {
  return (
    left.kind === right.kind &&
    left.startRow === right.startRow &&
    left.endRow === right.endRow &&
    left.startCol === right.startCol &&
    left.endCol === right.endCol
  );
}

function decodeEntities(text: string): string {
  return text
    .replaceAll('&lt;', '<')
    .replaceAll('&gt;', '>')
    .replaceAll('&quot;', '"')
    .replaceAll('&apos;', "'")
    .replaceAll('&amp;', '&');
}

function escapeXmlText(text: string): string {
  return text.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;');
}

/** The address a plan's formula rewrite points at, for reports and tests. */
export function formulaRewriteAddress(rewrite: FormulaRewrite): string {
  return formatCellRef(rewrite.row, rewrite.col);
}
