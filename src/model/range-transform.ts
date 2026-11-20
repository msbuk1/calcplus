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
  findDefinedName,
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
 * - comment cell keys, hyperlink cell keys and a hyperlink's internal
 *   `location` reference;
 * - the row and column dimension maps (hidden flags, heights and widths move
 *   or renumber with their positions);
 * - Tier B range-bearing parts: drawing anchors, chart series ranges, pivot
 *   source ranges, comment/threaded-comment positions and legacy note anchors
 *   (file-compat §4 and §8).
 *
 * A cell insert/delete carries a `band`: the perpendicular run of rows (for a
 * column edit) or columns (for a row edit) the shift applies to. Dimension
 * records never move for a banded edit because the sheet's rows and columns
 * themselves do not.
 *
 * **Transform or block, never a silently stale anchor** (file-compat §4 and
 * §10.7). Every range-bearing construct is either transformed for the edit or
 * the edit is blocked with a typed `TransformBlock`. Only three Tier B cases
 * block: a chart served by a defined name or another non-range source, a
 * pivot source range the edit would change, and an anchor whose markers cannot
 * be read. Authored formula text with a construct the module cannot rewrite
 * safely (R1C1, a 3D range, a union or intersection, implicit intersection)
 * blocks the edit when it can refer to the edited sheet, rather than being
 * left stale.
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

/**
 * One sheet's hyperlink map after a structural edit. The edited sheet's map is
 * `RangeTransformPlan.hyperlinks`; every other sheet whose internal target
 * moved appears here, so all sheets are renumbered by the same edit (ticket 32).
 */
export interface CrossSheetHyperlinkRenumber extends AttachmentRenumber<HyperlinkRecord> {
  readonly sheetId: SheetId;
  readonly changed: boolean;
}

/** One reference the edit had to break; it now reads `#REF!`. */
export interface BrokenReference {
  readonly kind: 'formula' | 'name' | 'hyperlink' | 'chart';
  readonly label: string;
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

/** The Tier B part kinds this module reads and can renumber (file-compat §4). */
export type OpaquePartKind =
  | 'drawing-anchors'
  | 'vml-anchors'
  | 'chart-series'
  | 'comment-refs'
  | 'threaded-comment-refs'
  | 'pivot-source'
  | 'pivot-location';

/**
 * One opaque part whose range-bearing content the edit changes. `before` is
 * the part exactly as saved, so a transaction can restore it for undo.
 */
export interface OpaquePartRewrite {
  readonly path: string;
  readonly kind: OpaquePartKind;
  readonly before: Uint8Array;
  readonly after: Uint8Array;
}

export type TransformBlockCode =
  | 'merge-split'
  | 'grouped-formula'
  | 'filter-split'
  | 'chart-source'
  | 'pivot-source'
  | 'anchor-untransformable'
  | 'unsupported-formula';

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
/** file-compat §4: a chart served by anything but a plain A1 range is not rewritten. */
export const CHART_SOURCE_BLOCK_MESSAGE =
  'This chart is served by a defined name or another non-range source, so a structural edit cannot keep it correct. The edit was blocked.';
/** file-compat §6.4: pivot caches are never refreshed, so a source edit is refused. */
export const PIVOT_SOURCE_BLOCK_MESSAGE =
  'This edit would change the source range of a PivotTable. Pivot caches are not refreshed in this version, so the edit was blocked.';
/** file-compat §4: "any anchor the module cannot transform" blocks the edit. */
export const ANCHOR_BLOCK_MESSAGE =
  'A drawing or comment on this sheet uses an anchor calcplus cannot transform, so the edit was blocked.';
/** An authored formula keeps a construct the module cannot rewrite safely. */
export const UNSUPPORTED_FORMULA_BLOCK_MESSAGE =
  'A formula uses a reference construct calcplus cannot transform safely, so the edit was blocked.';
/** file-compat §4 and §10.7: a hyperlink anchor the module cannot transform blocks. */
export const HYPERLINK_LOCATION_BLOCK_MESSAGE =
  'A hyperlink points at a cell or range this edit cannot keep correct, so the edit was blocked.';

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
  /** The edited sheet's hyperlinks, moved with their cells and targets. */
  readonly hyperlinks: AttachmentRenumber<HyperlinkRecord>;
  /** Other sheets whose hyperlink target points into the edited sheet. */
  readonly crossSheetHyperlinks: readonly CrossSheetHyperlinkRenumber[];
  /** Every reference the edit could not keep; it now reads `#REF!`. */
  readonly brokenReferences: readonly BrokenReference[];
  readonly validation: ValidationRenumber;
  readonly conditionalFormats: ConditionalFormatRenumber;
  readonly autoFilter: AutoFilterRenumber;
  /** Tier B parts whose anchors or ranges the edit rewrites. */
  readonly opaqueParts: readonly OpaquePartRewrite[];
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
  const brokenReferences: BrokenReference[] = [];
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
        if (block === null && formulaBlocksStructuralEdit(cell.formula, edit, options, sheet.id)) {
          block = {
            code: 'unsupported-formula',
            message: `The formula in ${formatCellRef(row, col)} uses a reference construct calcplus cannot transform safely, so the edit was blocked.`,
          };
        }
        const after = transformFormulaReferences(cell.formula, edit, options, sheet.id);
        if (after !== cell.formula) {
          formulaRewrites.push({ sheetId: sheet.id, row, col, before: cell.formula, after });
          formulaAfterByKey.set(`${sheet.id}\u0000${key}`, after);
          afterFormula = after;
          if (breaksReference(cell.formula, after)) {
            brokenReferences.push({
              kind: 'formula',
              label: `the formula in ${formatCellRef(row, col)} on ${sheet.name}`,
            });
          }
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
    if (block === null && formulaBlocksStructuralEdit(name.formula, edit, options, null)) {
      block = {
        code: 'unsupported-formula',
        message: `The defined name ${name.name} uses a reference construct calcplus cannot transform safely, so the edit was blocked.`,
      };
    }
    const after = transformDefinedNameFormula(name.formula, edit, options, null);
    if (after !== name.formula) {
      nameRewrites.push({
        scope: 'workbook',
        sheetId: null,
        name: name.name,
        before: name.formula,
        after,
      });
      if (breaksReference(name.formula, after)) {
        brokenReferences.push({ kind: 'name', label: `the defined name ${name.name}` });
      }
    }
  }
  for (const sheet of workbook.sheets) {
    for (const name of sheet.names) {
      if (block === null && formulaBlocksStructuralEdit(name.formula, edit, options, sheet.id)) {
        block = {
          code: 'unsupported-formula',
          message: `The defined name ${name.name} uses a reference construct calcplus cannot transform safely, so the edit was blocked.`,
        };
      }
      const after = transformDefinedNameFormula(name.formula, edit, options, sheet.id);
      if (after !== name.formula) {
        nameRewrites.push({
          scope: 'sheet',
          sheetId: sheet.id,
          name: name.name,
          before: name.formula,
          after,
        });
        if (breaksReference(name.formula, after)) {
          brokenReferences.push({
            kind: 'name',
            label: `the defined name ${name.name} on ${sheet.name}`,
          });
        }
      }
    }
  }

  const rowDimensions = dimensionRenumber(targetSheet.rowDimensions, edit, 'row');
  const columnDimensions = dimensionRenumber(targetSheet.columnDimensions, edit, 'column');
  const comments = attachmentRenumber(targetSheet.comments, edit);
  // A hyperlink carries a reference (its internal `location`) as well as a cell
  // key, so it is transformed rather than only renumbered (file-compat §4).
  const hyperlinks = planHyperlinkRenumber(targetSheet, edit, options);
  if (hyperlinks.block !== null) block ??= hyperlinks.block;
  for (const label of hyperlinks.broken) {
    brokenReferences.push({ kind: 'hyperlink', label });
  }
  // A link on another sheet can point into the edited sheet. It is a reference
  // exactly like a formula's, so it moves in the same transaction (ticket 32).
  const crossSheetHyperlinks: CrossSheetHyperlinkRenumber[] = [];
  for (const sheet of workbook.sheets) {
    if (sheet.id === targetSheet.id) continue;
    const cross = planCrossSheetHyperlinkRenumber(sheet, edit, options);
    if (cross.entry !== null) crossSheetHyperlinks.push(cross.entry);
    for (const label of cross.broken) {
      brokenReferences.push({ kind: 'hyperlink', label });
    }
  }

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

  const opaque = planOpaquePartTransforms(workbook, edit, options);
  if (opaque.block !== null) block ??= opaque.block;
  for (let index = 0; index < opaque.broken; index += 1) {
    brokenReferences.push({ kind: 'chart', label: 'a chart series' });
  }

  const changed =
    formulaRewrites.length > 0 ||
    cellMoves.length > 0 ||
    mergeRewrites.some((rewrite) => rewrite.after !== rewrite.before) ||
    nameRewrites.length > 0 ||
    dimensionsChanged(rowDimensions) ||
    dimensionsChanged(columnDimensions) ||
    attachmentChanged(comments) ||
    hyperlinks.changed ||
    crossSheetHyperlinks.length > 0 ||
    validationAfter.changed ||
    conditionalFormatsAfter.changed ||
    autoFilter.changed ||
    opaque.rewrites.length > 0;

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
    crossSheetHyperlinks,
    brokenReferences,
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
    opaqueParts: opaque.rewrites,
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

export type HyperlinkLocationTransform =
  | { readonly ok: true; readonly location: string; readonly changed: boolean }
  | { readonly ok: false; readonly block: TransformBlock };

/**
 * The A1 area a `location` names, with the offset the reference text starts at,
 * so a rewrite can keep any leading whitespace and the `#` some producers write
 * in front of an in-document jump. `null` when the text is not one area.
 */
function parseLocationArea(
  location: string,
): { readonly area: AreaReference; readonly offset: number } | null {
  const afterSpace = location.length - location.trimStart().length;
  const hash = location.startsWith('#', afterSpace) ? 1 : 0;
  const offset = afterSpace + hash;
  const area = parseWholeArea(location.slice(offset));
  return area === null ? null : { area, offset };
}

/**
 * Rewrite a hyperlink's internal `location` (file-compat §4, §10.7).
 *
 * Excel stores a same-workbook jump as `location="Sheet2!A1"`, with an
 * optionally quoted sheet prefix; a location without a prefix is relative to
 * the sheet that holds the hyperlink, and `$` markers are legal. That grammar is
 * a reference in exactly the sense a formula's is, so the edit transforms it or
 * the edit is blocked — a location left "as authored" would silently point at
 * the wrong cells.
 *
 * - A plain cell or range on the edited sheet is transformed like a formula
 *   reference; a target the edit removes collapses to `#REF!`, the module's one
 *   convention for a collapsed reference (names and formulas do the same).
 * - A location naming another sheet is untouched.
 * - A defined name is untouched: the same plan rewrites the name's own formula,
 *   so the jump keeps resolving to the intended cells.
 * - Anything else — arbitrary text, a 3D or structured reference, a name we
 *   cannot resolve — blocks the edit rather than going stale.
 *
 * Only `location` participates. An external `target` is a relationship URL, not
 * a reference, and is never touched.
 */
export function transformHyperlinkLocation(
  location: string,
  edit: StructuralEdit,
  options: RangeTransformOptions,
  hostSheetId: SheetId,
): HyperlinkLocationTransform {
  const parsed = parseLocationArea(location);
  if (parsed === null) {
    if (location.trim() === '') return { ok: true, location, changed: false };
    // A jump an earlier edit already collapsed is a broken link: no edit here
    // can make it staler, and it must not block the sheet from now on.
    if (location.trim().toUpperCase() === '#REF!') return { ok: true, location, changed: false };
    if (findDefinedName(options.workbook, location.trim(), hostSheetId) !== null) {
      return { ok: true, location, changed: false };
    }
    return {
      ok: false,
      block: {
        code: 'anchor-untransformable',
        message: `${HYPERLINK_LOCATION_BLOCK_MESSAGE} (the location "${location}" is not a cell, range or defined name this version understands.)`,
      },
    };
  }
  const { area, offset } = parsed;
  const targetSheetId =
    area.sheetName === null ? hostSheetId : findSheetIdByName(options.workbook, area.sheetName);
  // A location naming a sheet this workbook does not hold is already broken or
  // external; no local edit can make it staler than it is.
  if (targetSheetId === null || targetSheetId !== edit.sheetId) {
    return { ok: true, location, changed: false };
  }
  const transformed = transformArea(area, edit);
  const text =
    transformed === null
      ? '#REF!'
      : formatAreaReference({ ...transformed, prefix: area.prefix, single: area.single });
  const rewritten = `${location.slice(0, offset + area.start)}${text}${location.slice(
    offset + area.end,
  )}`;
  return { ok: true, location: rewritten, changed: rewritten !== location };
}

interface CrossSheetHyperlinkPlan {
  readonly entry: CrossSheetHyperlinkRenumber | null;
  readonly broken: readonly string[];
}

/**
 * Rewrite the hyperlink targets held on a sheet other than the edited one. The
 * anchor cell does not move, so only `location` participates. A location that
 * does not parse cannot be attributed to the edited sheet and is left alone; a
 * 3D, structured or otherwise unsupported reference is not guessed at either.
 * A target the edit removes collapses to `#REF!` and is reported once (ticket
 * 32: broken and marked, never silently stale).
 */
function planCrossSheetHyperlinkRenumber(
  sheet: Sheet,
  edit: StructuralEdit,
  options: RangeTransformOptions,
): CrossSheetHyperlinkPlan {
  const before: (readonly [string, HyperlinkRecord])[] = [];
  const after: (readonly [string, HyperlinkRecord])[] = [];
  const broken: string[] = [];
  let changed = false;
  for (const [key, record] of sheet.hyperlinks) {
    before.push([key, record]);
    let value = record;
    if (record.location !== undefined) {
      // `transformHyperlinkLocation` leaves a bare target alone (it belongs to
      // the sheet that holds the link) and only rewrites a target that names
      // the edited sheet. A location it cannot parse is not this edit's to fix.
      const outcome = transformHyperlinkLocation(record.location, edit, options, sheet.id);
      if (outcome.ok && outcome.changed) {
        value = { ...record, location: outcome.location };
        changed = true;
        if (breaksReference(record.location, outcome.location)) {
          broken.push(`a hyperlink on ${sheet.name}`);
        }
      }
    }
    after.push([key, value]);
  }
  if (!changed) return { entry: null, broken };
  return { entry: { sheetId: sheet.id, before, after, changed }, broken };
}

interface HyperlinkRenumber extends AttachmentRenumber<HyperlinkRecord> {
  readonly changed: boolean;
  readonly block: TransformBlock | null;
  readonly broken: readonly string[];
}

/**
 * Renumber a sheet's hyperlinks: the cell keys move with their cells and each
 * internal `location` is transformed by `transformHyperlinkLocation`. A block
 * anywhere leaves the edit refused rather than saved half-transformed.
 */
function planHyperlinkRenumber(
  sheet: Sheet,
  edit: StructuralEdit,
  options: RangeTransformOptions,
): HyperlinkRenumber {
  const before: (readonly [string, HyperlinkRecord])[] = [];
  const after: (readonly [string, HyperlinkRecord])[] = [];
  const broken: string[] = [];
  let changed = false;
  let block: TransformBlock | null = null;
  for (const [key, record] of sheet.hyperlinks) {
    before.push([key, record]);
    const address = parseCellKey(key);
    const next = transformCellAddress(address.row, address.col, edit);
    if (next === null) {
      changed = true;
      continue;
    }
    let value = record;
    if (record.location !== undefined) {
      const outcome = transformHyperlinkLocation(record.location, edit, options, sheet.id);
      if (!outcome.ok) block ??= outcome.block;
      else if (outcome.changed) {
        value = { ...record, location: outcome.location };
        changed = true;
        if (breaksReference(record.location, outcome.location)) {
          broken.push(`a hyperlink on ${sheet.name}`);
        }
      }
    }
    if (next.row !== address.row || next.col !== address.col) changed = true;
    after.push([cellKey(next.row, next.col), value]);
  }
  return { before, after, changed, block, broken };
}

/** Whether a rewrite added a broken marker, even if the text already had one. */
function breaksReference(before: string, after: string): boolean {
  return brokenMarkerCount(after) > brokenMarkerCount(before);
}

function brokenMarkerCount(text: string): number {
  return text.match(/#REF!/gi)?.length ?? 0;
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
// Tier B parts: drawing anchors, chart series, pivots and comments
// ---------------------------------------------------------------------------

/**
 * The outcome of rewriting one opaque part. A failure is a typed block: the
 * edit must be refused rather than saved with a stale anchor (file-compat §4).
 */
export type PartTransformResult =
  | {
      readonly ok: true;
      readonly xml: string;
      readonly changed: boolean;
      /** References this rewrite had to break; they now read `#REF!`. */
      readonly broken?: number;
    }
  | { readonly ok: false; readonly block: TransformBlock };

export interface OpaquePartTransformPlan {
  readonly rewrites: readonly OpaquePartRewrite[];
  readonly block: TransformBlock | null;
  /** References the edit had to break across these parts. */
  readonly broken: number;
}

const DRAWING_PATH = /^xl\/drawings\/(?!_rels\/)[^/]+\.xml$/i;
const VML_PATH = /^xl\/drawings\/(?!_rels\/)[^/]+\.vml$/i;
const COMMENTS_PATH = /^xl\/comments[^/]*\.xml$/i;
const THREADED_COMMENTS_PATH = /^xl\/threadedComments\/[^/]+\.xml$/i;
const CHART_PATH = /^xl\/charts\/[^/]+\.xml$/i;
const PIVOT_TABLE_PATH = /^xl\/pivotTables\/[^/]+\.xml$/i;
const PIVOT_CACHE_PATH = /^xl\/pivotCache\/pivotCacheDefinition[^/]*\.xml$/i;

function blockResult(code: TransformBlockCode, message: string): PartTransformResult {
  return { ok: false, block: { code, message } };
}

/**
 * The address a drawing anchor marker lands at. A marker that tracks a deleted
 * cell clamps to the cell that shifts into its place (or the last grid cell
 * when the delete runs off the end), so the anchor never points at a row that
 * no longer exists. `colOff`/`rowOff` offsets are relative to the tracked cell
 * and therefore survive untouched (Tier B anchor rule, file-compat §4).
 */
export function transformAnchorMarker(
  row: number,
  col: number,
  edit: StructuralEdit,
): { readonly row: number; readonly col: number } {
  if (edit.count <= 0) return { row, col };
  const band = edit.band;
  if (band !== undefined) {
    const perpendicular = edit.axis === 'row' ? col : row;
    if (perpendicular < band[0] || perpendicular > band[1]) return { row, col };
  }
  if (edit.axis === 'row') return { row: clampAnchorCoordinate(row, edit, MAX_ROW), col };
  return { row, col: clampAnchorCoordinate(col, edit, MAX_COLUMN) };
}

function clampAnchorCoordinate(value: number, edit: StructuralEdit, limit: number): number {
  const next = transformCoordinate(value, edit);
  if (next !== null) return next;
  return Math.min(edit.index, limit - 1);
}

/**
 * Rewrite `xdr:from`/`xdr:to` marker indices in a DrawingML drawing part.
 * `absoluteAnchor` shapes carry EMU offsets, not cell indices, so they are
 * unchanged by construction — they cannot go stale. A one- or two-cell anchor
 * that has a marker the module cannot read blocks the edit.
 */
export function transformDrawingAnchorsXml(
  xml: string,
  edit: StructuralEdit,
  partName = 'the drawing',
): PartTransformResult {
  const anchors = /<(?:[A-Za-z_][\w.-]*:)?(?:oneCellAnchor|twoCellAnchor)\b/g;
  const anchorCount = [...xml.matchAll(anchors)].length;
  let changed = false;
  let markerCount = 0;
  const markerPattern = /<((?:[A-Za-z_][\w.-]*):)?(from|to)\b[^>]*>([\s\S]*?)<\/(?:\1)?\2\s*>/g;
  let result: string;
  try {
    result = xml.replace(markerPattern, (whole, _prefix, _tag, content: string) => {
      markerCount += 1;
      const colMatch = elementInteger(content, 'col');
      const rowMatch = elementInteger(content, 'row');
      if (colMatch === null || rowMatch === null) {
        throw new AnchorReadError(partName);
      }
      const next = transformAnchorMarker(rowMatch, colMatch, edit);
      if (next.row === rowMatch && next.col === colMatch) return whole;
      changed = true;
      const withCol = replaceElementInteger(content, 'col', next.col);
      const withRow = replaceElementInteger(withCol, 'row', next.row);
      return whole.replace(content, withRow);
    });
  } catch (error) {
    if (error instanceof AnchorReadError) {
      return blockResult(
        'anchor-untransformable',
        `${ANCHOR_BLOCK_MESSAGE} (${error.partName} has an anchor marker without a readable column and row.)`,
      );
    }
    throw error;
  }
  if (anchorCount > 0 && markerCount === 0) {
    return blockResult(
      'anchor-untransformable',
      `${ANCHOR_BLOCK_MESSAGE} (${partName} has an anchor without from/to markers.)`,
    );
  }
  if (!changed) return { ok: true, xml, changed: false };
  return { ok: true, xml: result, changed: true };
}

class AnchorReadError extends Error {
  readonly partName: string;

  constructor(partName: string) {
    super(`Unreadable anchor marker in ${partName}`);
    this.partName = partName;
  }
}

/** A leaf element's non-negative integer text, e.g. `<xdr:col>3</xdr:col>`. */
function elementInteger(xml: string, local: string): number | null {
  const pattern = new RegExp(
    `<((?:[A-Za-z_][\\w.-]*):)?${local}\\b[^>]*>\\s*(-?\\d+)\\s*<\\/(?:\\1)?${local}\\s*>`,
  );
  const match = pattern.exec(xml);
  if (match === null) return null;
  const value = Number(match[2]);
  return Number.isInteger(value) && value >= 0 ? value : null;
}

function replaceElementInteger(xml: string, local: string, value: number): string {
  const pattern = new RegExp(
    `(<((?:[A-Za-z_][\\w.-]*):)?${local}\\b[^>]*>)\\s*-?\\d+\\s*(<\\/(?:\\2)?${local}\\s*>)`,
  );
  return xml.replace(pattern, (_whole, open: string, _prefix: string, close: string) => {
    return `${open}${value}${close}`;
  });
}

/**
 * Rewrite the legacy note anchors in a VML drawing part. A note shape carries
 * `<x:Anchor>col, dx, row, dy, col, dx, row, dy</x:Anchor>` plus
 * `<x:Row>`/`<x:Column>` in its client data; both must move together or the
 * note box detaches from its cell.
 */
export function transformVmlAnchorsXml(
  xml: string,
  edit: StructuralEdit,
  partName = 'the legacy drawing',
): PartTransformResult {
  let failure: TransformBlock | null = null;
  let changed = false;
  const shapePattern = /<((?:[A-Za-z_][\w.-]*):)?shape\b[^>]*>([\s\S]*?)<\/(?:\1)?shape\s*>/g;
  const output = xml.replace(shapePattern, (whole, _prefix, body: string) => {
    const anchorPattern = /<((?:[A-Za-z_][\w.-]*):)?Anchor\b[^>]*>([\s\S]*?)<\/(?:\1)?Anchor\s*>/;
    const anchor = anchorPattern.exec(body);
    let first: { readonly row: number; readonly col: number } | null = null;
    let nextBody = body;
    if (anchor !== null) {
      const parsed = parseVmlAnchor(anchor[2] ?? '');
      if (parsed === null) {
        failure ??= {
          code: 'anchor-untransformable',
          message: `${ANCHOR_BLOCK_MESSAGE} (${partName} has a note anchor that is not eight integers.)`,
        };
        return whole;
      }
      first = transformAnchorMarker(parsed[2]!, parsed[0]!, edit);
      const second = transformAnchorMarker(parsed[6]!, parsed[4]!, edit);
      const next = [
        first.col,
        parsed[1]!,
        first.row,
        parsed[3]!,
        second.col,
        parsed[5]!,
        second.row,
        parsed[7]!,
      ];
      if (next.some((value, index) => value !== parsed[index])) {
        changed = true;
        nextBody = nextBody.replace(anchor[0], anchor[0].replace(anchor[2] ?? '', next.join(', ')));
      }
    }
    // The note's client data repeats the first corner's cell; both must agree
    // with the anchor or the note box detaches from its cell.
    for (const local of ['Row', 'Column'] as const) {
      const element = new RegExp(
        `(<((?:[A-Za-z_][\\w.-]*):)?${local}\\b[^>]*>)\\s*(-?\\d+)\\s*(<\\/(?:\\2)?${local}\\s*>)`,
        'g',
      );
      nextBody = nextBody.replace(
        element,
        (match: string, open: string, _elementPrefix: string, value: string, close: string) => {
          const numeric = Number(value);
          const transformed =
            first !== null
              ? local === 'Row'
                ? first.row
                : first.col
              : local === 'Row'
                ? transformAnchorMarker(numeric, 0, { ...edit, band: undefined }).row
                : transformAnchorMarker(0, numeric, { ...edit, band: undefined }).col;
          if (transformed === numeric) return match;
          changed = true;
          return `${open}${transformed}${close}`;
        },
      );
    }
    return whole.replace(body, nextBody);
  });
  if (failure !== null) return { ok: false, block: failure };
  return { ok: true, xml: output, changed };
}

function parseVmlAnchor(text: string): number[] | null {
  const fields = text.split(',').map((field) => field.trim());
  if (fields.length !== 8 || fields.some((field) => !/^-?\d+$/.test(field))) return null;
  return fields.map(Number);
}

/**
 * Rewrite a chart's series ranges (`c:cat`/`c:val` and the scatter/bubble
 * equivalents, each a `numRef`/`strRef`/`multiLvlStrRef` with one `c:f`).
 * A source that is not a single plain A1 area — a defined name, a union, a
 * structured reference — blocks the edit (file-compat §4). A series whose
 * range the edit removes becomes `#REF!`, which Excel treats as a broken
 * series rather than showing the wrong data.
 */
export function transformChartXml(
  xml: string,
  edit: StructuralEdit,
  options: RangeTransformOptions,
  ownerSheetId: SheetId | null,
  partName = 'the chart',
): PartTransformResult {
  let block: TransformBlock | null = null;
  let changed = false;
  let broken = 0;
  // `c:f` is the first child of numRef/strRef/multiLvlStrRef (ECMA-376
  // CT_NumRef/CT_StrRef/CT_MultiLvlStrRef). Matching the parent avoids an
  // unrelated `f` element in a chart extension.
  const pattern =
    /<((?:[A-Za-z_][\w.-]*):)?(numRef|strRef|multiLvlStrRef)\b[^>]*>\s*<((?:[A-Za-z_][\w.-]*):)?f\b[^>]*>([\s\S]*?)<\/(?:\3)?f\s*>/g;
  const output = xml.replace(
    pattern,
    (whole, _parentPrefix, _parent: string, _prefix, text: string) => {
      const decoded = decodeEntities(text).trim();
      if (decoded === '') return whole;
      const outcome = transformChartSource(decoded, edit, options, ownerSheetId);
      if (outcome === null) {
        block ??= {
          code: 'chart-source',
          message: `${CHART_SOURCE_BLOCK_MESSAGE} (${partName} points at "${decoded}".)`,
        };
        return whole;
      }
      if (outcome === decoded) return whole;
      changed = true;
      if (breaksReference(decoded, outcome)) broken += 1;
      return whole.replace(text, escapeXmlText(outcome));
    },
  );
  if (block !== null) return { ok: false, block };
  return { ok: true, xml: output, changed, broken };
}

function transformChartSource(
  source: string,
  edit: StructuralEdit,
  options: RangeTransformOptions,
  ownerSheetId: SheetId | null,
): string | null {
  if (source.includes(',')) return null;
  const area = parseWholeArea(source);
  if (area === null) return null;
  if (area.sheetName === null && ownerSheetId === null) return null;
  const targetSheetId =
    area.sheetName === null ? ownerSheetId : findSheetIdByName(options.workbook, area.sheetName);
  // A sheet name that does not resolve is not this workbook's sheet (a broken
  // or external chart); no edit here can make it staler than it already is.
  if (targetSheetId === null) return source;
  if (targetSheetId !== edit.sheetId) return source;
  const transformed = transformArea(area, edit);
  if (transformed === null) return '#REF!';
  if (areasEqual(transformed, area)) return source;
  return formatAreaReference({ ...transformed, prefix: area.prefix, single: area.single });
}

/**
 * Rewrite `<comment ref>` and `<threadedComment ref>` positions. A comment on
 * a deleted cell is dropped, matching the model's own comment map, so the part
 * cannot point at a cell that now belongs to another row.
 */
export function transformCommentRefsXml(
  xml: string,
  edit: StructuralEdit,
  partName = 'the comments part',
): PartTransformResult {
  return transformCellRefElements(xml, 'comment', edit, partName);
}

export function transformThreadedCommentRefsXml(
  xml: string,
  edit: StructuralEdit,
  partName = 'the threaded comments part',
): PartTransformResult {
  return transformCellRefElements(xml, 'threadedComment', edit, partName);
}

function transformCellRefElements(
  xml: string,
  local: string,
  edit: StructuralEdit,
  partName: string,
): PartTransformResult {
  const pattern = new RegExp(
    `<((?:[A-Za-z_][\\w.-]*):)?${local}\\b([^>]*?)(\\/>|>[\\s\\S]*?<\\/(?:\\1)?${local}\\s*>)`,
    'g',
  );
  let block: TransformBlock | null = null;
  let changed = false;
  const output = xml.replace(
    pattern,
    (whole: string, prefix: string | undefined, attributes: string, tail: string) => {
      const match = /(\bref\s*=\s*)(["'])([^"']*)\2/.exec(attributes);
      if (match === null) return whole;
      const range = parseRangeText(match[3] ?? '');
      if (range === null || range.startRow !== range.endRow || range.startCol !== range.endCol) {
        block ??= {
          code: 'anchor-untransformable',
          message: `${ANCHOR_BLOCK_MESSAGE} (${partName} has a comment reference that is not one cell.)`,
        };
        return whole;
      }
      const next = transformCellAddress(range.startRow, range.startCol, edit);
      if (next === null) {
        changed = true;
        return '';
      }
      if (next.row === range.startRow && next.col === range.startCol) return whole;
      changed = true;
      const target = formatCellRef(next.row, next.col);
      const rewritten = `${attributes.slice(0, match.index)}${match[1]}${match[2]}${escapeXmlAttributeValue(
        target,
      )}${match[2]}${attributes.slice(match.index + match[0].length)}`;
      return `<${prefix ?? ''}${local}${rewritten}${tail}`;
    },
  );
  if (block !== null) return { ok: false, block };
  return { ok: true, xml: output, changed };
}

/**
 * Check a pivot cache definition's source. The source range is never rewritten
 * (file-compat §4 and §6.4); an edit that would change it blocks. A cache with
 * an external, consolidation or table-name source has no A1 range to compare
 * and is left alone.
 */
export function transformPivotCacheXml(
  xml: string,
  edit: StructuralEdit,
  options: RangeTransformOptions,
  ownerSheetId: SheetId | null,
  partName = 'the pivot cache',
): PartTransformResult {
  const pattern = /<((?:[A-Za-z_][\w.-]*):)?worksheetSource\b([^>]*?)\/?>/g;
  let block: TransformBlock | null = null;
  for (const match of xml.matchAll(pattern)) {
    const attributes = parseXmlAttributes(match[2] ?? '');
    const ref = attributes.get('ref');
    if (ref === undefined) continue;
    const sheetName = attributes.get('sheet');
    const sourceSheetId =
      sheetName === undefined
        ? ownerSheetId
        : (findSheetIdByName(options.workbook, sheetName) ?? null);
    if (sourceSheetId !== edit.sheetId) continue;
    const range = parseRangeText(ref);
    if (range === null) {
      block ??= {
        code: 'pivot-source',
        message: `${PIVOT_SOURCE_BLOCK_MESSAGE} (${partName} has a source that is not an A1 range.)`,
      };
      continue;
    }
    const transformed = transformRange(range, edit);
    if (transformed === null || !sameGridRange(range, transformed)) {
      block ??= {
        code: 'pivot-source',
        message: `${PIVOT_SOURCE_BLOCK_MESSAGE} (${partName} source is ${ref}.)`,
      };
    }
  }
  if (block !== null) return { ok: false, block };
  return { ok: true, xml, changed: false };
}

/**
 * Move a pivot table's output rectangle with the cells it occupies
 * (workbook-model §8: "anchor indexes for opaque charts, drawings and
 * pivots"). The source range is a separate, block-only rule. A rectangle the
 * edit removes cannot be represented, so the edit is blocked rather than
 * silently orphaning the pivot.
 */
export function transformPivotTableXml(
  xml: string,
  edit: StructuralEdit,
  partName = 'the pivot table',
): PartTransformResult {
  let block: TransformBlock | null = null;
  let changed = false;
  const pattern = /<((?:[A-Za-z_][\w.-]*):)?location\b([^>]*)>/g;
  const output = xml.replace(
    pattern,
    (whole: string, prefix: string | undefined, attributes: string) => {
      const match = /(\bref\s*=\s*)(["'])([^"']*)\2/.exec(attributes);
      if (match === null) return whole;
      const range = parseRangeText(match[3] ?? '');
      if (range === null) {
        block ??= {
          code: 'anchor-untransformable',
          message: `${ANCHOR_BLOCK_MESSAGE} (${partName} has a location that is not an A1 range.)`,
        };
        return whole;
      }
      const transformed = transformRange(range, edit);
      if (transformed === null) {
        block ??= {
          code: 'anchor-untransformable',
          message: `${ANCHOR_BLOCK_MESSAGE} (${partName} output is entirely inside the removed cells.)`,
        };
        return whole;
      }
      if (sameGridRange(range, transformed)) return whole;
      changed = true;
      const text = formatRangeText(transformed);
      const rewritten = `${attributes.slice(0, match.index)}${match[1]}${match[2]}${escapeXmlAttributeValue(
        text,
      )}${match[2]}${attributes.slice(match.index + match[0].length)}`;
      return `<${prefix ?? ''}location${rewritten}>`;
    },
  );
  if (block !== null) return { ok: false, block };
  return { ok: true, xml: output, changed };
}

/**
 * Discover every Tier B part the edit touches and transform it. Sheet-owned
 * parts (drawings, comments) come from the sheet's relationship targets;
 * charts and pivot caches are workbook parts attributed through the drawing
 * and pivot-table relationships. Parts that are not UTF-8 XML are skipped: the
 * module cannot read them, and they carry no A1 text to leave stale.
 */
export function planOpaquePartTransforms(
  workbook: Workbook,
  edit: StructuralEdit,
  options: RangeTransformOptions = { workbook },
): OpaquePartTransformPlan {
  const sheet = workbook.sheets.find((candidate) => candidate.id === edit.sheetId);
  if (sheet === undefined) return { rewrites: [], block: null, broken: 0 };

  const owners = resolvePartOwners(workbook);
  const candidates = new Map<string, Uint8Array>();
  for (const [path, part] of sheet.opaqueParts) {
    if (
      DRAWING_PATH.test(path) ||
      VML_PATH.test(path) ||
      COMMENTS_PATH.test(path) ||
      THREADED_COMMENTS_PATH.test(path) ||
      PIVOT_TABLE_PATH.test(path)
    ) {
      candidates.set(path, part.bytes);
    }
  }
  for (const [path, part] of workbook.opaqueParts) {
    if (CHART_PATH.test(path) || PIVOT_CACHE_PATH.test(path)) candidates.set(path, part.bytes);
  }

  const rewrites: OpaquePartRewrite[] = [];
  let block: TransformBlock | null = null;
  let broken = 0;
  for (const path of [...candidates.keys()].sort()) {
    const bytes = candidates.get(path);
    if (bytes === undefined) continue;
    const xml = decodePartXml(bytes);
    if (xml === null) continue;
    let transformed: PartTransformResult;
    if (CHART_PATH.test(path)) {
      transformed = transformChartXml(xml, edit, options, owners.charts.get(path) ?? null, path);
    } else if (PIVOT_CACHE_PATH.test(path)) {
      transformed = transformPivotCacheXml(
        xml,
        edit,
        options,
        owners.pivots.get(path) ?? null,
        path,
      );
    } else if (PIVOT_TABLE_PATH.test(path)) {
      transformed = transformPivotTableXml(xml, edit, path);
    } else if (VML_PATH.test(path)) {
      transformed = transformVmlAnchorsXml(xml, edit, path);
    } else if (THREADED_COMMENTS_PATH.test(path)) {
      transformed = transformThreadedCommentRefsXml(xml, edit, path);
    } else if (COMMENTS_PATH.test(path)) {
      transformed = transformCommentRefsXml(xml, edit, path);
    } else {
      transformed = transformDrawingAnchorsXml(xml, edit, path);
    }
    if (!transformed.ok) {
      block ??= transformed.block;
      continue;
    }
    broken += transformed.broken ?? 0;
    if (!transformed.changed) continue;
    rewrites.push({
      path,
      kind: opaqueKindForPath(path),
      before: bytes,
      after: encodePartXml(transformed.xml, bytes),
    });
  }
  // A blocked edit applies nothing, so a blocked plan carries no rewrites.
  if (block !== null) return { rewrites: [], block, broken: 0 };
  return { rewrites, block, broken };
}

function opaqueKindForPath(path: string): OpaquePartKind {
  if (CHART_PATH.test(path)) return 'chart-series';
  if (PIVOT_CACHE_PATH.test(path)) return 'pivot-source';
  if (PIVOT_TABLE_PATH.test(path)) return 'pivot-location';
  if (VML_PATH.test(path)) return 'vml-anchors';
  if (THREADED_COMMENTS_PATH.test(path)) return 'threaded-comment-refs';
  if (COMMENTS_PATH.test(path)) return 'comment-refs';
  return 'drawing-anchors';
}

interface PartOwners {
  readonly charts: ReadonlyMap<string, SheetId>;
  readonly pivots: ReadonlyMap<string, SheetId>;
}

/** Attribute chart and pivot-cache parts to the sheet whose relationships reach them. */
function resolvePartOwners(workbook: Workbook): PartOwners {
  const charts = new Map<string, SheetId>();
  const pivots = new Map<string, SheetId>();
  for (const sheet of workbook.sheets) {
    for (const path of sheet.opaqueParts.keys()) {
      if (DRAWING_PATH.test(path)) {
        for (const relationship of readRelationshipTargets(workbook, path)) {
          if (relationship.type.endsWith('/chart')) charts.set(relationship.target, sheet.id);
        }
      } else if (PIVOT_TABLE_PATH.test(path)) {
        for (const relationship of readRelationshipTargets(workbook, path)) {
          if (relationship.type.endsWith('/pivotCacheDefinition')) {
            pivots.set(relationship.target, sheet.id);
          }
        }
      }
    }
  }
  return { charts, pivots };
}

interface RelationshipTarget {
  readonly type: string;
  readonly target: string;
}

/**
 * Read a part's relationships without depending on the OOXML layer: the model
 * module must stay usable from the browser model alone.
 */
function readRelationshipTargets(workbook: Workbook, partPath: string): RelationshipTarget[] {
  const relsPath = relsPathFor(partPath);
  const bytes = workbook.opaqueParts.get(relsPath)?.bytes;
  if (bytes === undefined) return [];
  const xml = decodePartXml(bytes);
  if (xml === null) return [];
  const targets: RelationshipTarget[] = [];
  for (const match of xml.matchAll(/<Relationship\b([^>]*?)\/?>/g)) {
    const attributes = parseXmlAttributes(match[1] ?? '');
    if (attributes.get('TargetMode') === 'External') continue;
    const type = attributes.get('Type') ?? '';
    const target = attributes.get('Target');
    if (type === '' || target === undefined) continue;
    targets.push({ type, target: resolveRelativePartPath(partPath, target) });
  }
  return targets;
}

function relsPathFor(partPath: string): string {
  const slash = partPath.lastIndexOf('/');
  const directory = slash < 0 ? '' : partPath.slice(0, slash + 1);
  const fileName = slash < 0 ? partPath : partPath.slice(slash + 1);
  return `${directory}_rels/${fileName}.rels`;
}

function resolveRelativePartPath(partPath: string, target: string): string {
  if (target.startsWith('/')) return normalizePartPath(target.slice(1));
  const segments = partPath.split('/');
  segments.pop();
  return normalizePartPath([...segments, ...target.split('/')].join('/'));
}

function normalizePartPath(path: string): string {
  const segments: string[] = [];
  for (const segment of path.split('/')) {
    if (segment === '' || segment === '.') continue;
    if (segment === '..') segments.pop();
    else segments.push(segment);
  }
  return segments.join('/');
}

function parseXmlAttributes(text: string): Map<string, string> {
  const attributes = new Map<string, string>();
  for (const match of text.matchAll(/([A-Za-z_][\w.:-]*)\s*=\s*(["'])(.*?)\2/g)) {
    const name = match[1] ?? '';
    const local = name.includes(':') ? name.slice(name.lastIndexOf(':') + 1) : name;
    attributes.set(local, decodeEntities(match[3] ?? ''));
  }
  return attributes;
}

const PART_UTF8_BOM = [0xef, 0xbb, 0xbf] as const;

function hasUtf8Bom(bytes: Uint8Array): boolean {
  return (
    bytes[0] === PART_UTF8_BOM[0] && bytes[1] === PART_UTF8_BOM[1] && bytes[2] === PART_UTF8_BOM[2]
  );
}

/** `null` when the part is not UTF-8 XML (UTF-16, binary or damaged): leave it alone. */
function decodePartXml(bytes: Uint8Array): string | null {
  if (bytes.length >= 2) {
    const first = bytes[0];
    const second = bytes[1];
    if ((first === 0xff && second === 0xfe) || (first === 0xfe && second === 0xff)) return null;
  }
  if (bytes.includes(0)) return null;
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(bytes);
  } catch {
    return null;
  }
}

function encodePartXml(xml: string, original: Uint8Array): Uint8Array {
  const text = hasUtf8Bom(original) ? `\uFEFF${xml}` : xml;
  return new TextEncoder().encode(text);
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
  const stripped = stripDoubleQuoted(formula).replace(/'(?:[^']|'')*'/g, "''");
  if (/[[\]@]/.test(stripped)) return true;
  // R1C1 notation is v2 and is preserved as stale authored text.
  if (/\bR\d+C\d+\b/i.test(stripped)) return true;
  // 3D ranges: Sheet1:Sheet3!A1 or 'Sheet 1':'Sheet 3'!A1.
  if (/(?:''|[A-Za-z_][\w.]*)\s*:\s*(?:''|[A-Za-z_][\w.]*)\s*!/.test(stripped)) return true;
  return hasUnionComma(formula);
}

/**
 * Whether the edit must be refused because this authored formula keeps a
 * construct the module cannot rewrite and could point at the edited sheet.
 *
 * - A1 references are rewritten, including inside `_xlfn.`/`_xlws.` calls:
 *   the function name is not a reference, so rewriting its arguments is safe.
 * - External-workbook references (`[1]Sheet1!A1`) are never rewritten and are
 *   unaffected by a local edit (file-compat §6.2), so they never block.
 * - Structured table references may go stale and are reported (file-compat
 *   §4), so they do not block; a sibling A1 reference in the same formula does.
 * - R1C1, 3D ranges, unions, intersections and implicit intersection are
 *   relative to their host and cannot be rewritten; they block when they touch
 *   the edited sheet.
 */
function formulaBlocksStructuralEdit(
  formula: string,
  edit: StructuralEdit,
  options: RangeTransformOptions,
  hostSheetId: SheetId | null,
): boolean {
  if (formula.length === 0) return false;
  const preserving = hasPreservingConstruct(formula);
  // An intersection is the one untransformable construct the cheap scan does
  // not catch; two references must be separated by whitespace to form one.
  const intersection = !preserving && /\S\s+\S/.test(formula) && hasReferenceIntersection(formula);
  if (!preserving && !intersection) return false;

  const stripped = stripDoubleQuoted(formula);
  if (threeDimensionalSpanIncludes(stripped, edit, options)) return true;

  // Drop external-link references and table structured references, then ask
  // which sheet the remaining A1 tokens resolve to.
  const localText = stripExternalReferences(stripped).replace(/\[[^\]]*\]/g, '');
  const references = findAreaReferences(localText);
  for (const reference of references) {
    const target =
      reference.sheetName === null
        ? hostSheetId
        : findSheetIdByName(options.workbook, reference.sheetName);
    if (target === edit.sheetId) return true;
  }
  if (hostSheetId !== edit.sheetId) return false;
  // A construct with no parseable reference is relative to its host cell.
  if (/\bR\d+C\d+\b/i.test(stripped) || stripped.includes('@')) return true;
  return intersection || hasUnionComma(formula);
}

function stripDoubleQuoted(formula: string): string {
  return formula.replace(/"(?:[^"]|"")*"/g, '""');
}

/** A 3D span (`Sheet1:Sheet3!A1`) whose sheet range contains the edited sheet. */
function threeDimensionalSpanIncludes(
  formula: string,
  edit: StructuralEdit,
  options: RangeTransformOptions,
): boolean {
  const pattern =
    /(?:'((?:[^']|'')*)'|([A-Za-z_][\w.]*))\s*:\s*(?:'((?:[^']|'')*)'|([A-Za-z_][\w.]*))\s*!/g;
  for (const match of formula.matchAll(pattern)) {
    const first = unquoteSheetName(match[1] ?? match[2] ?? '');
    const second = unquoteSheetName(match[3] ?? match[4] ?? '');
    const from = options.workbook.sheets.findIndex((sheet) => sheetsNameEqual(sheet.name, first));
    const to = options.workbook.sheets.findIndex((sheet) => sheetsNameEqual(sheet.name, second));
    if (from < 0 || to < 0) continue;
    const start = Math.min(from, to);
    const end = Math.max(from, to);
    const target = options.workbook.sheets.findIndex((sheet) => sheet.id === edit.sheetId);
    if (target >= start && target <= end) return true;
  }
  return false;
}

function unquoteSheetName(name: string): string {
  return name.replaceAll("''", "'");
}

function sheetsNameEqual(left: string, right: string): boolean {
  return left.toUpperCase() === right.toUpperCase();
}

function stripExternalReferences(formula: string): string {
  return formula.replace(/\[\d+\][^!'"]*!\$?[A-Za-z]{1,3}\$?\d*(?::\$?[A-Za-z]{1,3}\$?\d*)?/g, '');
}

/** Whether two references are separated only by whitespace (Excel intersection). */
function hasReferenceIntersection(formula: string): boolean {
  const references = findAreaReferences(formula);
  for (let index = 1; index < references.length; index += 1) {
    const previous = references[index - 1]!;
    const current = references[index]!;
    if (formula.slice(previous.end, current.start).trim() === '') return true;
  }
  return false;
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

/**
 * Parse a plain `A1`, `A1:B2` or `$A$1:$B$2` reference. Anything the range
 * module cannot read as one rectangular area returns `null`.
 */
function parseRangeText(text: string): GridRange | null {
  const trimmed = text.trim();
  if (trimmed === '') return null;
  const [startText, endText = startText] = trimmed.split(':');
  const start = parseCellText(startText ?? '');
  const end = parseCellText(endText);
  if (start === null || end === null) return null;
  return {
    startRow: Math.min(start.row, end.row),
    startCol: Math.min(start.col, end.col),
    endRow: Math.max(start.row, end.row),
    endCol: Math.max(start.col, end.col),
  };
}

function parseCellText(text: string): { readonly row: number; readonly col: number } | null {
  const match = /^\$?([A-Z]{1,3})\$?(\d{1,7})$/i.exec(text.trim());
  if (match === null) return null;
  const row = Number(match[2]) - 1;
  if (!Number.isInteger(row) || row < 0) return null;
  return { row, col: letterIndex(match[1] ?? '') };
}

function formatRangeText(range: GridRange): string {
  const start = formatCellRef(range.startRow, range.startCol);
  const end = formatCellRef(range.endRow, range.endCol);
  return start === end ? start : `${start}:${end}`;
}

function letterIndex(letters: string): number {
  let value = 0;
  for (const character of letters.toUpperCase()) {
    value = value * 26 + (character.charCodeAt(0) - 64);
  }
  return value - 1;
}

function escapeXmlAttributeValue(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;');
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
