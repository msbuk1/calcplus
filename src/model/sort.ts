import type {
  Cell,
  CellAddress,
  CommentRecord,
  GridRange,
  HyperlinkRecord,
  Sheet,
  SheetId,
  StyleRecord,
  Workbook,
} from './types';
import {
  cellIsEmpty,
  columnToLetters,
  getCell,
  getSheetById,
  parseRangeRef,
  singleCellRange,
} from './types';
import type { CellValue } from './value';
import { BLANK } from './value';
import { MAX_ROW, findAreaReferences, formatAreaReference, type AreaReference } from './references';
import { getStyle } from './workbook';

/**
 * The sort engine (sort-autofilter §2).
 *
 * A sort is a row permutation over the target range, never a value rewrite:
 * every cell moves with its row, and a formula that moves shifts its relative
 * references by the row delta exactly as if the cell had been moved by hand.
 *
 * The module is pure: it reads the workbook model and returns a plan. The
 * caller applies the plan as one transaction and keeps it for undo (§2.7);
 * this module knows nothing about the command registry.
 */

/** Excel's key limit, matching the Sort dialog (§2.3). */
export const MAX_SORT_KEYS = 64;

export type SortDirection = 'ascending' | 'descending';

export interface SortKey {
  /** Absolute column index. Every key column must lie inside the sorted range. */
  readonly column: number;
  readonly direction: SortDirection;
}

/** Excel's message when merges inside the range are not uniform one-row merges. */
export const MERGED_SIZE_MESSAGE =
  'This operation requires the merged cells to be identically sized.';

/** Excel's message for an edit that would change part of a legacy array. */
export const ARRAY_BLOCK_MESSAGE = 'You cannot change part of an array.';

export const TABLE_BLOCK_MESSAGE = 'A sort cannot include cells that belong to a table.';

export const TOO_MANY_KEYS_MESSAGE = `A sort can use at most ${MAX_SORT_KEYS} keys.`;

export const KEY_OUTSIDE_RANGE_MESSAGE = 'A sort key must be a column inside the sorted range.';

export const HIDDEN_ROWS_WARNING =
  'The sorted range includes hidden rows. They stay in place; unhide them first to sort every row.';

export const SUB_RANGE_WARNING =
  'The sorted rows are part of a larger block. Sorting them separates them from the surrounding data.';

/** A warning the caller shows once before it applies the sort. */
export type SortWarningCode = 'hidden-rows' | 'sub-range';

export interface SortWarning {
  readonly code: SortWarningCode;
  readonly message: string;
  /** Hidden data rows inside the range; present for `hidden-rows`. */
  readonly count?: number;
}

export type SortBlockCode =
  'merged-cells' | 'table' | 'array-formula' | 'too-many-keys' | 'key-outside-range';

export type SortStopReason = 'no-keys' | 'too-few-rows' | 'nothing-to-order';

/** Destination row -> source row, for the rows that change position. */
export interface SortRowMove {
  readonly fromRow: number;
  readonly toRow: number;
}

/**
 * One address whose content the sort changes.
 *
 * `before` is the cell at `address` before the sort and `after` is the cell
 * that lands there, with its formula already shifted. Applying every record is
 * the sorted state; applying `before` everywhere restores the prior state.
 * Cells are shared with the model, which treats them as immutable.
 */
export interface SortCellRecord {
  readonly address: CellAddress;
  /** Where the content came from, or `null` when the address is left empty. */
  readonly source: CellAddress | null;
  readonly before: Cell | null;
  readonly after: Cell | null;
}

/** An authored formula that the move rewrote, at its destination address. */
export interface SortFormulaRewrite {
  readonly address: CellAddress;
  readonly before: string;
  readonly after: string;
}

export interface SortMergeRewrite {
  readonly before: GridRange;
  readonly after: GridRange;
}

/**
 * A comment or hyperlink that moves with its row. The applier removes every
 * comment and hyperlink in the moved rows first, then inserts these records,
 * because an incoming cell can leave a destination attachment empty.
 */
export interface SortAttachmentMove<T> {
  readonly from: CellAddress;
  readonly to: CellAddress;
  readonly record: T;
}

export interface SortPlan {
  readonly sheetId: SheetId;
  readonly range: GridRange;
  /** `true` when the first row is a header and stays in place. */
  readonly hasHeaders: boolean;
  /** The labels the Sort dialog shows, in key order. */
  readonly keyLabels: readonly string[];
  readonly keys: readonly SortKey[];
  /**
   * Destination offset inside the range -> source row. The array is total: a
   * row that keeps its position maps to itself, and so does the header row.
   */
  readonly rowPermutation: readonly number[];
  /** The same mapping as `rowPermutation`, listed only for rows that move. */
  readonly rowMoves: readonly SortRowMove[];
  readonly cellRecords: readonly SortCellRecord[];
  readonly formulaRewrites: readonly SortFormulaRewrite[];
  readonly mergeRewrites: readonly SortMergeRewrite[];
  readonly commentMoves: readonly SortAttachmentMove<CommentRecord>[];
  readonly hyperlinkMoves: readonly SortAttachmentMove<HyperlinkRecord>[];
  readonly visibleRowCount: number;
  readonly hiddenRowCount: number;
  readonly crossesHiddenRows: boolean;
}

export type SortResult =
  | { readonly kind: 'sorted'; readonly plan: SortPlan; readonly warnings: readonly SortWarning[] }
  | {
      readonly kind: 'unchanged';
      readonly reason: SortStopReason;
      readonly range: GridRange;
      readonly warnings: readonly SortWarning[];
    }
  | {
      readonly kind: 'blocked';
      readonly code: SortBlockCode;
      readonly message: string;
      readonly range: GridRange;
    };

export interface SortRequest {
  readonly workbook: Workbook;
  readonly sheetId: SheetId;
  /** The selection. A single cell expands to the current contiguous region. */
  readonly selection: GridRange;
  readonly keys: readonly SortKey[];
  /** `null` guesses with the §2.2 heuristic; `true` or `false` is the override. */
  readonly hasHeaders: boolean | null;
  /** Filter-hidden rows from the AutoFilter layer. Manual hidden rows come from the sheet. */
  readonly filterHiddenRows?: ReadonlySet<number>;
  /** Every table range on the sheet; tables are blocked inside the range (§2.5). */
  readonly tableRanges?: readonly GridRange[];
}

export interface HeaderHeuristicInput {
  readonly workbook: Workbook;
  readonly sheet: Sheet;
  readonly range: GridRange;
  /** Absolute key columns. Omitted or empty means every column in the range. */
  readonly keyColumns?: readonly number[];
}

// ---------------------------------------------------------------------------
// Comparator (§2.4)
// ---------------------------------------------------------------------------

/**
 * Excel's documented ascending character order for text: digits, space,
 * punctuation, then letters. Apostrophes and hyphens are ignored (§2.4), so
 * they are absent from this sequence. Localization is v2, so the table is
 * fixed rather than locale-derived, and folding is ASCII-only to stay
 * independent of the host locale.
 */
const PUNCTUATION_ORDER = '!"#$%&()*,./:;?@[\\]^_`{|}~+<=>';

const PUNCTUATION_WEIGHT = new Map<string, number>();
for (let index = 0; index < PUNCTUATION_ORDER.length; index += 1) {
  PUNCTUATION_WEIGHT.set(PUNCTUATION_ORDER[index]!, index);
}

const DIGIT_WEIGHT = 0x1000;
const SPACE_WEIGHT = 0x2000;
const PUNCTUATION_WEIGHT_BASE = 0x3000;
const LETTER_WEIGHT = 0x4000;
const OTHER_WEIGHT = 0x5000;

function textCharacterWeight(character: string): number {
  const codePoint = character.codePointAt(0) ?? 0;
  if (codePoint >= 0x30 && codePoint <= 0x39) return DIGIT_WEIGHT + codePoint;
  if (codePoint === 0x20) return SPACE_WEIGHT;
  const punctuation = PUNCTUATION_WEIGHT.get(character);
  if (punctuation !== undefined) return PUNCTUATION_WEIGHT_BASE + punctuation;
  if (codePoint >= 0x41 && codePoint <= 0x5a) return LETTER_WEIGHT + codePoint;
  // ASCII lower case folds onto upper case; a locale-aware fold is v2.
  if (codePoint >= 0x61 && codePoint <= 0x7a) return LETTER_WEIGHT + codePoint - 0x20;
  // Anything else falls back to code-point order, which also keeps non-ASCII
  // after every ASCII letter.
  return OTHER_WEIGHT + codePoint;
}

function textFold(text: string): number[] {
  const weights: number[] = [];
  for (const character of text) {
    if (character === "'" || character === '-') continue;
    weights.push(textCharacterWeight(character));
  }
  return weights;
}

function countMarkers(text: string): number {
  let count = 0;
  for (const character of text) {
    if (character === "'" || character === '-') count += 1;
  }
  return count;
}

/**
 * Ascending text order. Case is ignored, and apostrophes and hyphens only
 * break a tie, where the decorated string sorts after the plain one.
 */
export function compareText(left: string, right: string): number {
  const leftWeights = textFold(left);
  const rightWeights = textFold(right);
  const shared = Math.min(leftWeights.length, rightWeights.length);
  for (let index = 0; index < shared; index += 1) {
    const difference = leftWeights[index]! - rightWeights[index]!;
    if (difference !== 0) return difference;
  }
  if (leftWeights.length !== rightWeights.length) {
    return leftWeights.length - rightWeights.length;
  }
  return countMarkers(left) - countMarkers(right);
}

/**
 * Spec ranks 1-4 as zero-based values. Blank is rank 5 and is handled before
 * any rank comparison so that it stays last in both directions.
 */
const NUMBER_RANK = 0;
const TEXT_RANK = 1;
const LOGICAL_RANK = 2;
const ERROR_RANK = 3;

function typeRank(value: CellValue): number {
  switch (value.type) {
    case 'number':
      return NUMBER_RANK;
    case 'text':
      return TEXT_RANK;
    case 'boolean':
      return LOGICAL_RANK;
    case 'error':
      return ERROR_RANK;
    case 'blank':
      return 4;
  }
}

/**
 * The §2.4 comparator. Blank cells sort last in both directions; descending
 * reverses ranks 1-4 only. Text compares with `compareText`, errors compare
 * equal, and numbers stored as text are text by their typed value.
 */
export function compareCellValues(
  left: CellValue,
  right: CellValue,
  direction: SortDirection,
): number {
  if (left.type === 'blank' || right.type === 'blank') {
    if (left.type === 'blank' && right.type === 'blank') return 0;
    return left.type === 'blank' ? 1 : -1;
  }
  const sign = direction === 'ascending' ? 1 : -1;
  const rankDifference = typeRank(left) - typeRank(right);
  if (rankDifference !== 0) return rankDifference * sign;
  if (left.type === 'number' && right.type === 'number') {
    if (left.value === right.value) return 0;
    return (left.value < right.value ? -1 : 1) * sign;
  }
  if (left.type === 'text' && right.type === 'text') {
    return compareText(left.value, right.value) * sign;
  }
  if (left.type === 'boolean' && right.type === 'boolean') {
    if (left.value === right.value) return 0;
    return (left.value ? 1 : -1) * sign;
  }
  return 0;
}

function compareKeySnapshots(
  left: readonly CellValue[],
  right: readonly CellValue[],
  keys: readonly SortKey[],
): number {
  for (let index = 0; index < keys.length; index += 1) {
    const difference = compareCellValues(left[index]!, right[index]!, keys[index]!.direction);
    if (difference !== 0) return difference;
  }
  return 0;
}

// ---------------------------------------------------------------------------
// Target range and header (§2.1, §2.2)
// ---------------------------------------------------------------------------

/** Normalise a selection so the start is the top-left corner. */
export function normaliseRange(range: GridRange): GridRange {
  return {
    startRow: Math.min(range.startRow, range.endRow),
    startCol: Math.min(range.startCol, range.endCol),
    endRow: Math.max(range.startRow, range.endRow),
    endCol: Math.max(range.startCol, range.endCol),
  };
}

/** A sheet index of the non-empty cells, so region growth never rescans the grid. */
interface RegionIndex {
  readonly rows: Map<number, number[]>;
  readonly columns: Map<number, number[]>;
}

function buildRegionIndex(sheet: Sheet): RegionIndex {
  const rows = new Map<number, number[]>();
  const columns = new Map<number, number[]>();
  for (const [key, cell] of sheet.cells) {
    if (cellIsEmpty(cell)) continue;
    const separator = key.indexOf(':');
    const row = Number(key.slice(0, separator));
    const col = Number(key.slice(separator + 1));
    appendTo(rows, row, col);
    appendTo(columns, col, row);
  }
  for (const list of rows.values()) list.sort((a, b) => a - b);
  for (const list of columns.values()) list.sort((a, b) => a - b);
  return { rows, columns };
}

function appendTo(index: Map<number, number[]>, key: number, value: number): void {
  const list = index.get(key);
  if (list === undefined) index.set(key, [value]);
  else list.push(value);
}

function lowerBound(sorted: readonly number[], value: number): number {
  let low = 0;
  let high = sorted.length;
  while (low < high) {
    const middle = (low + high) >>> 1;
    if (sorted[middle]! < value) low = middle + 1;
    else high = middle;
  }
  return low;
}

function hasValueInSpan(list: number[] | undefined, low: number, high: number): boolean {
  if (list === undefined) return false;
  const index = lowerBound(list, low);
  return index < list.length && list[index]! <= high;
}

/**
 * The current contiguous region around a cell: the largest block of adjacent
 * non-empty rows and columns that contains the anchor. An empty anchor has no
 * region, so it stays a single cell.
 */
export function currentRegion(sheet: Sheet, row: number, col: number): GridRange {
  const anchor = getCell(sheet, row, col);
  if (anchor === null || cellIsEmpty(anchor)) return singleCellRange(row, col);
  const index = buildRegionIndex(sheet);
  let startRow = row;
  let endRow = row;
  let startCol = col;
  let endCol = col;
  let grew = true;
  while (grew) {
    grew = false;
    if (startRow > 0 && hasValueInSpan(index.rows.get(startRow - 1), startCol, endCol)) {
      startRow -= 1;
      grew = true;
    }
    if (hasValueInSpan(index.rows.get(endRow + 1), startCol, endCol)) {
      endRow += 1;
      grew = true;
    }
    if (startCol > 0 && hasValueInSpan(index.columns.get(startCol - 1), startRow, endRow)) {
      startCol -= 1;
      grew = true;
    }
    if (hasValueInSpan(index.columns.get(endCol + 1), startRow, endRow)) {
      endCol += 1;
      grew = true;
    }
  }
  return { startRow, startCol, endRow, endCol };
}

/** The sort target: the selection, or the current region for a single cell. */
export function resolveTargetRange(sheet: Sheet, selection: GridRange): GridRange {
  const range = normaliseRange(selection);
  if (range.startRow === range.endRow && range.startCol === range.endCol) {
    return currentRegion(sheet, range.startRow, range.startCol);
  }
  return range;
}

function regionGrowsBeyond(region: GridRange, range: GridRange): boolean {
  return (
    region.startRow < range.startRow ||
    region.startCol < range.startCol ||
    region.endRow > range.endRow ||
    region.endCol > range.endCol
  );
}

/**
 * The topmost-leftmost non-empty cell of a range. A selection whose corners are
 * blank can still sit inside a block, so the sub-range warning needs a real
 * anchor rather than the selection's top-left corner.
 */
function firstContentInRange(sheet: Sheet, range: GridRange): CellAddress | null {
  let found: CellAddress | null = null;
  for (const [key, cell] of sheet.cells) {
    if (cellIsEmpty(cell)) continue;
    const separator = key.indexOf(':');
    const row = Number(key.slice(0, separator));
    const col = Number(key.slice(separator + 1));
    if (row < range.startRow || row > range.endRow) continue;
    if (col < range.startCol || col > range.endCol) continue;
    if (found === null || row < found.row || (row === found.row && col < found.col)) {
      found = { row, col };
    }
  }
  return found;
}

interface EmphasisFacts {
  readonly any: boolean;
  readonly key: string;
}

function emphasisOf(style: StyleRecord): EmphasisFacts {
  const bold = style.font.bold;
  const filled = style.fill.pattern !== 'none' || style.fill.color !== null;
  const bordered =
    style.border.left.style !== 'none' ||
    style.border.right.style !== 'none' ||
    style.border.top.style !== 'none' ||
    style.border.bottom.style !== 'none';
  return { any: bold || filled || bordered, key: `${bold}/${filled}/${bordered}` };
}

const NO_EMPHASIS: EmphasisFacts = { any: false, key: 'false/false/false' };

function headerHasDistinctStyle(workbook: Workbook, sheet: Sheet, range: GridRange): boolean {
  for (let col = range.startCol; col <= range.endCol; col += 1) {
    const headerCell = getCell(sheet, range.startRow, col);
    if (headerCell === null) continue;
    const header = emphasisOf(getStyle(workbook, headerCell.styleId));
    if (!header.any) continue;
    for (let row = range.startRow + 1; row <= range.endRow; row += 1) {
      const cell = getCell(sheet, row, col);
      const body = cell === null ? NO_EMPHASIS : emphasisOf(getStyle(workbook, cell.styleId));
      if (body.key !== header.key) return true;
    }
  }
  return false;
}

/**
 * The §2.2 header guess. The first row is a header when every cell in it is
 * text and either a key column holds a non-text value below or the row carries
 * a distinct style. A blank cell below a key label is not evidence either way,
 * because a gap in the data says nothing about the first row.
 */
export function guessHeaderRow(input: HeaderHeuristicInput): boolean {
  const { workbook, sheet, range } = input;
  if (range.endRow <= range.startRow) return false;
  for (let col = range.startCol; col <= range.endCol; col += 1) {
    const cell = getCell(sheet, range.startRow, col);
    if (cell === null || cell.value.type !== 'text') return false;
  }
  const keyColumns =
    input.keyColumns !== undefined && input.keyColumns.length > 0
      ? input.keyColumns
      : columnsOf(range);
  for (const column of keyColumns) {
    for (let row = range.startRow + 1; row <= range.endRow; row += 1) {
      const value = getCell(sheet, row, column)?.value ?? BLANK;
      if (value.type !== 'text' && value.type !== 'blank') return true;
    }
  }
  return headerHasDistinctStyle(workbook, sheet, range);
}

function columnsOf(range: GridRange): number[] {
  const columns: number[] = [];
  for (let col = range.startCol; col <= range.endCol; col += 1) columns.push(col);
  return columns;
}

function keyLabelsFor(
  sheet: Sheet,
  range: GridRange,
  keys: readonly SortKey[],
  hasHeaders: boolean,
): string[] {
  return keys.map((key) => {
    if (hasHeaders) {
      const cell = getCell(sheet, range.startRow, key.column);
      if (cell !== null && cell.value.type === 'text') return cell.value.value;
    }
    return columnToLetters(key.column);
  });
}

// ---------------------------------------------------------------------------
// What moves (§2.5)
// ---------------------------------------------------------------------------

function rangesIntersect(left: GridRange, right: GridRange): boolean {
  return (
    left.startRow <= right.endRow &&
    right.startRow <= left.endRow &&
    left.startCol <= right.endCol &&
    right.startCol <= left.endCol
  );
}

function rangeContainsRange(outer: GridRange, inner: GridRange): boolean {
  return (
    inner.startRow >= outer.startRow &&
    inner.endRow <= outer.endRow &&
    inner.startCol >= outer.startCol &&
    inner.endCol <= outer.endCol
  );
}

/**
 * A moved formula shifts every relative row part by the row delta. Absolute
 * row parts keep their anchor, whole columns have no row part, and a relative
 * part that leaves the grid becomes `#REF!`, as Excel does. Sheet prefixes and
 * defined names are untouched because only parsed areas are rewritten.
 */
export function shiftFormulaRows(formula: string, deltaRows: number): string {
  if (deltaRows === 0 || formula.length === 0) return formula;
  const references = findAreaReferences(formula);
  if (references.length === 0) return formula;
  let output = '';
  let cursor = 0;
  let changed = false;
  for (const reference of references) {
    output += formula.slice(cursor, reference.start);
    const shifted = shiftAreaRows(reference, deltaRows);
    if (shifted === null) {
      output += '#REF!';
      changed = true;
    } else if (shifted === reference) {
      output += formula.slice(reference.start, reference.end);
    } else {
      output += formatAreaReference(shifted);
      changed = true;
    }
    cursor = reference.end;
  }
  return changed ? output + formula.slice(cursor) : formula;
}

function shiftAreaRows(area: AreaReference, deltaRows: number): AreaReference | null {
  if (area.kind === 'columns') return area;
  const startRow = area.startRowAbs ? area.startRow : area.startRow + deltaRows;
  const endRow = area.endRowAbs ? area.endRow : area.endRow + deltaRows;
  if (startRow < 0 || endRow < 0 || startRow >= MAX_ROW || endRow >= MAX_ROW) return null;
  if (startRow === area.startRow && endRow === area.endRow) return area;
  return { ...area, startRow, endRow };
}

function moveCellToRow(cell: Cell, deltaRows: number): Cell {
  if (deltaRows === 0) return cell;
  // Shared groups keep their authored text and storage metadata, exactly as the
  // M3C structural transform treats them; an array group blocks the sort.
  if (cell.formula === null || cell.formulaSource !== undefined) return cell;
  const shifted = shiftFormulaRows(cell.formula, deltaRows);
  if (shifted === cell.formula) return cell;
  return { ...cell, formula: shifted, stale: true };
}

/**
 * Every merge that touches the range must be one row tall and fully inside it;
 * anything else blocks with Excel's message. The containment rule matches
 * `arrayFormulaForCell` (formula-ranges.ts), but a sort asks the whole-range
 * question once instead of per cell.
 */
function collectRowMerges(
  sheet: Sheet,
  range: GridRange,
): { readonly kind: 'blocked' } | { readonly kind: 'ok'; readonly merges: GridRange[] } {
  const merges: GridRange[] = [];
  for (const merge of sheet.merges) {
    if (!rangesIntersect(merge.range, range)) continue;
    if (!rangeContainsRange(range, merge.range)) return { kind: 'blocked' };
    if (merge.range.startRow !== merge.range.endRow) return { kind: 'blocked' };
    merges.push(merge.range);
  }
  return { kind: 'ok', merges };
}

function findArrayBlock(sheet: Sheet, range: GridRange): GridRange | null {
  for (const cell of sheet.cells.values()) {
    const source = cell.formulaSource;
    if (source?.kind !== 'array' || source.range === undefined) continue;
    const covered = parseArrayRef(source.range);
    if (covered !== null && rangesIntersect(covered, range)) return covered;
  }
  return null;
}

function parseArrayRef(ref: string): GridRange | null {
  try {
    return parseRangeRef(ref.replaceAll('$', ''));
  } catch {
    // A malformed ref stays unmodelled; the worksheet patch preserves it.
    return null;
  }
}

function attachmentMoves<T>(
  records: Map<string, T>,
  range: GridRange,
  destinationOf: ReadonlyMap<number, number>,
): SortAttachmentMove<T>[] {
  if (destinationOf.size === 0) return [];
  const moves: SortAttachmentMove<T>[] = [];
  for (const [key, record] of records) {
    const separator = key.indexOf(':');
    const row = Number(key.slice(0, separator));
    const col = Number(key.slice(separator + 1));
    if (row < range.startRow || row > range.endRow) continue;
    if (col < range.startCol || col > range.endCol) continue;
    const destination = destinationOf.get(row);
    if (destination === undefined) continue;
    moves.push({ from: { row, col }, to: { row: destination, col }, record });
  }
  return moves;
}

// ---------------------------------------------------------------------------
// The plan (§2.3, §2.6, §2.7)
// ---------------------------------------------------------------------------

/**
 * Build the sort plan. Nothing is mutated: the caller applies the plan in one
 * transaction, and the plan carries the permutation, the formula rewrites and
 * the moved cell records that undo needs.
 */
export function planSort(request: SortRequest): SortResult {
  const sheet = getSheetById(request.workbook, request.sheetId);
  const range = resolveTargetRange(sheet, request.selection);
  const warnings: SortWarning[] = [];
  if (!isSingleCell(request.selection)) {
    const anchor = firstContentInRange(sheet, range);
    if (anchor !== null) {
      const region = currentRegion(sheet, anchor.row, anchor.col);
      if (regionGrowsBeyond(region, range)) {
        warnings.push({ code: 'sub-range', message: SUB_RANGE_WARNING });
      }
    }
  }

  if (request.keys.length === 0) {
    return { kind: 'unchanged', reason: 'no-keys', range, warnings };
  }
  if (request.keys.length > MAX_SORT_KEYS) {
    return { kind: 'blocked', code: 'too-many-keys', message: TOO_MANY_KEYS_MESSAGE, range };
  }
  for (const key of request.keys) {
    const inside = Number.isInteger(key.column) && key.column >= range.startCol;
    if (!inside || key.column > range.endCol) {
      return {
        kind: 'blocked',
        code: 'key-outside-range',
        message: KEY_OUTSIDE_RANGE_MESSAGE,
        range,
      };
    }
  }

  const hasHeaders =
    request.hasHeaders ??
    guessHeaderRow({
      workbook: request.workbook,
      sheet,
      range,
      keyColumns: request.keys.map((key) => key.column),
    });
  const firstDataRow = range.startRow + (hasHeaders ? 1 : 0);
  if (range.endRow - firstDataRow + 1 < 2) {
    return { kind: 'unchanged', reason: 'too-few-rows', range, warnings };
  }

  const mergePlan = collectRowMerges(sheet, range);
  if (mergePlan.kind === 'blocked') {
    return { kind: 'blocked', code: 'merged-cells', message: MERGED_SIZE_MESSAGE, range };
  }
  const table = (request.tableRanges ?? []).find((candidate) => rangesIntersect(candidate, range));
  if (table !== undefined) {
    return { kind: 'blocked', code: 'table', message: TABLE_BLOCK_MESSAGE, range };
  }
  if (findArrayBlock(sheet, range) !== null) {
    return { kind: 'blocked', code: 'array-formula', message: ARRAY_BLOCK_MESSAGE, range };
  }

  const hiddenRows = new Set<number>();
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    if (sheet.rowDimensions.get(row)?.hidden === true) hiddenRows.add(row);
  }
  for (const row of request.filterHiddenRows ?? []) {
    if (row >= range.startRow && row <= range.endRow) hiddenRows.add(row);
  }

  const visibleRows: number[] = [];
  let visibleHiddenCount = 0;
  for (let row = firstDataRow; row <= range.endRow; row += 1) {
    if (hiddenRows.has(row)) visibleHiddenCount += 1;
    else visibleRows.push(row);
  }
  if (visibleRows.length < 2) {
    return { kind: 'unchanged', reason: 'nothing-to-order', range, warnings };
  }
  if (visibleHiddenCount > 0) {
    warnings.push({
      code: 'hidden-rows',
      message: HIDDEN_ROWS_WARNING,
      count: visibleHiddenCount,
    });
  }

  // The key values are read once, before any row moves (§2.3).
  const snapshots = visibleRows.map((row) =>
    request.keys.map((key) => getCell(sheet, row, key.column)?.value ?? BLANK),
  );
  const order = visibleRows.map((_row, index) => index);
  order.sort(
    (left, right) =>
      compareKeySnapshots(snapshots[left]!, snapshots[right]!, request.keys) || left - right,
  );

  const rowPermutation: number[] = [];
  for (let row = range.startRow; row <= range.endRow; row += 1) rowPermutation.push(row);
  const destinationOf = new Map<number, number>();
  visibleRows.forEach((destinationRow, slot) => {
    const sourceRow = visibleRows[order[slot]!]!;
    rowPermutation[destinationRow - range.startRow] = sourceRow;
    if (sourceRow !== destinationRow) destinationOf.set(sourceRow, destinationRow);
  });

  const rowMoves: SortRowMove[] = [];
  const cellRecords: SortCellRecord[] = [];
  const formulaRewrites: SortFormulaRewrite[] = [];
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    const sourceRow = rowPermutation[row - range.startRow]!;
    if (sourceRow === row) continue;
    rowMoves.push({ fromRow: sourceRow, toRow: row });
    const deltaRows = row - sourceRow;
    for (let col = range.startCol; col <= range.endCol; col += 1) {
      const before = getCell(sheet, row, col);
      const source = getCell(sheet, sourceRow, col);
      const after = source === null ? null : moveCellToRow(source, deltaRows);
      if (before !== null || after !== null) {
        cellRecords.push({
          address: { row, col },
          source: source === null ? null : { row: sourceRow, col },
          before,
          after,
        });
      }
      if (source?.formula != null && after?.formula !== source.formula) {
        formulaRewrites.push({
          address: { row, col },
          before: source.formula,
          after: after?.formula ?? source.formula,
        });
      }
    }
  }

  const mergeRewrites: SortMergeRewrite[] = [];
  for (const merge of mergePlan.merges) {
    const destination = destinationOf.get(merge.startRow);
    if (destination === undefined) continue;
    mergeRewrites.push({
      before: merge,
      after: { ...merge, startRow: destination, endRow: destination },
    });
  }

  const plan: SortPlan = {
    sheetId: sheet.id,
    range,
    hasHeaders,
    keyLabels: keyLabelsFor(sheet, range, request.keys, hasHeaders),
    keys: request.keys.map((key) => ({ column: key.column, direction: key.direction })),
    rowPermutation,
    rowMoves,
    cellRecords,
    formulaRewrites,
    mergeRewrites,
    commentMoves: attachmentMoves(sheet.comments, range, destinationOf),
    hyperlinkMoves: attachmentMoves(sheet.hyperlinks, range, destinationOf),
    visibleRowCount: visibleRows.length,
    hiddenRowCount: visibleHiddenCount,
    crossesHiddenRows: visibleHiddenCount > 0,
  };
  return { kind: 'sorted', plan, warnings };
}

function isSingleCell(range: GridRange): boolean {
  return range.startRow === range.endRow && range.startCol === range.endCol;
}
