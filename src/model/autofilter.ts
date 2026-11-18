import { cellKey } from './types';
import type {
  AutoFilter,
  Cell,
  ColumnFilter,
  ColumnOffset,
  Condition,
  ConditionsFilter,
  DateComparisonOperator,
  DateCondition,
  DateOperator,
  GridRange,
  NumberComparisonOperator,
  NumberCondition,
  NumberOperator,
  Sheet,
  SheetId,
  TextCondition,
  TextOperator,
  UnsupportedFilter,
  ValueFilter,
} from './types';
import { BLANK, cellValueEquals, cloneCellValue, formatCellValue } from './value';
import type { CellValue } from './value';

/**
 * AutoFilter state, criteria evaluation and row visibility
 * (sort and AutoFilter §3).
 *
 * The module is pure: it reads the workbook model, never the engine, the DOM,
 * the clock or the file system. Dates compare as serial numbers, so the
 * workbook's 1900/1904 flag cannot change a result.
 *
 * The model's typed values are `CellValue` (workbook-model §1); the contract's
 * `TypedValue` is the same union. Criteria are document state; hidden rows are
 * derived from them (§3.3). This module returns those derived sets and the
 * before/after record one command needs (§3.9); it does not dispatch commands.
 */

/**
 * The AutoFilter state model (sort-autofilter §3.1-§3.2) lives in the core
 * model (`./types`), where `Sheet.autoFilter` carries it; this module owns its
 * behaviour and re-exports the shapes for callers.
 */
export type {
  AutoFilter,
  ColumnFilter,
  ColumnOffset,
  Condition,
  ConditionsFilter,
  DateComparisonOperator,
  DateCondition,
  DateOperator,
  NumberComparisonOperator,
  NumberCondition,
  NumberOperator,
  TextCondition,
  TextOperator,
  UnsupportedFilter,
  ValueFilter,
};

/** A filter can hold at most two conditions per column (§3.2). */
export const MAX_CONDITIONS_PER_COLUMN = 2;

/**
 * Build an AutoFilter, copying the range and column entries. `raw` is kept by
 * reference: it is opaque payload the model never interprets.
 */
export function createAutoFilter(
  sheetId: SheetId,
  range: GridRange,
  columns: Iterable<readonly [ColumnOffset, ColumnFilter]> = [],
  raw?: unknown,
): AutoFilter {
  const columnMap = new Map<ColumnOffset, ColumnFilter>();
  for (const [offset, columnFilter] of columns) columnMap.set(offset, columnFilter);
  return {
    sheetId,
    range: { ...range },
    columns: columnMap,
    ...(raw === undefined ? {} : { raw }),
  };
}

/**
 * A deep-enough copy for a transaction snapshot: a new map, new condition
 * arrays and cloned typed values. `raw` stays by reference because it is
 * opaque payload that must be preserved untouched.
 */
export function cloneAutoFilter(filter: AutoFilter): AutoFilter {
  const columns = new Map<ColumnOffset, ColumnFilter>();
  for (const [offset, columnFilter] of filter.columns) {
    columns.set(offset, cloneColumnFilter(columnFilter));
  }
  return {
    sheetId: filter.sheetId,
    range: { ...filter.range },
    columns,
    ...(filter.raw === undefined ? {} : { raw: filter.raw }),
  };
}

function cloneColumnFilter(filter: ColumnFilter): ColumnFilter {
  switch (filter.kind) {
    case 'values':
      return {
        kind: 'values',
        selected: filter.selected.map(cloneCellValue),
        includeBlanks: filter.includeBlanks,
      };
    case 'conditions':
      return {
        kind: 'conditions',
        join: filter.join,
        conditions: filter.conditions.map((condition) => ({ ...condition })),
      };
    case 'unsupported':
      return { kind: 'unsupported', raw: filter.raw };
  }
}

/** `true` for the row the contract always treats as the header (§3.1). */
export function isHeaderRow(filter: AutoFilter, row: number): boolean {
  return row === filter.range.startRow;
}

/** The model column a relative offset addresses, or `null` outside the filter range. */
export function columnForOffset(filter: AutoFilter, offset: ColumnOffset): number | null {
  const column = filter.range.startCol + offset;
  return column >= filter.range.startCol && column <= filter.range.endCol ? column : null;
}

/**
 * The current typed value a filter compares. `cell.value` is the value the
 * grid displays: opening a workbook fills it from the saved result and the
 * calculation layer writes both fields after a recalculation. Filtering reads
 * it and never asks the engine for a new result (§3.2).
 */
function currentValueOf(cell: Cell | null): CellValue {
  return cell === null ? BLANK : cell.value;
}

// ---------------------------------------------------------------------------
// Value checklist (§3.2)
// ---------------------------------------------------------------------------

/** One checklist row. Values that format alike are selected together. */
export interface ChecklistEntry {
  readonly display: string;
  /** Every typed value that displays as `display`; selecting the entry selects all of them. */
  readonly values: readonly CellValue[];
}

export interface ValueChecklist {
  readonly columnOffset: ColumnOffset;
  readonly entries: readonly ChecklistEntry[];
  /** `true` when at least one data row is blank, which drives `(Blanks)`. */
  readonly hasBlanks: boolean;
  /** Data rows scanned; the header row is never included. */
  readonly rowCount: number;
}

export interface ChecklistOptions {
  /**
   * Formatted display text, from the caller's number-format seam. Defaults to
   * `formatCellValue`, which has no formatting.
   */
  readonly displayText?: (cell: Cell, value: CellValue) => string;
  /**
   * §2.4 comparator. The default is the local implementation below; the sort
   * lane can inject its canonical comparator once `sort.ts` lands.
   */
  readonly compare?: (left: CellValue, right: CellValue) => number;
}

/**
 * Distinct typed values in a column, ordered by the §2.4 comparator, blanks
 * reported separately. Grouping is by type and formatted display: two numeric
 * cells that both display `1.00` become one entry selecting both, while number
 * `10` and text `"10"` stay distinct entries because selection is type-aware
 * (§3.2). Errors are ordinary entries.
 */
export function collectValueChecklist(
  sheet: Sheet,
  filter: AutoFilter,
  columnOffset: ColumnOffset,
  options: ChecklistOptions = {},
): ValueChecklist {
  const column = columnForOffset(filter, columnOffset);
  const groups = new Map<string, { display: string; values: CellValue[] }>();
  let hasBlanks = false;
  let rowCount = 0;

  if (column !== null) {
    for (let row = filter.range.startRow + 1; row <= filter.range.endRow; row += 1) {
      rowCount += 1;
      const cell = sheet.cells.get(cellKey(row, column)) ?? null;
      const value = currentValueOf(cell);
      if (cell === null || value.type === 'blank') {
        hasBlanks = true;
        continue;
      }
      const display =
        options.displayText === undefined
          ? formatCellValue(value)
          : options.displayText(cell, value);
      const key = `${value.type}\u0000${display}`;
      const group = groups.get(key);
      if (group === undefined) {
        groups.set(key, { display, values: [value] });
      } else if (!group.values.some((existing) => cellValueEquals(existing, value))) {
        group.values.push(value);
      }
    }
  }

  const compare = options.compare ?? compareFilterValues;
  const entries = [...groups.values()].sort((left, right) => {
    const leftValue = left.values[0];
    const rightValue = right.values[0];
    if (leftValue === undefined || rightValue === undefined) return 0;
    return compare(leftValue, rightValue);
  });
  return { columnOffset, entries, hasBlanks, rowCount };
}

/** The search box: case-insensitive display-text substring, in checklist order. */
export function searchChecklist(
  entries: readonly ChecklistEntry[],
  query: string,
): ChecklistEntry[] {
  const needle = query.toLowerCase();
  if (needle === '') return [...entries];
  return entries.filter((entry) => entry.display.toLowerCase().includes(needle));
}

// ---------------------------------------------------------------------------
// Comparison (§2.4)
// ---------------------------------------------------------------------------

const TYPE_RANK: Record<CellValue['type'], number> = {
  number: 1,
  text: 2,
  boolean: 3,
  error: 4,
  blank: 5,
};

/**
 * The §2.4 ascending order: numbers, text, logicals, errors, blanks last.
 * Descending reverses ranks 1-4 but keeps blanks last, so a caller must handle
 * direction separately from this function.
 *
 * The sort lane owns the canonical comparator; this local implementation keeps
 * the checklist self-contained until `sort.ts` lands, and
 * `ChecklistOptions.compare` is the seam for consolidating them.
 */
export function compareFilterValues(left: CellValue, right: CellValue): number {
  const rankDifference = TYPE_RANK[left.type] - TYPE_RANK[right.type];
  if (rankDifference !== 0) return rankDifference;
  if (left.type === 'number' && right.type === 'number') {
    return left.value < right.value ? -1 : left.value > right.value ? 1 : 0;
  }
  if (left.type === 'text' && right.type === 'text') {
    return compareExcelText(left.value, right.value);
  }
  if (left.type === 'boolean' && right.type === 'boolean') {
    return left.value === right.value ? 0 : left.value ? 1 : -1;
  }
  // Errors all compare equal, so a stable sort keeps their original order;
  // two blanks are equal as well.
  return 0;
}

/**
 * English-locale text order: `0-9`, space, punctuation, `A-Z`, then non-ASCII
 * by code point. Case is ignored. Apostrophes and hyphens are ignored except
 * as a final tie-break.
 */
function compareExcelText(left: string, right: string): number {
  const normalisedLeft = left.toLowerCase();
  const normalisedRight = right.toLowerCase();
  const primary = compareExcelCharacters(
    stripCollationIgnored(normalisedLeft),
    stripCollationIgnored(normalisedRight),
  );
  if (primary !== 0) return primary;
  return compareExcelCharacters(normalisedLeft, normalisedRight);
}

function stripCollationIgnored(text: string): string {
  return text.replace(/['-]/g, '');
}

function compareExcelCharacters(left: string, right: string): number {
  const leftCharacters = Array.from(left);
  const rightCharacters = Array.from(right);
  const shared = Math.min(leftCharacters.length, rightCharacters.length);
  for (let index = 0; index < shared; index += 1) {
    const difference = compareExcelCharacter(leftCharacters[index]!, rightCharacters[index]!);
    if (difference !== 0) return difference;
  }
  return leftCharacters.length - rightCharacters.length;
}

/** Band 0: digits and control characters, 1: space, 2: punctuation, 3: letters, 4: non-ASCII. */
function compareExcelCharacter(left: string, right: string): number {
  const [leftBand, leftWeight] = characterOrder(left);
  const [rightBand, rightWeight] = characterOrder(right);
  return leftBand === rightBand ? leftWeight - rightWeight : leftBand - rightBand;
}

function characterOrder(character: string): [number, number] {
  const code = character.codePointAt(0) ?? 0;
  if (code < 32) return [0, code];
  if (code >= 48 && code <= 57) return [0, code];
  if (code === 32) return [1, code];
  if (code < 128) {
    if (code >= 65 && code <= 90) return [3, code];
    if (code >= 97 && code <= 122) return [3, code - 32];
    return [2, code];
  }
  return [4, code];
}

// ---------------------------------------------------------------------------
// Condition evaluation (§3.2)
// ---------------------------------------------------------------------------

/** `unsupported` is a typed outcome, never a thrown string (§3.7). */
export type CriterionMatch = 'match' | 'no-match' | 'unsupported';

/**
 * One condition against one typed value. Text conditions match text cells,
 * number and date conditions match numeric cells, and every other type —
 * including blanks and errors — matches nothing. Negated operators negate only
 * within the applicable type, so `does not equal "x"` still hides blanks and
 * errors.
 */
export function evaluateCondition(condition: Condition, value: CellValue): boolean {
  return compileCondition(condition)(value);
}

/** One column criterion against one typed value. */
export function evaluateColumnFilter(filter: ColumnFilter, value: CellValue): CriterionMatch {
  const matches = compileColumnFilter(filter);
  if (matches === null) return 'unsupported';
  return matches(value) ? 'match' : 'no-match';
}

type ValuePredicate = (value: CellValue) => boolean;

/** `null` means an unsupported criterion, which cannot be evaluated (§3.7). */
function compileColumnFilter(filter: ColumnFilter): ValuePredicate | null {
  switch (filter.kind) {
    case 'values':
      return (value) => matchesValueFilter(filter, value);
    case 'conditions': {
      const predicates = filter.conditions.map(compileCondition);
      if (predicates.length === 0) return () => true;
      return filter.join === 'and'
        ? (value) => predicates.every((predicate) => predicate(value))
        : (value) => predicates.some((predicate) => predicate(value));
    }
    case 'unsupported':
      return null;
  }
}

function matchesValueFilter(filter: ValueFilter, value: CellValue): boolean {
  if (value.type === 'blank') return filter.includeBlanks;
  return filter.selected.some((selected) => cellValueEquals(selected, value));
}

function compileCondition(condition: Condition): ValuePredicate {
  switch (condition.kind) {
    case 'text': {
      const matches = compileTextCondition(condition);
      return (value) => value.type === 'text' && matches(value.value);
    }
    case 'number': {
      const matches = compileNumberCondition(condition);
      return (value) => value.type === 'number' && matches(value.value);
    }
    case 'date': {
      const matches = compileDateCondition(condition);
      return (value) => value.type === 'number' && matches(value.value);
    }
  }
}

type TextMatcher = (text: string) => boolean;

function compileTextCondition(condition: TextCondition): TextMatcher {
  switch (condition.operator) {
    case 'equals': {
      const pattern = wildcardRegExp(condition.value, 'whole');
      return (text) => pattern.test(text);
    }
    case 'not-equals': {
      const pattern = wildcardRegExp(condition.value, 'whole');
      return (text) => !pattern.test(text);
    }
    case 'begins-with': {
      const pattern = wildcardRegExp(condition.value, 'prefix');
      return (text) => pattern.test(text);
    }
    case 'ends-with': {
      const pattern = wildcardRegExp(condition.value, 'suffix');
      return (text) => pattern.test(text);
    }
    case 'contains': {
      const pattern = wildcardRegExp(condition.value, 'anywhere');
      return (text) => pattern.test(text);
    }
    case 'does-not-contain': {
      const pattern = wildcardRegExp(condition.value, 'anywhere');
      return (text) => !pattern.test(text);
    }
  }
}

function compileNumberCondition(condition: NumberCondition): (value: number) => boolean {
  if ('from' in condition) {
    const lower = Math.min(condition.from, condition.to);
    const upper = Math.max(condition.from, condition.to);
    const inside = (value: number): boolean => value >= lower && value <= upper;
    return condition.operator === 'between' ? inside : (value) => value < lower || value > upper;
  }
  const operand = condition.value;
  switch (condition.operator) {
    case '=':
      return (value) => value === operand;
    case '<>':
      return (value) => value !== operand;
    case '>':
      return (value) => value > operand;
    case '>=':
      return (value) => value >= operand;
    case '<':
      return (value) => value < operand;
    case '<=':
      return (value) => value <= operand;
  }
}

function compileDateCondition(condition: DateCondition): (serial: number) => boolean {
  if ('from' in condition) {
    const lower = Math.min(condition.from, condition.to);
    const upper = Math.max(condition.from, condition.to);
    return condition.operator === 'between'
      ? (serial) => serial >= lower && serial <= upper
      : (serial) => serial < lower || serial > upper;
  }
  const operand = condition.value;
  switch (condition.operator) {
    case 'equals':
      return (serial) => serial === operand;
    case 'not-equals':
      return (serial) => serial !== operand;
    case 'before':
      return (serial) => serial < operand;
    case 'after':
      return (serial) => serial > operand;
  }
}

type WildcardAnchor = 'whole' | 'prefix' | 'suffix' | 'anywhere';

/**
 * Excel wildcards: `*` is any run, `?` is one character, and `~` escapes the
 * next character. Matching is case-insensitive and `.` spans newlines. These
 * mirror the engine shim's lookup wildcards; the engine's helpers are private
 * to that layer.
 */
function wildcardRegExp(pattern: string, anchor: WildcardAnchor): RegExp {
  let source = '';
  for (let index = 0; index < pattern.length; index += 1) {
    const character = pattern[index]!;
    if (character === '~' && index + 1 < pattern.length) {
      index += 1;
      source += escapeRegExp(pattern[index]!);
      continue;
    }
    if (character === '*') {
      source += '.*';
      continue;
    }
    if (character === '?') {
      source += '.';
      continue;
    }
    source += escapeRegExp(character);
  }
  if (anchor === 'whole' || anchor === 'prefix') source = `^${source}`;
  if (anchor === 'whole' || anchor === 'suffix') source = `${source}$`;
  return new RegExp(source, 'is');
}

function escapeRegExp(text: string): string {
  return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

// ---------------------------------------------------------------------------
// Whole-filter evaluation (§3.2)
// ---------------------------------------------------------------------------

export interface AutoFilterEvaluation {
  /** Data rows that fail at least one filtered column. The header is never here. */
  readonly filterHidden: ReadonlySet<number>;
  /**
   * Criteria calcplus cannot evaluate (§3.7). Their rows are not hidden by
   * this layer; imported hidden flags keep them hidden through the manual
   * layer, and the UI offers only Clear.
   */
  readonly unsupportedColumns: readonly ColumnOffset[];
  /** Number of data rows the filter range covers. */
  readonly rowsEvaluated: number;
}

/**
 * Evaluate every filtered column over the range's data rows. Columns combine
 * with AND, the header row is always excluded, and a cell's current value is
 * read without recalculation. Unsupported columns pass every row but are
 * reported so the caller can show the read-only notice.
 */
export function evaluateAutoFilter(sheet: Sheet, filter: AutoFilter): AutoFilterEvaluation {
  const columns: { column: number; matches: ValuePredicate }[] = [];
  const unsupportedColumns: ColumnOffset[] = [];

  for (const [offset, columnFilter] of filter.columns) {
    if (columnForOffset(filter, offset) === null) continue;
    const matches = compileColumnFilter(columnFilter);
    if (matches === null) {
      unsupportedColumns.push(offset);
      continue;
    }
    columns.push({ column: filter.range.startCol + offset, matches });
  }

  const filterHidden = new Set<number>();
  for (let row = filter.range.startRow + 1; row <= filter.range.endRow; row += 1) {
    for (const column of columns) {
      const cell = sheet.cells.get(cellKey(row, column.column)) ?? null;
      if (!column.matches(currentValueOf(cell))) {
        filterHidden.add(row);
        break;
      }
    }
  }

  const rowsEvaluated = Math.max(0, filter.range.endRow - filter.range.startRow);
  return { filterHidden, unsupportedColumns, rowsEvaluated };
}

// ---------------------------------------------------------------------------
// Visibility layers (§3.3)
// ---------------------------------------------------------------------------

const NO_HIDDEN_ROWS: ReadonlySet<number> = new Set<number>();

/**
 * The two independent layers and their union. `manualHidden` is the row's own
 * imported hidden flag; `filterHidden` is derived from the active criteria.
 */
export interface HiddenLayers {
  readonly manualHidden: ReadonlySet<number>;
  readonly filterHidden: ReadonlySet<number>;
  readonly effectiveHidden: ReadonlySet<number>;
}

/**
 * Derive both layers from criteria plus the manual layer. Applying or changing
 * a filter only replaces `filterHidden`; the manual layer is echoed unchanged.
 */
export function computeHiddenLayers(
  sheet: Sheet,
  filter: AutoFilter | null,
  manualHidden: Iterable<number> = [],
): HiddenLayers {
  const manual = new Set(manualHidden);
  const filterHidden =
    filter === null
      ? new Set<number>()
      : new Set<number>(evaluateAutoFilter(sheet, filter).filterHidden);
  const effective = new Set(manual);
  for (const row of filterHidden) effective.add(row);
  return { manualHidden: manual, filterHidden, effectiveHidden: effective };
}

/**
 * Rebuild the layers from a saved effective hidden set (§3.3): every hidden
 * row that fails the saved criteria is filter-hidden, and every other hidden
 * row is manual. Rows hidden by an unsupported criterion cannot be evaluated,
 * so they stay manual — which keeps them hidden after open and after Clear.
 */
export function restoreHiddenLayers(
  sheet: Sheet,
  filter: AutoFilter | null,
  savedHidden: Iterable<number> = [],
): HiddenLayers {
  const filterHidden =
    filter === null
      ? new Set<number>()
      : new Set<number>(evaluateAutoFilter(sheet, filter).filterHidden);
  const manual = new Set<number>();
  for (const row of savedHidden) {
    if (!filterHidden.has(row)) manual.add(row);
  }
  const effective = new Set(manual);
  for (const row of filterHidden) effective.add(row);
  return { manualHidden: manual, filterHidden, effectiveHidden: effective };
}

/**
 * Clearing removes only the filter layer. Manually hidden rows stay hidden
 * and every filter-hidden row becomes visible (§3.3).
 */
export function clearFilterLayers(layers: HiddenLayers): HiddenLayers {
  const manual = new Set(layers.manualHidden);
  return {
    manualHidden: manual,
    filterHidden: new Set<number>(),
    effectiveHidden: manual,
  };
}

/**
 * A row predicate that reads both layers. `manualHidden` is read live, so a
 * caller can pass a set it updates between sorts without rebuilding the
 * predicate. The predicate does not range-check: rows outside the filter
 * range can only be manually hidden.
 */
export function visibleRowPredicate(
  sheet: Sheet,
  filter: AutoFilter | null,
  manualHidden: ReadonlySet<number> = NO_HIDDEN_ROWS,
): (row: number) => boolean {
  const filterHidden =
    filter === null ? NO_HIDDEN_ROWS : evaluateAutoFilter(sheet, filter).filterHidden;
  return (row) => !manualHidden.has(row) && !filterHidden.has(row);
}

/** Every visible row of `range`, in document order. */
export function visibleRowsInRange(
  sheet: Sheet,
  filter: AutoFilter | null,
  manualHidden: ReadonlySet<number>,
  range: GridRange,
): number[] {
  const isVisible = visibleRowPredicate(sheet, filter, manualHidden);
  const rows: number[] = [];
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    if (isVisible(row)) rows.push(row);
  }
  return rows;
}

// ---------------------------------------------------------------------------
// Sort interplay (§3.4)
// ---------------------------------------------------------------------------

export interface FilteredSortPlan {
  /** The same criteria object; a sort never changes filter state (§3.4). */
  readonly filter: AutoFilter | null;
  /** Rows the sort may reorder, in document order, into their own visible slots. */
  readonly visibleRows: readonly number[];
  /** Rows hidden by either layer; they stay in their positions. */
  readonly fixedRows: readonly number[];
}

/**
 * Divide a sort target into rows that may move and rows that stay (§2.6,
 * §3.4). The values in `visibleRows` all matched before the sort, and sorting
 * only permutes them, so re-evaluating the same criteria afterwards yields the
 * same hidden set. The sort layer removes the header row when headers are on.
 */
export function planFilteredSort(
  sheet: Sheet,
  filter: AutoFilter | null,
  manualHidden: ReadonlySet<number>,
  targetRange: GridRange,
): FilteredSortPlan {
  const isVisible = visibleRowPredicate(sheet, filter, manualHidden);
  const visibleRows: number[] = [];
  const fixedRows: number[] = [];
  for (let row = targetRange.startRow; row <= targetRange.endRow; row += 1) {
    if (isVisible(row)) visibleRows.push(row);
    else fixedRows.push(row);
  }
  return { filter, visibleRows, fixedRows };
}

// ---------------------------------------------------------------------------
// Transaction record (§3.9)
// ---------------------------------------------------------------------------

/** One row whose effective hidden flag changes. Undo restores the before state. */
export interface HiddenRowChange {
  readonly row: number;
  readonly before: boolean;
  readonly after: boolean;
}

/**
 * What one AutoFilter transaction stores: the criteria before and after plus
 * the effective hidden-row changes (§3.9). Applying, changing and clearing are
 * each one such record; data only, so the command layer owns dispatch and undo.
 */
export interface AutoFilterTransaction {
  readonly sheetId: SheetId;
  readonly before: AutoFilter | null;
  readonly after: AutoFilter | null;
  readonly hiddenChanges: readonly HiddenRowChange[];
}

/**
 * Diff the effective hidden flags of two filter states against the same
 * manual layer. Only rows inside either filter range can change, and a
 * no-op reapply produces an empty change list.
 */
export function planAutoFilterChange(
  sheet: Sheet,
  before: AutoFilter | null,
  after: AutoFilter | null,
  manualHidden: ReadonlySet<number> = NO_HIDDEN_ROWS,
): AutoFilterTransaction {
  const beforeFilterHidden =
    before === null ? NO_HIDDEN_ROWS : evaluateAutoFilter(sheet, before).filterHidden;
  const afterFilterHidden =
    after === null ? NO_HIDDEN_ROWS : evaluateAutoFilter(sheet, after).filterHidden;

  const rows = new Set<number>();
  if (before !== null) addDataRows(rows, before.range);
  if (after !== null) addDataRows(rows, after.range);

  const hiddenChanges: HiddenRowChange[] = [];
  for (const row of [...rows].sort((left, right) => left - right)) {
    const hiddenBefore = manualHidden.has(row) || beforeFilterHidden.has(row);
    const hiddenAfter = manualHidden.has(row) || afterFilterHidden.has(row);
    if (hiddenBefore !== hiddenAfter)
      hiddenChanges.push({ row, before: hiddenBefore, after: hiddenAfter });
  }

  return { sheetId: sheet.id, before, after, hiddenChanges };
}

function addDataRows(rows: Set<number>, range: GridRange): void {
  for (let row = range.startRow + 1; row <= range.endRow; row += 1) rows.add(row);
}
