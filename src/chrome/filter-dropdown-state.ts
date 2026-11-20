import { evaluateAutoFilter } from '@model/autofilter';
import type { ChecklistEntry, ColumnFilter, Condition, ValueChecklist } from '@model/autofilter';
import { datePartsToSerial, serialToDateParts } from '@model/date-system';
import type { AutoFilter, DateOperator, NumberOperator, Sheet, TextOperator } from '@model/types';
import { cellValueEquals, cloneCellValue } from '@model/value';

/**
 * The AutoFilter dropdown's state (sort-autofilter §3.2, §3.8; design-refs
 * m5-menus-dialogs §2). The checklist selection, the condition drafts and the
 * status-bar record count are pure data, so they are testable without a DOM.
 *
 * The rules the model owns — grouping, ordering and formatting — arrive as the
 * `ValueChecklist` from `collectValueChecklist`; this module only decides which
 * entries are ticked and what criterion Apply dispatches.
 */

export type FilterColumnType = 'text' | 'number' | 'date';

/**
 * The value checklist is virtualised so a column with thousands of distinct
 * values does not build a row for every entry. These numbers must match the
 * row height the dropdown sets inline and the list's CSS `max-height` in
 * `styles.css`.
 */
export const FILTER_ROW_HEIGHT = 21;

/** The value list's CSS `max-height`, in CSS pixels. */
export const FILTER_LIST_HEIGHT = 210;

/** Rows drawn outside the viewport so a scroll does not reveal blank space. */
export const FILTER_ROW_OVERSCAN = 4;

/** Which slice of the checklist rows to render, and the spacer heights around it. */
export interface FilterListWindow {
  /** First rendered row slot, inclusive. */
  readonly start: number;
  /** Last rendered row slot, exclusive. */
  readonly end: number;
  /** Total row slots: every matching value plus `(Blanks)` when offered. */
  readonly total: number;
  /** Height of the spacer above the rendered rows. */
  readonly leading: number;
  /** Height of the spacer below the rendered rows. */
  readonly trailing: number;
}

/**
 * The virtual window for a fixed-height row list. The top is clamped to the
 * scrollable range so a stale scroll position after the search box narrows
 * the list still renders the first rows rather than nothing.
 */
export function filterListWindow(
  totalRows: number,
  scrollTop: number,
  viewportHeight: number = FILTER_LIST_HEIGHT,
  rowHeight: number = FILTER_ROW_HEIGHT,
  overscan: number = FILTER_ROW_OVERSCAN,
): FilterListWindow {
  const total = Math.max(0, Math.trunc(totalRows));
  if (total === 0 || rowHeight <= 0 || viewportHeight <= 0) {
    return { start: 0, end: 0, total, leading: 0, trailing: 0 };
  }
  const maxScroll = Math.max(0, total * rowHeight - viewportHeight);
  const top = Math.min(Math.max(0, scrollTop), maxScroll);
  const first = Math.floor(top / rowHeight);
  const visible = Math.ceil(viewportHeight / rowHeight) + 1;
  const start = Math.max(0, first - Math.max(0, Math.trunc(overscan)));
  const end = Math.min(total, first + visible + Math.max(0, Math.trunc(overscan)));
  return {
    start,
    end,
    total,
    leading: start * rowHeight,
    trailing: (total - end) * rowHeight,
  };
}

/** Which checklist entries the user has ticked, by index into the full list. */
export interface FilterChecklistState {
  readonly checked: ReadonlySet<number>;
  readonly blanksChecked: boolean;
  /** `false` when the column has no blank data row, so `(Blanks)` is not shown. */
  readonly hasBlanks: boolean;
}

/**
 * A fresh value criterion ticks every entry. An existing value criterion ticks
 * the entries its selected values display as; a conditions or unsupported
 * criterion leaves the checklist fully ticked, because it does not describe a
 * value subset.
 */
export function initialChecklistState(
  entries: readonly ChecklistEntry[],
  hasBlanks: boolean,
  criterion: ColumnFilter | null,
): FilterChecklistState {
  if (criterion !== null && criterion.kind === 'values') {
    const checked = new Set<number>();
    entries.forEach((entry, index) => {
      if (
        entry.values.some((value) =>
          criterion.selected.some((selected) => cellValueEquals(selected, value)),
        )
      ) {
        checked.add(index);
      }
    });
    return { checked, blanksChecked: criterion.includeBlanks, hasBlanks };
  }
  const checked = new Set<number>();
  entries.forEach((_, index) => checked.add(index));
  return { checked, blanksChecked: hasBlanks, hasBlanks };
}

export function toggleChecklistEntry(
  state: FilterChecklistState,
  index: number,
): FilterChecklistState {
  const checked = new Set(state.checked);
  if (checked.has(index)) checked.delete(index);
  else checked.add(index);
  return { ...state, checked };
}

export function toggleChecklistBlanks(state: FilterChecklistState): FilterChecklistState {
  return { ...state, blanksChecked: !state.blanksChecked };
}

/** Select All is checked when every entry and `(Blanks)` are, and mixed when some are. */
export function selectAllState(
  state: FilterChecklistState,
  entries: readonly ChecklistEntry[],
): { readonly all: boolean; readonly some: boolean } {
  const all = state.checked.size === entries.length && (!state.hasBlanks || state.blanksChecked);
  const some = state.checked.size > 0 || (state.hasBlanks && state.blanksChecked);
  return { all, some: all ? false : some };
}

/** Select All ticks everything, or clears everything when it was fully ticked. */
export function toggleSelectAll(
  state: FilterChecklistState,
  entries: readonly ChecklistEntry[],
): FilterChecklistState {
  const { all } = selectAllState(state, entries);
  if (all) return { ...state, checked: new Set(), blanksChecked: false };
  const checked = new Set<number>();
  entries.forEach((_, index) => checked.add(index));
  return { ...state, checked, blanksChecked: state.hasBlanks };
}

/**
 * The criterion Apply dispatches, or `null` when every value is ticked — the
 * column then has no filter, so the caller clears it instead (Excel's
 * behaviour, and one undo step either way, §3.9).
 */
export function criterionFromChecklist(
  state: FilterChecklistState,
  entries: readonly ChecklistEntry[],
): ColumnFilter | null {
  const { all } = selectAllState(state, entries);
  if (all) return null;
  const selected = [...state.checked]
    .sort((left, right) => left - right)
    .flatMap((index) => entries[index]?.values ?? [])
    .map(cloneCellValue);
  return {
    kind: 'values',
    selected,
    includeBlanks: state.hasBlanks && state.blanksChecked,
  };
}

/** Nothing ticked would hide every row, so Apply refuses it (§3.2). */
export function checklistProblem(
  state: FilterChecklistState,
  entries: readonly ChecklistEntry[],
): string | null {
  if (state.checked.size === 0 && !(state.hasBlanks && state.blanksChecked)) {
    return 'Select at least one value or (Blanks) before applying the filter.';
  }
  if (entries.length === 0 && !state.hasBlanks) {
    return 'This column has no values to filter on.';
  }
  return null;
}

/**
 * Excel's type-dependent "Text/Number/Date Filters" menu. A column of numbers
 * carrying a date number format gets the date operators; anything mixed or
 * textual gets the text operators, whose conditions apply to text cells only.
 */
export function detectFilterColumnType(
  checklist: ValueChecklist,
  isDateFormatted: boolean,
): FilterColumnType {
  let sawNumber = false;
  let sawOther = false;
  for (const entry of checklist.entries) {
    const value = entry.values[0];
    if (value === undefined) continue;
    if (value.type === 'number') sawNumber = true;
    else sawOther = true;
  }
  if (sawOther || !sawNumber) return 'text';
  return isDateFormatted ? 'date' : 'number';
}

/**
 * A conservative date-format test for the filter-type menu. Quoted literals,
 * escaped characters, `_`/`*` padding and bracketed colour/locale sections are
 * removed first, so `0.00 "m"` and `[Red]0.00` stay numbers.
 */
export function isDateFormatCode(code: string | null): boolean {
  if (code === null) return false;
  const stripped = code
    .replace(/"[^"]*"/g, '')
    .replace(/\\./g, '')
    .replace(/\[[^\]]*\]/g, '')
    .replace(/_.|\*./g, '')
    .toLowerCase();
  return /[ymdhs]|am\/pm|a\/p/.test(stripped);
}

export interface FilterOperatorOption {
  readonly value: string;
  readonly label: string;
}

/** §3.2's text operators. */
export const TEXT_OPERATORS: readonly FilterOperatorOption[] = [
  { value: 'equals', label: 'equals' },
  { value: 'not-equals', label: 'does not equal' },
  { value: 'begins-with', label: 'begins with' },
  { value: 'ends-with', label: 'ends with' },
  { value: 'contains', label: 'contains' },
  { value: 'does-not-contain', label: 'does not contain' },
];

/** §3.2's number operators, including the two ranges. */
export const NUMBER_OPERATORS: readonly FilterOperatorOption[] = [
  { value: '=', label: 'equals' },
  { value: '<>', label: 'does not equal' },
  { value: '>', label: 'is greater than' },
  { value: '>=', label: 'is greater than or equal to' },
  { value: '<', label: 'is less than' },
  { value: '<=', label: 'is less than or equal to' },
  { value: 'between', label: 'is between' },
  { value: 'not-between', label: 'is not between' },
];

/** §3.2's date operators, compared as serial numbers. */
export const DATE_OPERATORS: readonly FilterOperatorOption[] = [
  { value: 'equals', label: 'equals' },
  { value: 'not-equals', label: 'does not equal' },
  { value: 'before', label: 'is before' },
  { value: 'after', label: 'is after' },
  { value: 'between', label: 'is between' },
  { value: 'not-between', label: 'is not between' },
];

export function operatorsFor(type: FilterColumnType): readonly FilterOperatorOption[] {
  if (type === 'number') return NUMBER_OPERATORS;
  if (type === 'date') return DATE_OPERATORS;
  return TEXT_OPERATORS;
}

/** The type-dependent menu title the capture shows: "Text Filters", "Number Filters"… */
export function filterMenuLabel(type: FilterColumnType): string {
  if (type === 'number') return 'Number Filters';
  if (type === 'date') return 'Date Filters';
  return 'Text Filters';
}

export function operatorAcceptsRange(operator: string): boolean {
  return operator === 'between' || operator === 'not-between';
}

export interface ConditionDraft {
  readonly operator: string;
  readonly value: string;
  /** The upper bound for a range operator; unused otherwise. */
  readonly to: string;
}

export interface ConditionsDraft {
  readonly join: 'and' | 'or';
  readonly first: ConditionDraft;
  readonly second: ConditionDraft;
}

export function defaultConditionsDraft(type: FilterColumnType): ConditionsDraft {
  const operator = operatorsFor(type)[0]?.value ?? 'equals';
  const empty: ConditionDraft = { operator, value: '', to: '' };
  return { join: 'and', first: empty, second: empty };
}

/** Prefill the conditions window from an existing conditions criterion. */
export function conditionsDraftFrom(
  criterion: ColumnFilter | null,
  type: FilterColumnType,
  date1904: boolean,
): ConditionsDraft | null {
  if (criterion === null || criterion.kind !== 'conditions') return null;
  const draft = defaultConditionsDraft(type);
  const first = criterion.conditions[0];
  const second = criterion.conditions[1];
  return {
    join: criterion.join,
    first: first === undefined ? draft.first : conditionToDraft(first, date1904),
    second: second === undefined ? draft.second : conditionToDraft(second, date1904),
  };
}

function conditionToDraft(condition: Condition, date1904: boolean): ConditionDraft {
  switch (condition.kind) {
    case 'text':
      return { operator: condition.operator, value: condition.value, to: '' };
    case 'number':
      return 'from' in condition
        ? { operator: condition.operator, value: String(condition.from), to: String(condition.to) }
        : { operator: condition.operator, value: String(condition.value), to: '' };
    case 'date':
      return 'from' in condition
        ? {
            operator: condition.operator,
            value: formatDateInput(condition.from, date1904),
            to: formatDateInput(condition.to, date1904),
          }
        : {
            operator: condition.operator,
            value: formatDateInput(condition.value, date1904),
            to: '',
          };
  }
}

/** An ISO date for the date input; the serial is read in the workbook's system. */
export function formatDateInput(serial: number, date1904: boolean): string {
  const parts = serialToDateParts(serial, date1904);
  if (parts === null) return '';
  const month = String(parts.month).padStart(2, '0');
  const day = String(parts.day).padStart(2, '0');
  return `${String(parts.year).padStart(4, '0')}-${month}-${day}`;
}

export type ConditionsBuild =
  | {
      readonly ok: true;
      readonly join: 'and' | 'or';
      readonly conditions: readonly Condition[];
    }
  | { readonly ok: false; readonly message: string };

/**
 * Build the conditions criterion from the two-row window. A blank second row
 * is ignored, so one condition is the common case; two is the §3.2 limit.
 */
export function buildConditions(
  draft: ConditionsDraft,
  type: FilterColumnType,
  date1904: boolean,
): ConditionsBuild {
  const first = conditionFromDraft(draft.first, type, date1904);
  if (first.ok === false) return first;
  const conditions: Condition[] = [first.condition];
  if (!draftIsBlank(draft.second)) {
    const second = conditionFromDraft(draft.second, type, date1904);
    if (second.ok === false) return second;
    conditions.push(second.condition);
  }
  return { ok: true, join: draft.join, conditions };
}

function draftIsBlank(draft: ConditionDraft): boolean {
  return draft.value.trim() === '' && draft.to.trim() === '';
}

type ConditionBuild =
  | { readonly ok: true; readonly condition: Condition }
  | { readonly ok: false; readonly message: string };

function conditionFromDraft(
  draft: ConditionDraft,
  type: FilterColumnType,
  date1904: boolean,
): ConditionBuild {
  const allowed = operatorsFor(type).some((option) => option.value === draft.operator);
  if (!allowed) return { ok: false, message: 'Choose an operator for the filter condition.' };
  if (type === 'text') {
    if (draft.value.trim() === '') {
      return { ok: false, message: 'Enter a value for the filter condition.' };
    }
    return {
      ok: true,
      condition: {
        kind: 'text',
        operator: draft.operator as TextOperator,
        value: draft.value,
      },
    };
  }
  if (type === 'number') {
    return numberCondition(draft);
  }
  return dateCondition(draft, date1904);
}

function numberCondition(draft: ConditionDraft): ConditionBuild {
  const operator = draft.operator as NumberOperator;
  if (operatorAcceptsRange(operator)) {
    const from = parseNumber(draft.value);
    const to = parseNumber(draft.to);
    if (from === null || to === null) {
      return { ok: false, message: 'Enter both numbers for the filter range.' };
    }
    return {
      ok: true,
      condition: { kind: 'number', operator: operator as 'between' | 'not-between', from, to },
    };
  }
  const value = parseNumber(draft.value);
  if (value === null) return { ok: false, message: 'Enter a number for the filter condition.' };
  return {
    ok: true,
    condition: {
      kind: 'number',
      operator: operator as Exclude<NumberOperator, 'between' | 'not-between'>,
      value,
    },
  };
}

function dateCondition(draft: ConditionDraft, date1904: boolean): ConditionBuild {
  const operator = draft.operator as DateOperator;
  if (operatorAcceptsRange(operator)) {
    const from = parseDate(draft.value, date1904);
    const to = parseDate(draft.to, date1904);
    if (from === null || to === null) {
      return { ok: false, message: 'Enter both dates for the filter range.' };
    }
    return {
      ok: true,
      condition: { kind: 'date', operator: operator as 'between' | 'not-between', from, to },
    };
  }
  const value = parseDate(draft.value, date1904);
  if (value === null) return { ok: false, message: 'Enter a date for the filter condition.' };
  return {
    ok: true,
    condition: {
      kind: 'date',
      operator: operator as Exclude<DateOperator, 'between' | 'not-between'>,
      value,
    },
  };
}

function parseNumber(text: string): number | null {
  const trimmed = text.trim();
  if (trimmed === '') return null;
  const value = Number(trimmed);
  return Number.isFinite(value) ? value : null;
}

/** `YYYY-MM-DD` from the date input, converted in the workbook's date system. */
function parseDate(text: string, date1904: boolean): number | null {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(text.trim());
  if (match === null) return null;
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  if (month < 1 || month > 12 || day < 1 || day > 31) return null;
  return datePartsToSerial(year, month, day, date1904);
}

/**
 * The status bar's `X of Y records` (§3.8), computed from the same two
 * visibility layers the filter payload reports (§3.3). Deriving it from the
 * workbook keeps it right after undo, redo and a file open, not only directly
 * after a filter command.
 */
export function filterRecordCount(
  sheet: Sheet,
  filter: AutoFilter,
): { readonly visible: number; readonly total: number } {
  const evaluation = evaluateAutoFilter(sheet, filter);
  const manualHidden = new Set<number>();
  for (const [row, dimension] of sheet.rowDimensions) {
    if (dimension.hidden === true) manualHidden.add(row);
  }
  const total = Math.max(0, filter.range.endRow - filter.range.startRow);
  let visible = 0;
  for (let row = filter.range.startRow + 1; row <= filter.range.endRow; row += 1) {
    if (!manualHidden.has(row) && !evaluation.filterHidden.has(row)) visible += 1;
  }
  return { visible, total };
}
