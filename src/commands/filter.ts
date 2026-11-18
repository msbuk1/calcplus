import {
  MAX_CONDITIONS_PER_COLUMN,
  cloneAutoFilter,
  evaluateAutoFilter,
  planAutoFilterChange,
} from '@model/autofilter';
import type { AutoFilter, ColumnFilter, ColumnOffset, Condition } from '@model/autofilter';
import { resolveTargetRange } from '@model/sort';
import type { GridRange, Sheet, SheetId } from '@model/types';
import { cellValueEquals } from '@model/value';
import type {
  CommandChange,
  CommandContext,
  CommandDefinition,
  CommandResult,
  FilterHiddenRowChange,
  FilterPayload,
} from './types';

/**
 * The AutoFilter commands (sort-autofilter §3): toggle, one column criterion,
 * clear one column, clear all criteria and reapply. Each change is one
 * transaction that stores the `autoFilter` before and after plus the effective
 * hidden-row changes (§3.9); undo restores the criteria and the hidden rows
 * together because filter-hidden rows are derived from the criteria (§3.3).
 *
 * The commands never touch view state and draw no dropdown: the checklist, the
 * condition builder and the funnel icon read these outcomes from their own
 * layer.
 */

/** Enablement shared by the filter commands: a filter exists to act on (§3.8). */
function hasFilter(context: CommandContext): boolean {
  const active =
    context.activeSheetId === null
      ? undefined
      : context.workbook.sheets.find((sheet) => sheet.id === context.activeSheetId);
  if (active !== undefined) return active.autoFilter !== null;
  return context.workbook.sheets.some((sheet) => sheet.autoFilter !== null);
}

/** Every filter command needs a worksheet; `hasFilter` adds the criteria check. */
function hasSheet(context: CommandContext): boolean {
  return context.workbook.sheets.length > 0;
}

// ---------------------------------------------------------------------------
// Toggle (Ctrl+Shift+L, §3.1, §3.8)
// ---------------------------------------------------------------------------

export interface FilterToggleArgs {
  readonly sheetId?: SheetId;
  /** The range to attach the filter to; defaults to the selection (§3.1). */
  readonly range?: GridRange;
  /** `true` or `false` sets the state; omitted toggles it. */
  readonly on?: boolean;
}

export const filterToggleCommand: CommandDefinition<FilterToggleArgs> = {
  id: 'data.filter',
  label: 'Filter',
  category: 'data',
  keybindings: ['Ctrl+Shift+L'],
  isEnabled: hasSheet,
  isChecked(context) {
    const active =
      context.activeSheetId === null
        ? undefined
        : context.workbook.sheets.find((sheet) => sheet.id === context.activeSheetId);
    return active !== undefined && active.autoFilter !== null;
  },
  run(context, args = {}): CommandResult {
    const sheet = sheetFor(context, args.sheetId);
    if (sheet === undefined) return toast('Open a worksheet before using AutoFilter.');
    const before = sheet.autoFilter === null ? null : cloneAutoFilter(sheet.autoFilter);
    // Ctrl+Shift+L passes no range and toggles. An explicit range is a request
    // to filter that range, so it turns the filter on unless `on` says otherwise.
    const wantsOn = args.on ?? (args.range !== undefined ? true : before === null);

    if (!wantsOn) {
      // Toggling off removes only the filter layer. Manual hidden rows are the
      // row's own flag and stay hidden (§3.3).
      if (before === null) {
        return {
          changes: [],
          filter: filterPayload(sheet, null, null, [], 'none', false, manualHiddenRows(sheet)),
        };
      }
      return filterResult(sheet, before, null, 'cleared');
    }

    const selection = args.range ?? context.selection ?? rangeOfActiveCell(context);
    if (selection === null) return toast('Select a cell or range for AutoFilter first.');
    const range = resolveTargetRange(sheet, selection);
    if (range.endRow - range.startRow < 1) {
      return toast('Select a header row and at least one data row before turning on AutoFilter.');
    }
    if (before !== null && sameRange(before.range, range)) {
      // Already on for this range: the criteria stay and nothing is logged.
      return filterResult(sheet, before, cloneAutoFilter(before), 'applied');
    }
    // One AutoFilter per sheet (§3.1): enabling on another range replaces the
    // old filter, criteria included, in this one transaction (fixture F16).
    return filterResult(sheet, before, emptyFilter(sheet.id, range), 'applied');
  },
};

// ---------------------------------------------------------------------------
// One column criterion (§3.2, §3.7)
// ---------------------------------------------------------------------------

export interface FilterColumnArgs {
  readonly sheetId?: SheetId;
  /** The absolute column index of a header cell inside the filter range. */
  readonly column: number;
  /** The checklist or the two conditions; `null` clears this column. */
  readonly criterion: ColumnFilter | null;
  /** Replacing an imported criterion calcplus cannot edit asks first (§3.7). */
  readonly confirm?: boolean;
}

export const filterColumnCommand: CommandDefinition<FilterColumnArgs> = {
  id: 'data.filterColumn',
  label: 'Filter by column',
  category: 'data',
  keybindings: [],
  isEnabled: hasFilter,
  run(context, args): CommandResult {
    return applyColumnCriterion(context, args.sheetId, args.column, args.criterion, args.confirm);
  },
};

export interface FilterClearColumnArgs {
  readonly sheetId?: SheetId;
  readonly column: number;
}

export const filterClearColumnCommand: CommandDefinition<FilterClearColumnArgs> = {
  id: 'data.filterClearColumn',
  label: 'Clear column filter',
  category: 'data',
  keybindings: [],
  isEnabled: hasFilter,
  run(context, args): CommandResult {
    return applyColumnCriterion(context, args.sheetId, args.column, null, undefined);
  },
};

function applyColumnCriterion(
  context: CommandContext,
  sheetId: SheetId | undefined,
  column: number,
  criterion: ColumnFilter | null,
  confirm: boolean | undefined,
): CommandResult {
  const sheet = sheetFor(context, sheetId);
  if (sheet === undefined) return toast('Open a worksheet before filtering.');
  const filter = sheet.autoFilter;
  if (filter === null) return toast('Turn on AutoFilter before filtering a column.');
  if (!Number.isInteger(column) || column < filter.range.startCol || column > filter.range.endCol) {
    return toast('That column is outside the AutoFilter range.');
  }
  const offset = column - filter.range.startCol;
  const before = cloneAutoFilter(filter);
  const existing = filter.columns.get(offset);

  if (criterion === null) {
    // Clearing one column is its own step; a column with no criterion is a no-op.
    return filterResult(sheet, before, withoutColumn(filter, offset), 'cleared');
  }
  const problem = criterionProblem(criterion);
  if (problem !== null) return toast(problem);
  if (existing?.kind === 'unsupported' && confirm !== true) {
    return {
      changes: [],
      ui: {
        type: 'dialog',
        message: 'This filter uses a feature calcplus cannot edit. Replace it with the new filter?',
      },
    };
  }
  return filterResult(sheet, before, withColumn(filter, offset, criterion), 'changed');
}

// ---------------------------------------------------------------------------
// Clear All and Reapply (§3.8)
// ---------------------------------------------------------------------------

export interface FilterClearArgs {
  readonly sheetId?: SheetId;
}

/**
 * Clear All keeps the AutoFilter and its dropdowns and drops every criterion,
 * which restores the filter-hidden rows and leaves manual ones hidden (§3.3).
 * Turning the filter itself off is the toggle.
 */
export const filterClearCommand: CommandDefinition<FilterClearArgs> = {
  id: 'data.filterClear',
  label: 'Clear filter',
  category: 'data',
  keybindings: [],
  isEnabled: hasFilter,
  run(context, args = {}): CommandResult {
    const sheet = sheetFor(context, args.sheetId);
    if (sheet === undefined) return toast('Open a worksheet before clearing a filter.');
    const filter = sheet.autoFilter;
    if (filter === null) return toast('There is no AutoFilter to clear.');
    const before = cloneAutoFilter(filter);
    const after = cloneAutoFilter(filter);
    after.columns.clear();
    return filterResult(sheet, before, after, 'cleared');
  },
};

export interface FilterReapplyArgs {
  readonly sheetId?: SheetId;
}

/**
 * Reapply recomputes the filter layer. calcplus derives filter-hidden rows from
 * the criteria (§3.3), so with unchanged criteria there is no stored state to
 * rewrite: the transaction is the empty one §3.9 describes.
 */
export const filterReapplyCommand: CommandDefinition<FilterReapplyArgs> = {
  id: 'data.filterReapply',
  label: 'Reapply filter',
  category: 'data',
  keybindings: [],
  isEnabled: hasFilter,
  run(context, args = {}): CommandResult {
    const sheet = sheetFor(context, args.sheetId);
    if (sheet === undefined) return toast('Open a worksheet before reapplying a filter.');
    const filter = sheet.autoFilter;
    if (filter === null) return toast('There is no AutoFilter to reapply.');
    return filterResult(sheet, cloneAutoFilter(filter), cloneAutoFilter(filter), 'reapplied');
  },
};

// ---------------------------------------------------------------------------
// The transaction (§3.9)
// ---------------------------------------------------------------------------

function filterResult(
  sheet: Sheet,
  before: AutoFilter | null,
  after: AutoFilter | null,
  state: FilterPayload['state'],
): CommandResult {
  const manualHidden = manualHiddenRows(sheet);
  const record = planAutoFilterChange(sheet, before, after, manualHidden);
  const changed = !autoFilterStateEquals(before, after) || record.hiddenChanges.length > 0;
  const payload = filterPayload(
    sheet,
    before,
    after,
    record.hiddenChanges,
    state,
    changed,
    manualHidden,
  );
  if (!changed) return { changes: [], filter: payload };
  const range = (after ?? before)?.range;
  if (range === undefined) return { changes: [], filter: payload };
  const change: CommandChange = {
    type: 'filter.set',
    sheetId: sheet.id,
    range: { ...range },
    before,
    after,
    hiddenChanges: record.hiddenChanges,
  };
  return { changes: [change], filter: payload };
}

function filterPayload(
  sheet: Sheet,
  before: AutoFilter | null,
  after: AutoFilter | null,
  hiddenChanges: readonly FilterHiddenRowChange[],
  state: FilterPayload['state'],
  changed: boolean,
  manualHidden: ReadonlySet<number>,
): FilterPayload {
  const evaluation = after === null ? null : evaluateAutoFilter(sheet, after);
  const filterHidden = evaluation === null ? new Set<number>() : new Set(evaluation.filterHidden);
  const range = (after ?? before)?.range ?? null;
  let rowsVisible = 0;
  let rowsTotal = 0;
  if (after !== null) {
    rowsTotal = Math.max(0, after.range.endRow - after.range.startRow);
    for (let row = after.range.startRow + 1; row <= after.range.endRow; row += 1) {
      if (!manualHidden.has(row) && !filterHidden.has(row)) rowsVisible += 1;
    }
  }
  return {
    state,
    changed,
    sheetId: sheet.id,
    range: range === null ? null : { ...range },
    before,
    after,
    hiddenChanges,
    rowsVisible,
    rowsTotal,
    unsupportedColumns: evaluation?.unsupportedColumns ?? [],
  };
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/** The manual layer: the row's own hidden flag, imported from the file (§3.3). */
function manualHiddenRows(sheet: Sheet): ReadonlySet<number> {
  const hidden = new Set<number>();
  for (const [row, dimension] of sheet.rowDimensions) {
    if (dimension.hidden === true) hidden.add(row);
  }
  return hidden;
}

function sheetFor(context: CommandContext, sheetId?: SheetId): Sheet | undefined {
  const id = sheetId ?? context.activeSheetId;
  if (id === null || id === undefined) return undefined;
  return context.workbook.sheets.find((sheet) => sheet.id === id);
}

function rangeOfActiveCell(context: CommandContext): GridRange | null {
  const active = context.activeCell;
  if (active === null) return null;
  return { startRow: active.row, startCol: active.col, endRow: active.row, endCol: active.col };
}

function sameRange(left: GridRange, right: GridRange): boolean {
  return (
    left.startRow === right.startRow &&
    left.startCol === right.startCol &&
    left.endRow === right.endRow &&
    left.endCol === right.endCol
  );
}

/** A filter on a fresh range with no criteria; the shape enabling creates. */
function emptyFilter(sheetId: SheetId, range: GridRange): AutoFilter {
  return {
    sheetId,
    range: { ...range },
    columns: new Map<ColumnOffset, ColumnFilter>(),
  };
}

function withColumn(filter: AutoFilter, offset: ColumnOffset, criterion: ColumnFilter): AutoFilter {
  const next = cloneAutoFilter(filter);
  next.columns.set(offset, criterion);
  return next;
}

function withoutColumn(filter: AutoFilter, offset: ColumnOffset): AutoFilter {
  const next = cloneAutoFilter(filter);
  next.columns.delete(offset);
  return next;
}

/**
 * A checklist with nothing selected would hide every row, so it is refused the
 * way Excel refuses an empty checklist; conditions follow §3.2's two-condition
 * limit. An imported unsupported criterion is not authorable.
 */
function criterionProblem(criterion: ColumnFilter): string | null {
  switch (criterion.kind) {
    case 'values':
      if (criterion.selected.length === 0 && !criterion.includeBlanks) {
        return 'Select at least one value or (Blanks) before applying the filter.';
      }
      return null;
    case 'conditions':
      if (criterion.conditions.length === 0) {
        return 'Add at least one condition before applying the filter.';
      }
      if (criterion.conditions.length > MAX_CONDITIONS_PER_COLUMN) {
        return `A filter column can hold at most ${MAX_CONDITIONS_PER_COLUMN} conditions.`;
      }
      return null;
    case 'unsupported':
      return 'This filter uses a feature calcplus cannot edit; only clearing it is offered.';
  }
}

/** Two filter states are the same when the range and every criterion match. */
function autoFilterStateEquals(left: AutoFilter | null, right: AutoFilter | null): boolean {
  if (left === null || right === null) return left === right;
  if (!sameRange(left.range, right.range)) return false;
  if (left.columns.size !== right.columns.size) return false;
  for (const [offset, columnFilter] of left.columns) {
    const other = right.columns.get(offset);
    if (other === undefined || !columnFilterEquals(columnFilter, other)) return false;
  }
  return left.raw === right.raw;
}

function columnFilterEquals(left: ColumnFilter, right: ColumnFilter): boolean {
  switch (left.kind) {
    case 'values': {
      if (right.kind !== 'values') return false;
      if (left.includeBlanks !== right.includeBlanks) return false;
      if (left.selected.length !== right.selected.length) return false;
      return left.selected.every((value) =>
        right.selected.some((candidate) => cellValueEquals(value, candidate)),
      );
    }
    case 'conditions': {
      if (right.kind !== 'conditions') return false;
      if (left.join !== right.join) return false;
      if (left.conditions.length !== right.conditions.length) return false;
      return left.conditions.every((condition, index) =>
        conditionEquals(condition, right.conditions[index]!),
      );
    }
    case 'unsupported':
      return right.kind === 'unsupported' && left.raw === right.raw;
  }
}

function conditionEquals(left: Condition, right: Condition): boolean {
  if (left.kind !== right.kind || left.operator !== right.operator) return false;
  if ('from' in left || 'from' in right) {
    return 'from' in left && 'from' in right && left.from === right.from && left.to === right.to;
  }
  return 'value' in left && 'value' in right && left.value === right.value;
}

function toast(message: string): CommandResult {
  return { changes: [], ui: { type: 'toast', message } };
}
