import { evaluateAutoFilter } from '@model/autofilter';
import type { SortAttachmentMove, SortCellRecord, SortKey, SortPlan } from '@model/sort';
import { guessHeaderRow, planSort, resolveTargetRange } from '@model/sort';
import type { Cell, GridRange, Sheet, SheetId } from '@model/types';
import { cellKey, cloneCell, formatCellRef, parseCellKey } from '@model/types';
import type {
  CommandChange,
  CommandContext,
  CommandDefinition,
  CommandResult,
  SortApplyChange,
  SortAttachmentChange,
} from './types';

/**
 * `data.sort` (sort-autofilter §2). The model's `planSort` is the whole
 * decision — target range, header guess, comparator, blocking and warnings; the
 * command turns the plan into one transaction and reports the outcome.
 *
 * One sort is one transaction (§2.7): every moved row arrives as a `cell.set`
 * patch so the calculation layer recalculates it, and one `sort.apply` record
 * carries the permutation, the formula rewrites, the merged ranges and the
 * comments and hyperlinks that move with their rows. Undo restores the exact
 * prior state and redo reapplies the same permutation, because both sides of
 * every patch are stored.
 */

export interface SortArgs {
  readonly sheetId?: SheetId;
  /** The target range; defaults to the selection. One cell expands to its region (§2.1). */
  readonly range?: GridRange;
  /**
   * Up to 64 keys in order (§2.3). Omitting them is the ribbon/menu path: the
   * command asks the Sort dialog to collect them instead of guessing.
   */
  readonly keys?: readonly (SortKey | SortQuickKey)[];
  /** `null` or omitted guesses with the §2.2 heuristic; a boolean overrides it. */
  readonly hasHeaders?: boolean | null;
}

/**
 * A menu key: Sort Ascending / Sort Descending sort by the active column, as
 * Excel does, so the map does not need to know the column when it is authored.
 */
export interface SortQuickKey {
  readonly direction: SortKey['direction'];
}

export const sortCommand: CommandDefinition<SortArgs> = {
  id: 'data.sort',
  label: 'Sort',
  category: 'data',
  keybindings: [],
  isEnabled(context) {
    return context.workbook.sheets.length > 0;
  },
  run(context, args = {}): CommandResult {
    return runSort(context, args);
  },
};

function runSort(context: CommandContext, args: SortArgs): CommandResult {
  const sheet = resolveSheet(context, args.sheetId);
  if (sheet === undefined) return toast('Open a worksheet before sorting.');
  const selection = args.range ?? context.selection ?? rangeOfActiveCell(context);
  if (selection === null) return toast('Select a cell or range to sort.');
  const range = resolveTargetRange(sheet, selection);

  const keys = normaliseKeys(args.keys ?? [], range, context);
  if (keys.length === 0) {
    // Dialogs are commands: the ribbon and menu open the sort dialog, and the
    // dialog submits this command again with keys.
    return {
      changes: [],
      sort: {
        state: 'needs-keys',
        sheetId: sheet.id,
        range,
        hasHeaders: guessHeaderRow({ workbook: context.workbook, sheet, range }),
        keyLabels: [],
        warnings: [],
      },
      ui: {
        type: 'dialog',
        message: `Sort ${formatRange(range)}. Choose at least one sort key, then confirm.`,
      },
    };
  }

  const filter = sheet.autoFilter;
  const result = planSort({
    workbook: context.workbook,
    sheetId: sheet.id,
    selection,
    keys,
    hasHeaders: args.hasHeaders ?? null,
    ...(filter === null
      ? {}
      : {
          // Filter-hidden rows stay in place, exactly like manual hidden rows
          // (§2.6, §3.4). The criteria are unchanged by the sort.
          filterHiddenRows: evaluateAutoFilter(sheet, filter).filterHidden,
        }),
  });

  switch (result.kind) {
    case 'blocked':
      return {
        changes: [],
        sort: {
          state: 'blocked',
          sheetId: sheet.id,
          range: result.range,
          code: result.code,
          message: result.message,
          warnings: [],
        },
        ui: { type: 'toast', message: result.message },
      };
    case 'unchanged':
      return {
        changes: [],
        sort: {
          state: 'unchanged',
          sheetId: sheet.id,
          range: result.range,
          reason: result.reason,
          warnings: result.warnings,
        },
      };
    case 'sorted': {
      const plan = result.plan;
      const changes = changesForPlan(sheet, plan);
      if (changes === null) {
        // The permutation is empty, or it only moved rows with no content and
        // no attached record: nothing observable changed, so no transaction.
        return {
          changes: [],
          sort: {
            state: 'unchanged',
            sheetId: sheet.id,
            range: plan.range,
            reason: 'nothing-to-order',
            hasHeaders: plan.hasHeaders,
            keyLabels: plan.keyLabels,
            warnings: result.warnings,
            visibleRowCount: plan.visibleRowCount,
            hiddenRowCount: plan.hiddenRowCount,
          },
        };
      }
      return {
        changes,
        sort: {
          state: 'sorted',
          sheetId: sheet.id,
          range: plan.range,
          hasHeaders: plan.hasHeaders,
          keyLabels: plan.keyLabels,
          warnings: result.warnings,
          visibleRowCount: plan.visibleRowCount,
          hiddenRowCount: plan.hiddenRowCount,
        },
      };
    }
  }
}

/**
 * The transaction's changes, or `null` when the sort changes nothing. The moved
 * rows are `cell.set` patches because that is the change the calculation layer
 * applies; the sort-specific state rides in one `sort.apply` record.
 */
function changesForPlan(sheet: Sheet, plan: SortPlan): CommandChange[] | null {
  const cells = cellChanges(sheet, plan);
  const sortApply = sortApplyChange(sheet, plan);
  if (
    cells.length === 0 &&
    plan.mergeRewrites.length === 0 &&
    plan.commentMoves.length === 0 &&
    plan.hyperlinkMoves.length === 0
  ) {
    return null;
  }
  return [...cells, sortApply];
}

/** One rectangular patch per destination row the sort rewrites. */
function cellChanges(sheet: Sheet, plan: SortPlan): CommandChange[] {
  const records = new Map<string, SortCellRecord>();
  for (const record of plan.cellRecords) {
    records.set(cellKey(record.address.row, record.address.col), record);
  }
  const destinations = [...new Set(plan.rowMoves.map((move) => move.toRow))].sort(
    (left, right) => left - right,
  );
  const changes: CommandChange[] = [];
  for (const row of destinations) {
    const before: (Cell | null)[] = [];
    const after: (Cell | null)[] = [];
    for (let col = plan.range.startCol; col <= plan.range.endCol; col += 1) {
      const record = records.get(cellKey(row, col));
      // The plan omits a destination whose old and new content are both empty,
      // so an absent record is empty on both sides.
      before.push(record === undefined ? null : cloneCell(record.before));
      after.push(record === undefined ? null : cloneCell(record.after));
    }
    changes.push({
      type: 'cell.set',
      sheetId: sheet.id,
      range: {
        startRow: row,
        startCol: plan.range.startCol,
        endRow: row,
        endCol: plan.range.endCol,
      },
      before: [before],
      after: [after],
    });
  }
  return changes;
}

function sortApplyChange(sheet: Sheet, plan: SortPlan): SortApplyChange {
  const before = sheet.merges.map((merge) => ({ ...merge.range }));
  const moved = new Map<string, GridRange>();
  for (const rewrite of plan.mergeRewrites) moved.set(rangeKey(rewrite.before), rewrite.after);
  const after = before.map((range) => {
    const destination = moved.get(rangeKey(range));
    return destination === undefined ? range : { ...destination };
  });
  return {
    type: 'sort.apply',
    sheetId: sheet.id,
    range: { ...plan.range },
    rowPermutation: [...plan.rowPermutation],
    formulaRewrites: plan.formulaRewrites.map((rewrite) => ({
      address: { ...rewrite.address },
      before: rewrite.before,
      after: rewrite.after,
    })),
    merges: { before, after },
    comments: attachmentChange(sheet.comments, plan, plan.commentMoves),
    hyperlinks: attachmentChange(sheet.hyperlinks, plan, plan.hyperlinkMoves),
  };
}

/**
 * The applier contract `src/model/sort.ts` documents: clear every attachment in
 * a moved row that lies inside the sorted columns, then insert the moved
 * records at their destinations. Attachments outside the sorted columns keep
 * their address, because the sort only permutes the selected range.
 */
function attachmentChange<T extends object>(
  records: ReadonlyMap<string, T>,
  plan: SortPlan,
  moves: readonly SortAttachmentMove<T>[],
): SortAttachmentChange<T> {
  const before: (readonly [string, T])[] = [];
  for (const [key, record] of records) before.push([key, { ...record }]);

  const movingRows = new Set<number>();
  for (const move of plan.rowMoves) {
    movingRows.add(move.fromRow);
    movingRows.add(move.toRow);
  }

  const after = new Map<string, T>();
  for (const [key, record] of records) {
    const address = parseCellKey(key);
    const cleared =
      movingRows.has(address.row) &&
      address.col >= plan.range.startCol &&
      address.col <= plan.range.endCol;
    if (!cleared) after.set(key, { ...record });
  }
  for (const move of moves) after.set(cellKey(move.to.row, move.to.col), { ...move.record });

  return { before, after: [...after.entries()] };
}

function rangeKey(range: GridRange): string {
  return `${range.startRow}:${range.startCol}:${range.endRow}:${range.endCol}`;
}

function formatRange(range: GridRange): string {
  const start = formatCellRef(range.startRow, range.startCol);
  const end = formatCellRef(range.endRow, range.endCol);
  return start === end ? start : `${start}:${end}`;
}

function resolveSheet(context: CommandContext, sheetId?: SheetId): Sheet | undefined {
  const id = sheetId ?? context.activeSheetId;
  if (id === null || id === undefined) return undefined;
  return context.workbook.sheets.find((sheet) => sheet.id === id);
}

function rangeOfActiveCell(context: CommandContext): GridRange | null {
  const active = context.activeCell;
  if (active === null) return null;
  return { startRow: active.row, startCol: active.col, endRow: active.row, endCol: active.col };
}

/**
 * A quick key ("Sort Ascending" in the Sort & Filter menu) sorts by the active
 * column, clamped into the target range, exactly as the Data tab's A→Z button
 * does. A full key passes through untouched.
 */
function normaliseKeys(
  keys: readonly (SortKey | SortQuickKey)[],
  range: GridRange,
  context: CommandContext,
): SortKey[] {
  const quickColumn = Math.min(
    Math.max(context.activeCell?.col ?? range.startCol, range.startCol),
    range.endCol,
  );
  return keys.map((key) =>
    'column' in key ? key : { column: quickColumn, direction: key.direction },
  );
}

function toast(message: string): CommandResult {
  return { changes: [], ui: { type: 'toast', message } };
}
