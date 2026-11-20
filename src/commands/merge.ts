import type { Cell, GridRange, Sheet, SheetId, StyleRecord } from '@model/types';
import { cellIsEmpty, cloneCell, getCell, singleCellRange } from '@model/types';
import { BLANK } from '@model/value';
import {
  MERGE_DATA_WARNING,
  UNMERGE_NOTHING_MESSAGE,
  mergeListsEqual,
  normalizeRange,
  planMerge,
  planUnmerge,
  rangesIntersect,
  type MergeMode,
} from '@model/merge';
import { findOrAppendStyle, patchedStyleInput } from '@model/style-table';
import type {
  CellSetChange,
  CommandChange,
  CommandContext,
  CommandDefinition,
  CommandResult,
  StyleSetChange,
} from './types';

/**
 * Merge authoring (feature matrix "Merge & Center, Merge Across, Merge Cells,
 * Unmerge"; command-system §2 and §4).
 *
 * Each command is one transaction: `merge.set` carries the sheet's exact before
 * and after merge lists, the covered cells travel as `cell.set` patches so their
 * values are discarded exactly as Excel discards them, and Merge & Center adds
 * one `style.set` for the anchor's centred alignment. Undo reverses all of them
 * together and restores the prior alignment, values and merges.
 */

export interface MergeArgs {
  readonly sheetId?: SheetId;
  readonly range?: GridRange;
  /**
   * The data-loss prompt result. Without it, a merge that would discard
   * content returns a dialog request, the way `sheet.delete` asks before its
   * terminal edit. The UI re-dispatches with `confirm: true` when the user
   * accepts.
   */
  readonly confirm?: boolean;
}

export const mergeCenterCommand: CommandDefinition<MergeArgs> = mergeCommand(
  'merge.center',
  'Merge & Center',
  'center',
);

export const mergeAcrossCommand: CommandDefinition<MergeArgs> = mergeCommand(
  'merge.across',
  'Merge Across',
  'across',
);

export const mergeCellsCommand: CommandDefinition<MergeArgs> = mergeCommand(
  'merge.cells',
  'Merge Cells',
  'cells',
);

export const mergeUnmergeCommand: CommandDefinition<MergeArgs> = {
  id: 'merge.unmerge',
  label: 'Unmerge',
  category: 'format',
  keybindings: [],
  isEnabled: hasMergedSelection,
  run(context, args = {}) {
    return runUnmerge(context, args);
  },
};

export const mergeCommands = (): readonly CommandDefinition<MergeArgs>[] => [
  mergeCenterCommand,
  mergeAcrossCommand,
  mergeCellsCommand,
  mergeUnmergeCommand,
];

// ---------------------------------------------------------------------------
// Execution
// ---------------------------------------------------------------------------

function mergeCommand(id: string, label: string, mode: MergeMode): CommandDefinition<MergeArgs> {
  return {
    id,
    label,
    category: 'format',
    keybindings: [],
    isEnabled: hasMergeSelection,
    run(context, args = {}) {
      return runMerge(context, args, mode);
    },
  };
}

function runMerge(context: CommandContext, args: MergeArgs, mode: MergeMode): CommandResult {
  const sheet = resolveSheet(context, args.sheetId);
  if (sheet === undefined) return toast('Open a worksheet before merging cells.');
  const range = resolveRange(context, args);
  if (range === null) return toast('Select a range before merging cells.');

  const planned = planMerge(sheet, range, mode);
  if (planned.kind === 'blocked') return toast(planned.message);
  const plan = planned.plan;

  // A merge discards every value outside the anchor cells. Excel asks first,
  // and a declined prompt leaves the sheet untouched and creates no history.
  if (plan.discarded.length > 0 && args.confirm !== true) {
    return { changes: [], ui: { type: 'dialog', message: MERGE_DATA_WARNING } };
  }

  const changes: CommandChange[] = [];
  if (!mergeListsEqual(plan.mergesBefore, plan.mergesAfter)) {
    changes.push({
      type: 'merge.set',
      sheetId: sheet.id,
      before: plan.mergesBefore,
      after: plan.mergesAfter,
    });
  }
  for (const discarded of plan.discarded) {
    changes.push(clearCellChange(sheet.id, discarded.row, discarded.col, discarded.cell));
  }
  if (mode === 'center') {
    const centered = centerAnchorStyle(context, sheet, plan.range);
    if (centered !== null) changes.push(centered);
  }
  return { changes };
}

function runUnmerge(context: CommandContext, args: MergeArgs): CommandResult {
  const sheet = resolveSheet(context, args.sheetId);
  if (sheet === undefined) return toast('Open a worksheet before unmerging cells.');
  const range = resolveRange(context, args);
  if (range === null) return toast('Select a cell or range before unmerging cells.');

  const plan = planUnmerge(sheet, range);
  if (plan.mergesAfter.length === plan.mergesBefore.length) {
    return toast(UNMERGE_NOTHING_MESSAGE);
  }
  return {
    changes: [
      {
        type: 'merge.set',
        sheetId: sheet.id,
        before: plan.mergesBefore,
        after: plan.mergesAfter,
      },
    ],
  };
}

// ---------------------------------------------------------------------------
// Changes
// ---------------------------------------------------------------------------

/**
 * Discard one covered cell's content. Its style survives because a merge hides
 * formatting, it does not erase it; a plain unstyled cell is removed so the
 * sparse map stays sparse.
 */
function clearCellChange(sheetId: SheetId, row: number, col: number, cell: Cell): CellSetChange {
  const after: Cell | null =
    cell.styleId === 0
      ? null
      : {
          value: BLANK,
          formula: null,
          formulaDerived: false,
          savedResult: null,
          styleId: cell.styleId,
          richText: null,
          stale: false,
        };
  return {
    type: 'cell.set',
    sheetId,
    range: singleCellRange(row, col),
    before: [[cloneCell(cell)]],
    after: [[after]],
  };
}

/** Merge & Center writes the anchor's horizontal alignment, nothing else. */
function centerAnchorStyle(
  context: CommandContext,
  sheet: Sheet,
  range: GridRange,
): StyleSetChange | null {
  const cell = getCell(sheet, range.startRow, range.startCol);
  const base = styleRecordById(context.workbook, cell?.styleId ?? 0);
  const centeredId = findOrAppendStyle(
    context.workbook,
    patchedStyleInput(base, { alignment: { horizontal: 'center' } }),
  );
  const beforeId = cell === null || (cell.styleId === 0 && cellIsEmpty(cell)) ? null : cell.styleId;
  if (beforeId === centeredId) return null;
  // A style id of 0 on a cell that did not exist is not a record the applier
  // can create, so the patch would be a no-op.
  if (centeredId === 0 && beforeId === null) return null;
  return {
    type: 'style.set',
    sheetId: sheet.id,
    range: singleCellRange(range.startRow, range.startCol),
    before: [[beforeId]],
    after: [[centeredId]],
  };
}

function styleRecordById(workbook: CommandContext['workbook'], styleId: number): StyleRecord {
  return workbook.styles.find((style) => style.id === styleId) ?? workbook.styles[0]!;
}

// ---------------------------------------------------------------------------
// Target resolution and enablement
// ---------------------------------------------------------------------------

function resolveRange(context: CommandContext, args: MergeArgs): GridRange | null {
  if (args.range !== undefined) return normalizeRange(args.range);
  if (context.selection !== null) return normalizeRange(context.selection);
  const active = context.activeCell;
  if (active === null) return null;
  return normalizeRange(singleCellRange(active.row, active.col));
}

function resolveSheet(context: CommandContext, sheetId?: SheetId): Sheet | undefined {
  const id = sheetId ?? context.activeSheetId;
  if (id === null || id === undefined) return undefined;
  return context.workbook.sheets.find((sheet) => sheet.id === id);
}

/** The three merge commands need an active sheet and a multi-cell selection. */
function hasMergeSelection(context: CommandContext): boolean {
  const sheet = resolveSheet(context);
  if (sheet === undefined || context.selection === null) return false;
  const range = normalizeRange(context.selection);
  return range.startRow !== range.endRow || range.startCol !== range.endCol;
}

/** Unmerge needs an active sheet and a merged cell in the selection. */
function hasMergedSelection(context: CommandContext): boolean {
  const sheet = resolveSheet(context);
  if (sheet === undefined || context.selection === null) return false;
  const range = normalizeRange(context.selection);
  return sheet.merges.some((merge) => rangesIntersect(merge.range, range));
}

function toast(message: string): CommandResult {
  return { changes: [], ui: { type: 'toast', message } };
}
