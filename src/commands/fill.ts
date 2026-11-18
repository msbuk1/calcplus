import type { FillBlockCode, FillStopReason } from '@model/fill';
import { planFill } from '@model/fill';
import type { GridRange, Sheet, SheetId } from '@model/types';
import type {
  CommandChange,
  CommandContext,
  CommandDefinition,
  CommandResult,
  FillPayload,
} from './types';

/**
 * Fill commands (feature matrix §1 "Fill handle: copy, relative-reference
 * adjustment, double-click to adjacent data"; command-system §4: "Fill-handle
 * drag or double-click fill | One").
 *
 * One gesture is one transaction. The model's `planFill` returns the filled
 * bands and the exact before/after records; each band becomes one `cell.set`
 * patch in the same `CommandResult`, exactly like the sort applier's row
 * patches. That keeps the calculation layer seeing the new cells and makes
 * undo exact without a new change type in `session.ts`.
 *
 * Series, growth and custom-list fills are v2 (feature matrix §9).
 */

/**
 * The fill handle gesture. `source` is the selection the drag started from and
 * `destination` is the whole rectangle the drag covered, source included (the
 * grid passes both). A double-click passes the destination the adjacent data
 * dictates. With neither, the selection fills onto itself and nothing happens.
 */
export interface FillHandleArgs {
  readonly sheetId?: SheetId;
  readonly source?: GridRange;
  readonly destination?: GridRange;
}

/** Ctrl+D and Ctrl+R: the whole selection, filled from its first row or column. */
export interface FillDirectionArgs {
  readonly sheetId?: SheetId;
  readonly range?: GridRange;
}

export const fillHandleCommand: CommandDefinition<FillHandleArgs> = {
  id: 'fill.handle',
  label: 'Fill',
  category: 'edit',
  keybindings: [],
  isEnabled: hasSheet,
  run(context, args = {}) {
    const sheet = resolveSheet(context, args.sheetId);
    if (sheet === undefined) return toast('Open a worksheet before filling cells.');
    const source = args.source ?? context.selection ?? rangeOfActiveCell(context);
    if (source === null) return toast('Select a cell or range to fill from.');
    return applyFill(sheet, source, args.destination ?? source);
  },
};

/** Ctrl+D: the first row of the selection fills every row below it. */
export const fillDownCommand: CommandDefinition<FillDirectionArgs> = {
  id: 'edit.fillDown',
  label: 'Fill Down',
  category: 'edit',
  keybindings: ['Ctrl+D'],
  isEnabled: hasSheet,
  run(context, args = {}) {
    return runDirectional(context, args, 'down');
  },
};

/** Ctrl+R: the first column of the selection fills every column to its right. */
export const fillRightCommand: CommandDefinition<FillDirectionArgs> = {
  id: 'edit.fillRight',
  label: 'Fill Right',
  category: 'edit',
  keybindings: ['Ctrl+R'],
  isEnabled: hasSheet,
  run(context, args = {}) {
    return runDirectional(context, args, 'right');
  },
};

/** Shared by the registry and tests. */
export function fillCommands(): readonly CommandDefinition<never>[] {
  return [
    fillHandleCommand as unknown as CommandDefinition<never>,
    fillDownCommand as unknown as CommandDefinition<never>,
    fillRightCommand as unknown as CommandDefinition<never>,
  ];
}

function runDirectional(
  context: CommandContext,
  args: FillDirectionArgs,
  direction: 'down' | 'right',
): CommandResult {
  const sheet = resolveSheet(context, args.sheetId);
  if (sheet === undefined) return toast('Open a worksheet before filling cells.');
  const range = args.range ?? context.selection ?? rangeOfActiveCell(context);
  if (range === null) return toast('Select a cell or range to fill.');
  // The whole selection is the destination; its first row (Ctrl+D) or first
  // column (Ctrl+R) is the block that repeats. A one-row or one-column
  // selection has no destination, so the command is a silent no-op exactly as
  // Excel's is, with no transaction.
  const source: GridRange =
    direction === 'down'
      ? {
          startRow: range.startRow,
          endRow: range.startRow,
          startCol: range.startCol,
          endCol: range.endCol,
        }
      : {
          startRow: range.startRow,
          endRow: range.endRow,
          startCol: range.startCol,
          endCol: range.startCol,
        };
  return applyFill(sheet, source, range);
}

function applyFill(sheet: Sheet, source: GridRange, destination: GridRange): CommandResult {
  const result = planFill({ sheet, source, destination });
  switch (result.kind) {
    case 'blocked':
      return {
        changes: [],
        fill: payload('blocked', sheet.id, source, destination, {
          code: result.code,
          message: result.message,
        }),
        ui: { type: 'toast', message: result.message },
      };
    case 'unchanged':
      return {
        changes: [],
        fill: payload('unchanged', sheet.id, source, destination, { reason: result.reason }),
      };
    case 'filled': {
      const plan = result.plan;
      const changes: CommandChange[] = plan.bands.map((band) => ({
        type: 'cell.set',
        sheetId: sheet.id,
        range: band.range,
        before: band.before,
        after: band.after,
      }));
      return {
        changes,
        // The filled rectangle becomes the selection, as Excel leaves it.
        selectionAfter: { ...plan.destination },
        fill: payload('filled', sheet.id, source, destination, {
          cellCount: plan.cellCount,
          formulaCount: plan.formulaCount,
        }),
      };
    }
  }
}

function payload(
  state: FillPayload['state'],
  sheetId: SheetId,
  source: GridRange,
  destination: GridRange,
  extra: {
    readonly reason?: FillStopReason;
    readonly code?: FillBlockCode;
    readonly message?: string;
    readonly cellCount?: number;
    readonly formulaCount?: number;
  } = {},
): FillPayload {
  return {
    state,
    sheetId,
    source: { ...source },
    destination: { ...destination },
    ...(extra.reason === undefined ? {} : { reason: extra.reason }),
    ...(extra.code === undefined ? {} : { code: extra.code }),
    ...(extra.message === undefined ? {} : { message: extra.message }),
    cellCount: extra.cellCount ?? 0,
    formulaCount: extra.formulaCount ?? 0,
  };
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

function hasSheet(context: CommandContext): boolean {
  return context.workbook.sheets.length > 0;
}

function toast(message: string): CommandResult {
  return { changes: [], ui: { type: 'toast', message } };
}
