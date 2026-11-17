import {
  createSheet,
  nextNewSheetId,
  nextSheetName,
  nextWorksheetPartPath,
  validateSheetName,
} from '@model/workbook';
import {
  planSheetDeleteReferences,
  planSheetRename,
  type SheetReferencePlan,
} from '@model/sheet-references';
import type { Cell, SheetId, Workbook } from '@model/types';
import { cellKey, cloneCell, singleCellRange } from '@model/types';
import { errorValue } from '@model/value';
import { isSheetVisible, visibleSheetCount } from '@view/sheet-visibility';
import type { CommandChange, CommandDefinition, CommandResult } from './types';

export interface SheetInsertArgs {
  /** Optional Excel-style name; the command picks `SheetN` when omitted. */
  readonly name?: string;
  /** Tab position; defaults to the end of the tab order. */
  readonly index?: number;
}

export interface SheetDeleteArgs {
  readonly sheetId?: SheetId;
  /**
   * The UI prompt result. Without it the command returns a dialog request.
   * Deleting a sheet is terminal, so every delete asks first; there is no
   * undo to fall back on when the answer was a mistake.
   */
  readonly confirm?: boolean;
}

export interface SheetRenameArgs {
  readonly sheetId?: SheetId;
  readonly name: string;
}

export const sheetInsertCommand: CommandDefinition<SheetInsertArgs> = {
  id: 'sheet.insert',
  label: 'Insert sheet',
  category: 'edit',
  keybindings: ['Shift+F11'],
  isEnabled() {
    return true;
  },
  run(context, args): CommandResult {
    const index =
      args.index === undefined || !Number.isInteger(args.index)
        ? context.workbook.sheets.length
        : Math.min(Math.max(0, args.index), context.workbook.sheets.length);
    const name = args.name ?? nextSheetName(context.workbook);
    const problem = validateSheetName(context.workbook, name);
    if (problem !== null) {
      return { changes: [], ui: { type: 'toast', message: problem } };
    }
    const partPath = nextWorksheetPartPath(context.workbook);
    const sheet = createSheet(name, partPath, nextNewSheetId(context.workbook, partPath));
    // A new sheet takes Excel's default grid and sheet settings: no frozen
    // panes, gridlines and headings visible, 100% zoom, 15pt rows.
    return {
      changes: [{ type: 'sheet.add', index, before: null, after: sheet }],
      activeSheetId: sheet.id,
    };
  },
};

export const sheetDeleteCommand: CommandDefinition<SheetDeleteArgs> = {
  id: 'sheet.delete',
  label: 'Delete sheet',
  category: 'edit',
  keybindings: [],
  history: 'terminal',
  isEnabled(context) {
    return context.workbook.sheets.length > 0;
  },
  run(context, args): CommandResult {
    const sheetId = args.sheetId ?? context.activeSheetId;
    const index =
      sheetId === null
        ? -1
        : context.workbook.sheets.findIndex((candidate) => candidate.id === sheetId);
    if (index < 0) return { changes: [] };
    const sheet = context.workbook.sheets[index]!;
    // Deleting a hidden sheet leaves the visible count alone, so only a
    // visible sheet can be the last visible one. A workbook must always keep
    // at least one visible sheet.
    const visibleAfterDelete =
      visibleSheetCount(context.workbook) - (isSheetVisible(context.workbook, sheet.id) ? 1 : 0);
    if (visibleAfterDelete < 1) {
      return {
        changes: [],
        ui: {
          type: 'toast',
          message:
            'A workbook must contain at least one visible sheet; this sheet was not deleted.',
        },
      };
    }
    const plan = planSheetDeleteReferences(context.workbook, sheet.id);
    if (plan.blockedReason !== null) {
      return { changes: [], ui: { type: 'toast', message: plan.blockedReason } };
    }
    // The content prompt folds into the permanent-delete warning: the sheet's
    // dirty cells can never be written once the sheet is gone.
    if (!args.confirm) {
      return {
        changes: [],
        ui: {
          type: 'dialog',
          message: `Deleting "${sheet.name}" cannot be undone, and its data may be lost. Delete the sheet?`,
        },
      };
    }
    const changes: CommandChange[] = [
      { type: 'sheet.delete', index, before: sheet, after: null },
      ...formulaAndNameChanges(context.workbook, plan, 'delete'),
    ];
    const activeSheetId =
      context.activeSheetId === null || context.activeSheetId === sheet.id
        ? neighbourSheetId(context.workbook, sheet.id, index)
        : context.activeSheetId;
    return { changes, activeSheetId };
  },
};

export const sheetRenameCommand: CommandDefinition<SheetRenameArgs> = {
  id: 'sheet.rename',
  label: 'Rename sheet',
  category: 'edit',
  keybindings: [],
  isEnabled(context) {
    return context.workbook.sheets.length > 0;
  },
  run(context, args): CommandResult {
    const sheetId = args.sheetId ?? context.activeSheetId;
    const sheet =
      sheetId === null
        ? undefined
        : context.workbook.sheets.find((candidate) => candidate.id === sheetId);
    if (sheet === undefined) return { changes: [] };
    if (!isSheetVisible(context.workbook, sheet.id)) {
      return {
        changes: [],
        ui: {
          type: 'toast',
          message: `Sheet "${sheet.name}" is hidden and cannot be renamed in this milestone.`,
        },
      };
    }
    const problem = validateSheetName(context.workbook, args.name, sheet.id);
    if (problem !== null) {
      return { changes: [], ui: { type: 'toast', message: problem } };
    }
    if (args.name === sheet.name) return { changes: [] };

    const plan = planSheetRename(context.workbook, sheet.id, args.name);
    if (plan.blockedReason !== null) {
      return { changes: [], ui: { type: 'toast', message: plan.blockedReason } };
    }
    const changes: CommandChange[] = [
      { type: 'sheet.rename', sheetId: sheet.id, before: sheet.name, after: args.name },
      ...formulaAndNameChanges(context.workbook, plan, 'rename'),
    ];
    return { changes, activeSheetId: context.activeSheetId ?? sheet.id };
  },
};

function formulaAndNameChanges(
  workbook: Workbook,
  plan: SheetReferencePlan,
  mode: 'rename' | 'delete',
): CommandChange[] {
  const changes: CommandChange[] = [];
  for (const rewrite of plan.formulaRewrites) {
    const sheet = workbook.sheets.find((candidate) => candidate.id === rewrite.sheetId);
    const cell = sheet?.cells.get(cellKey(rewrite.row, rewrite.col));
    if (sheet === undefined || cell === undefined) continue;
    const before = cloneCell(cell);
    if (before === null) continue;
    const after: Cell =
      mode === 'delete'
        ? {
            ...before,
            formula: rewrite.after,
            value: errorValue('#REF!'),
            savedResult: null,
            stale: true,
          }
        : { ...before, formula: rewrite.after };
    changes.push({
      type: 'cell.set',
      sheetId: sheet.id,
      range: singleCellRange(rewrite.row, rewrite.col),
      before: [[before]],
      after: [[after]],
    });
  }
  for (const name of plan.nameChanges) {
    changes.push({
      type: 'name.set',
      scope: name.scope,
      sheetId: name.sheetId,
      name: name.name,
      before: name.before,
      after: name.after,
    });
  }
  return changes;
}

function neighbourSheetId(
  workbook: Workbook,
  removedSheetId: SheetId,
  removedIndex: number,
): SheetId | null {
  if (workbook.sheets.length <= 1) return null;
  const remaining = workbook.sheets.filter((sheet) => sheet.id !== removedSheetId);
  if (remaining.length === 0) return null;
  const position = Math.min(Math.max(0, removedIndex), remaining.length - 1);
  // Excel activates the next visible tab; imported hidden sheets are skipped.
  for (let index = position; index < remaining.length; index += 1) {
    const sheet = remaining[index]!;
    if (isSheetVisible(workbook, sheet.id)) return sheet.id;
  }
  for (let index = position - 1; index >= 0; index -= 1) {
    const sheet = remaining[index]!;
    if (isSheetVisible(workbook, sheet.id)) return sheet.id;
  }
  return remaining[0]!.id;
}
