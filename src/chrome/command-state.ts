import { createCommandRegistry } from '@commands/registry';
import type { CellAddress, GridRange, SheetId, Workbook } from '@model/types';
import type { CommandContext, CommandViewState } from '@commands/types';
import { createEmptyWorkbook } from '@model/workbook';
import type { CommandState } from './ribbon-model';

/** One registry for chrome metadata; the document session keeps its own. */
const registry = createCommandRegistry();
const emptyWorkbook = createEmptyWorkbook();

export interface ChromeCommandContext {
  readonly workbook: Workbook | null;
  readonly activeSheetId: SheetId | null;
  readonly selection: GridRange | null;
  readonly activeCell: CellAddress | null;
  readonly view: CommandViewState;
  readonly busy: boolean;
  readonly canUndo: boolean;
  readonly canRedo: boolean;
}

/**
 * Enablement and checked state come from the command registry (ui-chrome §3).
 * With no document open the registry still answers, using an empty workbook so
 * New and Open stay available while document commands stay disabled.
 */
export function resolveCommandState(
  context: ChromeCommandContext,
  commandId: string,
): CommandState {
  const command = registry.get(commandId);
  if (command === undefined) return { enabled: false, checked: false };
  const commandContext: CommandContext = {
    workbook: context.workbook ?? emptyWorkbook,
    activeSheetId: context.activeSheetId,
    selection: context.selection,
    activeCell: context.activeCell,
    view: context.view,
    busy: context.busy,
    canUndo: context.canUndo,
    canRedo: context.canRedo,
  };
  return {
    enabled: command.isEnabled(commandContext),
    checked: command.isChecked?.(commandContext) ?? false,
  };
}
