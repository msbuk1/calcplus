import { resolveReferenceText } from '@model/references';
import type { CommandDefinition, CommandResult } from './types';

/**
 * Go To submission command (feature matrix §1).
 *
 * Accepts an A1 address, a range, a whole column/row, or a workbook- or
 * sheet-scoped defined name. Invalid text returns a message and no
 * `selectionAfter`, so the caller leaves the selection exactly where it was.
 */

export interface GoToArgs {
  readonly target: string;
}

export const goToCommand: CommandDefinition<GoToArgs> = {
  id: 'goto.goto',
  label: 'Go To',
  category: 'edit',
  keybindings: ['Ctrl+G'],
  isEnabled(context) {
    return context.workbook.sheets.length > 0;
  },
  run(context, args): CommandResult {
    const target = typeof args?.target === 'string' ? args.target : '';
    if (target.trim() === '') {
      return invalid('Type a cell address or defined name.');
    }
    const resolved = resolveReferenceText(context.workbook, target, context.activeSheetId);
    if (resolved.kind === 'invalid') return invalid(resolved.message);
    return {
      changes: [],
      selectionAfter: { ...resolved.range },
      activeSheetId: resolved.sheetId,
    };
  },
};

function invalid(message: string): CommandResult {
  return { changes: [], ui: { type: 'toast', message } };
}
