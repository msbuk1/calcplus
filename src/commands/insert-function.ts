import type { CommandDefinition, CommandResult } from './types';

/**
 * Insert Function (feature matrix §4, ticket 30).
 *
 * The command owns only the request: it asks the chrome to open the dialog and
 * changes nothing. The chosen function's skeleton is an edit in the formula
 * bar and becomes a `cell.set` only when the user commits the edit, so opening
 * the dialog is never an undoable transaction.
 */
export const insertFunctionCommand: CommandDefinition<void> = {
  id: 'formula.insertFunction',
  label: 'Insert Function',
  category: 'formula',
  keybindings: [],
  isEnabled(context) {
    return context.activeSheetId !== null && context.workbook.sheets.length > 0;
  },
  run(): CommandResult {
    return { changes: [], ui: { type: 'dialog', message: 'Insert Function' } };
  },
};
