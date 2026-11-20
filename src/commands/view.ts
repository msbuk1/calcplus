import type { CommandDefinition, CommandResult } from './types';

/**
 * `view.freeze` — Excel 365's Freeze Panes at the active cell (feature matrix
 * §6 "Freeze panes (rows and columns)").
 *
 * Oracle behaviour matched: Microsoft's documented rule freezes the rows above
 * the selected cell and the columns to its left. The task also pins the
 * repeat-press rule: pressing again from the split point unfreezes, pressing
 * elsewhere moves the split, and freezing at A1 leaves no split. Excel's
 * separate "Freeze Top Row" and "Freeze First Column" menu entries are not part
 * of the matrix row, so they are not built.
 *
 * Freeze state is document state because it round-trips in the sheet view, so
 * the command returns one `sheet.freeze` patch and undo restores the prior
 * split as one transaction.
 */

export interface FreezeArgs {
  readonly sheetId?: string;
}

export const freezePanesCommand: CommandDefinition<FreezeArgs> = {
  id: 'view.freeze',
  label: 'Freeze Panes',
  category: 'view',
  keybindings: [],
  isEnabled(context) {
    return context.workbook.sheets.length > 0;
  },
  isChecked(context) {
    return (context.view?.freezeRows ?? 0) > 0 || (context.view?.freezeCols ?? 0) > 0;
  },
  run(context, args = {}): CommandResult {
    const sheetId = args.sheetId ?? context.activeSheetId;
    const sheet =
      sheetId === null || sheetId === undefined
        ? undefined
        : context.workbook.sheets.find((candidate) => candidate.id === sheetId);
    if (sheet === undefined) return toast('Open a worksheet before freezing panes.');

    const active = context.activeCell;
    const row = active?.row ?? 0;
    const col = active?.col ?? 0;
    const before = { rows: sheet.view.freezeRows, cols: sheet.view.freezeCols };
    // Excel's split point for the pair of counts is the top-left cell visible
    // in the scrolling pane: B3 for two frozen rows and one frozen column.
    const atSplitPoint = before.rows === row && before.cols === col;
    const after =
      atSplitPoint || (row === 0 && col === 0) ? { rows: 0, cols: 0 } : { rows: row, cols: col };
    if (before.rows === after.rows && before.cols === after.cols) {
      return toast(
        'Select a cell below the first row or to the right of the first column to freeze panes.',
      );
    }
    const message =
      after.rows === 0 && after.cols === 0
        ? 'Unfroze panes.'
        : `Froze ${[
            ...(after.rows > 0 ? [plural(after.rows, 'row')] : []),
            ...(after.cols > 0 ? [plural(after.cols, 'column')] : []),
          ].join(' and ')}.`;
    return {
      changes: [{ type: 'sheet.freeze', sheetId: sheet.id, before, after }],
      ui: { type: 'toast', message },
    };
  },
};

function plural(count: number, noun: string): string {
  if (count === 1) return `1 ${noun}`;
  return `${count} ${noun}s`;
}

function toast(message: string): CommandResult {
  return { changes: [], ui: { type: 'toast', message } };
}
