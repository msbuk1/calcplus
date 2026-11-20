import type { Cell, GridRange, Sheet, SheetId } from '@model/types';
import { cellStateEquals, cloneCell, getCell } from '@model/types';
import { BLANK } from '@model/value';
import type { CommandChange, CommandContext, CommandDefinition, CommandResult } from './types';

export interface ClearArgs {
  readonly sheetId?: SheetId;
  readonly range?: GridRange;
}

/** Clearing a huge selection is one bounded patch; larger targets ask for a smaller range. */
const MAX_CLEAR_CELLS = 250_000;

export const clearContentsCommand: CommandDefinition<ClearArgs> = {
  id: 'edit.clearContents',
  label: 'Clear contents',
  category: 'edit',
  keybindings: ['Delete', 'Backspace'],
  isEnabled: hasSheet,
  run(context, args = {}) {
    return clearRange(context, args, (before) => {
      if (before === null) return null;
      if (before.styleId === 0) return null;
      return {
        value: BLANK,
        formula: null,
        formulaDerived: false,
        savedResult: null,
        styleId: before.styleId,
        richText: null,
        // Clear contents is an explicit deletion; the marker records that the
        // dropped runs were deliberate (file-compat §9 F17, save guard).
        ...(before.richText === null ? {} : { richTextFlattened: true }),
        stale: false,
      };
    });
  },
};

/**
 * Clear All removes every modelled thing the range holds: values, formulas,
 * rich text, cached results and formatting. A cell record with nothing left is
 * deleted outright, which is how `cell.set` represents an empty cell; the whole
 * range travels as one patch and therefore one undo transaction.
 *
 * Imported comments and hyperlinks stay on the sheet. They are read-only in
 * v1 (feature matrix §8) and are not part of the cell record this command
 * clears.
 */
export const clearAllCommand: CommandDefinition<ClearArgs> = {
  id: 'edit.clearAll',
  label: 'Clear all',
  category: 'edit',
  keybindings: [],
  isEnabled: hasSheet,
  run(context, args = {}) {
    return clearRange(context, args, () => null);
  },
};

function clearRange(
  context: CommandContext,
  args: ClearArgs,
  afterFor: (before: Cell | null) => Cell | null,
): CommandResult {
  const sheetId = args.sheetId ?? context.activeSheetId;
  const sheet =
    sheetId === null
      ? undefined
      : context.workbook.sheets.find((candidate) => candidate.id === sheetId);
  if (sheet === undefined) return toast('Open a worksheet before clearing cells.');
  const range = args.range ?? context.selection ?? rangeOfActiveCell(context);
  if (range === null || !validRange(range)) return toast('Select a cell or range first.');
  if (cellCount(range) > MAX_CLEAR_CELLS) {
    return toast(
      `Select at most ${MAX_CLEAR_CELLS.toLocaleString('en-US')} cells before clearing.`,
    );
  }
  const change = buildChange(sheet, range, afterFor);
  return change === null ? { changes: [] } : { changes: [change] };
}

function buildChange(
  sheet: Sheet,
  range: GridRange,
  afterFor: (before: Cell | null) => Cell | null,
): CommandChange | null {
  const beforeRows: (Cell | null)[][] = [];
  const afterRows: (Cell | null)[][] = [];
  let changed = false;
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    const beforeRow: (Cell | null)[] = [];
    const afterRow: (Cell | null)[] = [];
    for (let col = range.startCol; col <= range.endCol; col += 1) {
      const before = cloneCell(getCell(sheet, row, col));
      const after = afterFor(before);
      beforeRow.push(before);
      afterRow.push(after);
      if (!cellStateEquals(before, after)) changed = true;
    }
    beforeRows.push(beforeRow);
    afterRows.push(afterRow);
  }
  if (!changed) return null;
  return {
    type: 'cell.set',
    sheetId: sheet.id,
    range,
    before: beforeRows,
    after: afterRows,
  };
}

function hasSheet(context: CommandContext): boolean {
  return context.workbook.sheets.length > 0;
}

function rangeOfActiveCell(context: CommandContext): GridRange | null {
  const active = context.activeCell;
  if (active === null) return null;
  return {
    startRow: active.row,
    startCol: active.col,
    endRow: active.row,
    endCol: active.col,
  };
}

function validRange(range: GridRange): boolean {
  return (
    Number.isInteger(range.startRow) &&
    Number.isInteger(range.startCol) &&
    Number.isInteger(range.endRow) &&
    Number.isInteger(range.endCol) &&
    range.startRow >= 0 &&
    range.startCol >= 0 &&
    range.endRow >= range.startRow &&
    range.endCol >= range.startCol &&
    range.endRow < 1_048_576 &&
    range.endCol < 16_384
  );
}

function cellCount(range: GridRange): number {
  return (range.endRow - range.startRow + 1) * (range.endCol - range.startCol + 1);
}

function toast(message: string): CommandResult {
  return { changes: [], ui: { type: 'toast', message } };
}
