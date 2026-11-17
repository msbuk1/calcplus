import type { Cell, GridRange, Sheet } from '@model/types';
import { cellIsEmpty, cellStateEquals, cloneCell, formatCellRef, getCell } from '@model/types';
import { BLANK } from '@model/value';
import type { CommandChange, CommandDefinition, CommandResult } from './types';

/**
 * AutoSum inserts `SUM(...)` over the nearest contiguous block above the
 * active cell, then to its left. With no adjacent block it inserts `SUM()` and
 * leaves the caret to the user. The command is one cell.set change, so undo is
 * one transaction.
 */
export const autoSumCommand: CommandDefinition<void> = {
  id: 'formula.autosum',
  label: 'AutoSum',
  category: 'formula',
  keybindings: ['Alt+='],
  isEnabled(context) {
    return context.workbook.sheets.length > 0;
  },
  run(context): CommandResult {
    const sheetId = context.activeSheetId;
    const sheet =
      sheetId === null
        ? undefined
        : context.workbook.sheets.find((candidate) => candidate.id === sheetId);
    const target = context.activeCell;
    if (sheet === undefined || target === null) {
      return { changes: [], ui: { type: 'toast', message: 'Select a cell before using AutoSum.' } };
    }
    const before = cloneCell(getCell(sheet, target.row, target.col));
    if (
      before?.formulaSource?.kind === 'array' ||
      before?.formulaSource?.kind === 'shared-master'
    ) {
      return readOnly('Legacy array formulas are read-only in this milestone.');
    }
    if (before?.formulaSource?.kind === 'shared-follower') {
      return readOnly('Shared formulas are read-only in this milestone.');
    }
    if (before?.richText !== null && before?.richText !== undefined) {
      return readOnly('Rich text is read-only in this milestone.');
    }
    const source = adjacentSource(sheet, target.row, target.col);
    const formula = source === null ? 'SUM()' : sumFormula(source);
    const after: Cell = {
      value: BLANK,
      formula,
      formulaDerived: false,
      savedResult: null,
      styleId: before?.styleId ?? 0,
      richText: null,
      stale: true,
    };
    if (cellStateEquals(before, after)) return { changes: [] };
    const range: GridRange = {
      startRow: target.row,
      startCol: target.col,
      endRow: target.row,
      endCol: target.col,
    };
    const change: CommandChange = {
      type: 'cell.set',
      sheetId: sheet.id,
      range,
      before: [[before]],
      after: [[after]],
    };
    return { changes: [change], selectionAfter: range };
  },
};

function adjacentSource(sheet: Sheet, row: number, col: number): GridRange | null {
  let startRow = row - 1;
  while (startRow >= 0 && hasContent(getCell(sheet, startRow, col))) startRow -= 1;
  if (row - 1 > startRow) {
    return {
      startRow: startRow + 1,
      startCol: col,
      endRow: row - 1,
      endCol: col,
    };
  }
  let startCol = col - 1;
  while (startCol >= 0 && hasContent(getCell(sheet, row, startCol))) startCol -= 1;
  if (col - 1 > startCol) {
    return {
      startRow: row,
      startCol: startCol + 1,
      endRow: row,
      endCol: col - 1,
    };
  }
  return null;
}

function hasContent(cell: Cell | null): boolean {
  return cell !== null && !cellIsEmpty(cell);
}

function sumFormula(range: GridRange): string {
  const first = formatCellRef(range.startRow, range.startCol);
  const last = formatCellRef(range.endRow, range.endCol);
  return first === last ? `SUM(${first})` : `SUM(${first}:${last})`;
}

function readOnly(message: string): CommandResult {
  return { changes: [], ui: { type: 'toast', message: `${message} The cell was not changed.` } };
}
