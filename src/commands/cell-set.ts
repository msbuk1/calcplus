import { arrayFormulaForCell } from '@model/formula-ranges';
import { invalidXmlCharacterMessage } from '@model/text';
import type { Cell } from '@model/types';
import { cellStateEquals, cloneCell } from '@model/types';
import type { CellValue } from '@model/value';
import type { CommandChange, CommandDefinition, CommandResult } from './types';

export interface CellSetArgs {
  readonly sheetId: string;
  readonly row: number;
  readonly col: number;
  readonly value: CellValue;
}

export const cellSetCommand: CommandDefinition<CellSetArgs> = {
  id: 'cell.set',
  label: 'Set cell',
  category: 'edit',
  keybindings: [],
  isEnabled(context) {
    return context.workbook.sheets.length > 0;
  },
  run(context, args): CommandResult {
    const sheet = context.workbook.sheets.find((candidate) => candidate.id === args.sheetId);
    if (sheet === undefined) throw new Error(`Unknown sheet: ${args.sheetId}`);
    if (!Number.isInteger(args.row) || args.row < 0 || args.row >= 1_048_576) {
      throw new RangeError('Row is outside the supported grid');
    }
    if (!Number.isInteger(args.col) || args.col < 0 || args.col >= 16_384) {
      throw new RangeError('Column is outside the supported grid');
    }

    const key = `${args.row}:${args.col}`;
    const before = cloneCell(sheet.cells.get(key) ?? null);
    if (before?.formulaSource?.kind === 'array' || arrayFormulaForCell(sheet, args.row, args.col)) {
      return {
        changes: [],
        ui: {
          type: 'toast',
          message:
            'Legacy array formulas are read-only in this milestone; the cell was not changed.',
        },
      };
    }
    if (
      before?.formulaSource?.kind === 'shared-master' ||
      before?.formulaSource?.kind === 'shared-follower'
    ) {
      return {
        changes: [],
        ui: {
          type: 'toast',
          message: 'Shared formulas are read-only in this milestone; the cell was not changed.',
        },
      };
    }
    if (before?.richText !== null && before?.richText !== undefined) {
      return {
        changes: [],
        ui: {
          type: 'toast',
          message:
            'This cell contains rich text. Rich text is read-only in this milestone; editing it would flatten its formatting, so the cell was not changed.',
        },
      };
    }
    if (args.value.type === 'text') {
      const unsafe = invalidXmlCharacterMessage(args.value.value);
      if (unsafe !== null) {
        return {
          changes: [],
          ui: { type: 'toast', message: `${unsafe} The cell was not changed.` },
        };
      }
    }
    const after = nextCellState(before, args.value);
    if (cellStateEquals(before, after)) {
      return { changes: [] };
    }
    const change: CommandChange = {
      type: 'cell.set',
      sheetId: args.sheetId,
      range: {
        startRow: args.row,
        startCol: args.col,
        endRow: args.row,
        endCol: args.col,
      },
      before: [[before]],
      after: [[after]],
    };
    return {
      changes: [change],
      selectionAfter: { ...change.range },
    };
  },
};

function nextCellState(before: Cell | null, value: CellValue): Cell | null {
  if (value.type === 'blank') {
    if (before === null) return null;
    if (before.styleId === 0) return null;
    return {
      value,
      formula: null,
      formulaDerived: false,
      savedResult: null,
      styleId: before.styleId,
      richText: null,
    };
  }
  return {
    value,
    formula: null,
    formulaDerived: false,
    savedResult: null,
    styleId: before?.styleId ?? 0,
    richText: null,
  };
}
