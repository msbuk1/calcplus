import { arrayFormulaForCell } from '@model/formula-ranges';
import { invalidXmlCharacterMessage } from '@model/text';
import type { Cell } from '@model/types';
import { cellStateEquals, cloneCell } from '@model/types';
import { BLANK, type CellValue } from '@model/value';
import type { CommandChange, CommandDefinition, CommandResult } from './types';

export interface CellSetArgs {
  readonly sheetId: string;
  readonly row: number;
  readonly col: number;
  readonly value: CellValue;
  /**
   * The authored formula, when the user typed one. The model keeps the exact
   * text; the engine recalculates the cell asynchronously. The cached result
   * is dropped on a formula change so a save before recalculation asks Excel
   * to recalculate (file-compat §6.1).
   */
  readonly formula?: string | null;
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
    // File-compat §9 F17 and workbook-model §9: editing a rich-text cell warns
    // and flattens it to plain text in the same transaction. The warning
    // travels with the change, so one undo restores the runs and the save
    // policy in rich-text-xml.ts keeps them on disk.
    const flattensRichText = before?.richText !== null && before?.richText !== undefined;
    if (args.value.type === 'text') {
      const unsafe = invalidXmlCharacterMessage(args.value.value);
      if (unsafe !== null) {
        return {
          changes: [],
          ui: { type: 'toast', message: `${unsafe} The cell was not changed.` },
        };
      }
    }
    const after = nextCellState(before, args.value, args.formula ?? null);
    if (cellStateEquals(before, after)) {
      return { changes: [] };
    }
    // A warned flatten is marked so the save guard lets the plain text through;
    // every other way of dropping rich runs is still refused (rich-text-xml.ts).
    const flattened =
      after !== null && flattensRichText ? { ...after, richTextFlattened: true } : after;
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
      after: [[flattened]],
    };
    return {
      changes: [change],
      selectionAfter: { ...change.range },
      ...(flattensRichText
        ? {
            ui: {
              type: 'toast' as const,
              message:
                'This cell held rich text. It was flattened to plain text; Undo restores the formatting.',
            },
          }
        : {}),
    };
  },
};

function nextCellState(before: Cell | null, value: CellValue, formula: string | null): Cell | null {
  if (formula !== null) {
    // The authored text is stored as it was typed. The formula engine
    // recalculates it asynchronously; until then the dropped cached result
    // keeps the saved file asking Excel to recalculate (file-compat §6.1).
    const unchanged = before !== null && before.formula === formula;
    return {
      value: before?.value ?? BLANK,
      formula,
      formulaDerived: false,
      savedResult: unchanged ? before.savedResult : null,
      styleId: before?.styleId ?? 0,
      richText: null,
      stale: unchanged ? (before?.stale ?? true) : true,
    };
  }
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
      stale: false,
    };
  }
  return {
    value,
    formula: null,
    formulaDerived: false,
    savedResult: null,
    styleId: before?.styleId ?? 0,
    richText: null,
    stale: false,
  };
}
