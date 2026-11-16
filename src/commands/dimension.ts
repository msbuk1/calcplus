import type { SheetId } from '@model/types';
import type { CommandDefinition, CommandResult, DimensionRecord } from './types';

/**
 * Row-height and column-width changes, including the `customHeight`,
 * `customWidth` and `bestFit` flags a resize or autofit records. Autofit
 * measurement happens in the grid, which owns the font metrics; this command
 * stores the resulting document dimension as one undoable transaction.
 */

export interface DimensionSetArgs {
  readonly sheetId?: SheetId;
  readonly axis: 'row' | 'column';
  readonly index: number;
  /** Points for a row, Excel character units for a column. */
  readonly size?: number;
  readonly hidden?: boolean;
  readonly custom?: boolean;
  readonly bestFit?: boolean;
}

const MAX_ROWS = 1_048_576;
const MAX_COLUMNS = 16_384;

type MutableDimension = {
  -readonly [Key in keyof DimensionRecord]: DimensionRecord[Key];
};

export const dimensionSetCommand: CommandDefinition<DimensionSetArgs> = {
  id: 'dimension.set',
  label: 'Row height and column width',
  category: 'format',
  keybindings: [],
  isEnabled(context) {
    return context.workbook.sheets.length > 0;
  },
  run(context, args): CommandResult {
    const sheetId = args.sheetId ?? context.activeSheetId;
    if (sheetId === null) return { changes: [] };
    const sheet = context.workbook.sheets.find((candidate) => candidate.id === sheetId);
    if (sheet === undefined) return { changes: [] };
    const limit = args.axis === 'row' ? MAX_ROWS : MAX_COLUMNS;
    if (!Number.isInteger(args.index) || args.index < 0 || args.index >= limit) {
      return { changes: [] };
    }
    if (args.size !== undefined && (!Number.isFinite(args.size) || args.size < 0)) {
      return { changes: [] };
    }
    const dimensions = args.axis === 'row' ? sheet.rowDimensions : sheet.columnDimensions;
    const current = dimensions.get(args.index);
    const before: DimensionRecord | null = current === undefined ? null : { ...current };
    const after: MutableDimension = { ...(before ?? {}) };
    if (args.size !== undefined) {
      if (args.axis === 'row') after.height = args.size;
      else after.width = args.size;
    }
    if (args.hidden !== undefined) after.hidden = args.hidden;
    if (args.custom !== undefined) {
      if (args.axis === 'row') after.customHeight = args.custom;
      else after.customWidth = args.custom;
    }
    if (args.axis === 'column' && args.bestFit !== undefined) after.bestFit = args.bestFit;
    if (sameDimension(before, after)) return { changes: [] };
    return {
      changes: [
        {
          type: 'dimension.set',
          sheetId,
          axis: args.axis,
          index: args.index,
          before,
          after,
        },
      ],
    };
  },
};

function sameDimension(before: DimensionRecord | null, after: DimensionRecord): boolean {
  if (before === null) return Object.keys(after).length === 0;
  const keys = new Set([...Object.keys(before), ...Object.keys(after)]);
  for (const key of keys) {
    if ((before as Record<string, unknown>)[key] !== (after as Record<string, unknown>)[key]) {
      return false;
    }
  }
  return true;
}
