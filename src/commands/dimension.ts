import type { SheetId } from '@model/types';
import type { CommandChange, CommandDefinition, CommandResult, DimensionRecord } from './types';

/**
 * Row-height and column-width changes, including the `customHeight`,
 * `customWidth` and `bestFit` flags a resize or autofit records. Autofit
 * measurement happens in the grid, which owns the font metrics; this command
 * stores the resulting document dimensions as one undoable transaction.
 */

export interface DimensionSetArgs {
  readonly sheetId?: SheetId;
  readonly axis: 'row' | 'column';
  readonly index?: number;
  /** Points for a row, Excel character units for a column. */
  readonly size?: number;
  readonly hidden?: boolean;
  readonly custom?: boolean;
  readonly bestFit?: boolean;
  /**
   * A multi-selection resize changes several dimensions in one transaction.
   * Each entry carries its own index, size and imported hidden flag.
   */
  readonly entries?: readonly DimensionSetEntry[];
}

export interface DimensionSetEntry {
  readonly index: number;
  /** Points for a row, Excel character units for a column. */
  readonly size?: number;
  readonly hidden?: boolean;
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
    const entries = dimensionEntries(args);
    if (entries === null) return { changes: [] };
    const limit = args.axis === 'row' ? MAX_ROWS : MAX_COLUMNS;
    for (const entry of entries) {
      if (!Number.isInteger(entry.index) || entry.index < 0 || entry.index >= limit) {
        return { changes: [] };
      }
      if (entry.size !== undefined && (!Number.isFinite(entry.size) || entry.size < 0)) {
        return { changes: [] };
      }
    }
    const dimensions = args.axis === 'row' ? sheet.rowDimensions : sheet.columnDimensions;
    const changes: CommandChange[] = [];
    for (const entry of entries) {
      const current = dimensions.get(entry.index);
      const before: DimensionRecord | null = current === undefined ? null : { ...current };
      const after: MutableDimension = { ...(before ?? {}) };
      if (entry.size !== undefined) {
        if (args.axis === 'row') after.height = entry.size;
        else after.width = entry.size;
      }
      if (entry.hidden !== undefined) after.hidden = entry.hidden;
      if (args.custom !== undefined) {
        if (args.axis === 'row') after.customHeight = args.custom;
        else after.customWidth = args.custom;
      }
      if (args.axis === 'column' && args.bestFit !== undefined) after.bestFit = args.bestFit;
      if (sameDimension(before, after)) continue;
      changes.push({
        type: 'dimension.set',
        sheetId,
        axis: args.axis,
        index: entry.index,
        before,
        after,
      });
    }
    return { changes };
  },
};

/**
 * The dimensions a command call changes. A caller may use the legacy single
 * `index`/`size` pair or pass `entries` for a multi-header selection. A
 * duplicated index keeps its first occurrence so one transaction cannot carry
 * two competing before/after snapshots for the same dimension.
 */
function dimensionEntries(args: DimensionSetArgs): readonly DimensionSetEntry[] | null {
  if (args.entries === undefined) {
    if (args.index === undefined) return null;
    return [{ index: args.index, size: args.size, hidden: args.hidden }];
  }
  const seen = new Set<number>();
  const entries: DimensionSetEntry[] = [];
  for (const entry of args.entries) {
    if (seen.has(entry.index)) continue;
    seen.add(entry.index);
    entries.push(entry);
  }
  return entries.length === 0 ? null : entries;
}

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
