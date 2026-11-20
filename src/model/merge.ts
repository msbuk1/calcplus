import type { Cell, GridRange, Sheet, SheetId } from './types';
import { cellIsEmpty, cellKey, getCell, parseRangeRef } from './types';
import { ARRAY_BLOCK_MESSAGE } from './sort';

/**
 * Merge authoring (feature matrix "Merge & Center, Merge Across, Merge Cells,
 * Unmerge"; command-system §2 and §4).
 *
 * The module is pure: it reads a sheet and returns the merge list on both sides
 * of the edit, plus the covered cells whose content the merge would discard.
 * The command layer turns that plan into one `merge.set` transaction with the
 * cell clears and the alignment patch, so undo restores everything exactly.
 */

/** Excel's warning before a merge throws away data outside the anchor cells. */
export const MERGE_DATA_WARNING = 'There is data here. Do you want to merge anyway?';

/** Excel's message for an edit that would change part of a merged cell. */
export const MERGE_PARTIAL_MESSAGE =
  'Cannot change part of a merged cell. Select the whole merged range or unmerge it first.';

/** Merge Across cannot turn one multi-row merge into a merge per row. */
export const MERGE_ACROSS_SPAN_MESSAGE =
  'Merge Across cannot split a merged cell that spans more than one row. Unmerge it first.';

/** Merge Across needs width; a one-column selection has nothing to merge. */
export const MERGE_ACROSS_COLUMN_MESSAGE =
  'Merge Across needs a selection that is more than one column wide.';

export const MERGE_SINGLE_CELL_MESSAGE = 'Select more than one cell to merge.';

export const UNMERGE_NOTHING_MESSAGE = 'The selection does not contain a merged cell.';

/**
 * Shared formula groups are read-only in v1 (file-compat §6.1). A merge would
 * clear the covered cells, and the save guard would refuse that later; the
 * merge blocks up front with the same reason instead.
 */
export const SHARED_FORMULA_BLOCK_MESSAGE =
  'You cannot change part of a shared formula. Shared formulas are read-only in this version.';

export type MergeMode = 'center' | 'across' | 'cells';

export type MergeBlockCode =
  'merged-partial' | 'merged-across' | 'array-formula' | 'shared-formula' | 'invalid-selection';

/** A covered cell whose value, formula or rich text the merge discards. */
export interface MergeDiscardedCell {
  readonly row: number;
  readonly col: number;
  readonly cell: Cell;
}

/** The full before/after merge lists plus the data the merge would clear. */
export interface MergePlan {
  readonly sheetId: SheetId;
  readonly range: GridRange;
  readonly mode: MergeMode;
  readonly mergesBefore: readonly GridRange[];
  readonly mergesAfter: readonly GridRange[];
  /** Non-anchor cells that hold content; empty when the merge loses nothing. */
  readonly discarded: readonly MergeDiscardedCell[];
}

export type MergePlanResult =
  | { readonly kind: 'planned'; readonly plan: MergePlan }
  | {
      readonly kind: 'blocked';
      readonly code: MergeBlockCode;
      readonly message: string;
      readonly range: GridRange;
    };

export interface UnmergePlan {
  readonly sheetId: SheetId;
  readonly range: GridRange;
  readonly mergesBefore: readonly GridRange[];
  readonly mergesAfter: readonly GridRange[];
}

export function normalizeRange(range: GridRange): GridRange {
  return {
    startRow: Math.min(range.startRow, range.endRow),
    startCol: Math.min(range.startCol, range.endCol),
    endRow: Math.max(range.startRow, range.endRow),
    endCol: Math.max(range.startCol, range.endCol),
  };
}

export function rangesIntersect(left: GridRange, right: GridRange): boolean {
  return (
    left.startRow <= right.endRow &&
    right.startRow <= left.endRow &&
    left.startCol <= right.endCol &&
    right.startCol <= left.endCol
  );
}

/** The merge list identity that lets a no-op plan skip its `merge.set`. */
export function mergeListsEqual(left: readonly GridRange[], right: readonly GridRange[]): boolean {
  if (left.length !== right.length) return false;
  const leftKeys = left.map(rangeKey).sort();
  const rightKeys = right.map(rangeKey).sort();
  return leftKeys.every((key, index) => key === rightKeys[index]);
}

/**
 * Plan one of the three merge commands. The selection may contain existing
 * merges only when each is wholly inside it, so the merge can replace them
 * without splitting a merged cell. Merge Across additionally refuses a merge
 * that spans rows because it would have to become several merges.
 */
export function planMerge(sheet: Sheet, selection: GridRange, mode: MergeMode): MergePlanResult {
  const range = normalizeRange(selection);
  const rows = range.endRow - range.startRow + 1;
  const columns = range.endCol - range.startCol + 1;
  if (rows * columns <= 1) {
    return blocked('invalid-selection', MERGE_SINGLE_CELL_MESSAGE, range);
  }
  if (mode === 'across' && columns === 1) {
    return blocked('invalid-selection', MERGE_ACROSS_COLUMN_MESSAGE, range);
  }

  const array = findArrayFormulaBlock(sheet, range);
  if (array !== null) {
    return blocked('array-formula', ARRAY_BLOCK_MESSAGE, range);
  }
  if (findSharedFormulaBlock(sheet, range)) {
    return blocked('shared-formula', SHARED_FORMULA_BLOCK_MESSAGE, range);
  }

  const intersecting = sheet.merges.filter((merge) => rangesIntersect(merge.range, range));
  for (const merge of intersecting) {
    if (!rangeContainsRange(range, merge.range)) {
      return blocked('merged-partial', MERGE_PARTIAL_MESSAGE, range);
    }
    if (mode === 'across' && merge.range.startRow !== merge.range.endRow) {
      return blocked('merged-across', MERGE_ACROSS_SPAN_MESSAGE, range);
    }
  }

  const mergesBefore = sheet.merges.map((merge) => ({ ...merge.range }));
  const kept = mergesBefore.filter((candidate) => !rangesIntersect(candidate, range));
  const created = mergesForMode(range, mode);
  const mergesAfter = [...kept, ...created].sort(compareRanges);

  const anchors = new Set(created.map((merge) => cellKey(merge.startRow, merge.startCol)));
  const discarded: MergeDiscardedCell[] = [];
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    for (let col = range.startCol; col <= range.endCol; col += 1) {
      if (anchors.has(cellKey(row, col))) continue;
      const cell = getCell(sheet, row, col);
      if (cell !== null && !cellIsEmpty(cell)) discarded.push({ row, col, cell });
    }
  }

  return {
    kind: 'planned',
    plan: {
      sheetId: sheet.id,
      range,
      mode,
      mergesBefore,
      mergesAfter,
      discarded,
    },
  };
}

/**
 * Plan an unmerge: every merge that touches the selection is removed, whether
 * it is wholly inside the selection or merely overlaps an edge. Content is not
 * moved; the anchor already holds it.
 */
export function planUnmerge(sheet: Sheet, selection: GridRange): UnmergePlan {
  const range = normalizeRange(selection);
  const mergesBefore = sheet.merges.map((merge) => ({ ...merge.range }));
  const mergesAfter = mergesBefore.filter((candidate) => !rangesIntersect(candidate, range));
  return { sheetId: sheet.id, range, mergesBefore, mergesAfter };
}

function mergesForMode(range: GridRange, mode: MergeMode): GridRange[] {
  if (mode !== 'across') return [range];
  const merges: GridRange[] = [];
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    merges.push({
      startRow: row,
      startCol: range.startCol,
      endRow: row,
      endCol: range.endCol,
    });
  }
  return merges;
}

function findArrayFormulaBlock(sheet: Sheet, range: GridRange): GridRange | null {
  for (const cell of sheet.cells.values()) {
    const source = cell.formulaSource;
    if (source?.kind !== 'array' || source.range === undefined) continue;
    try {
      const covered = parseRangeRef(source.range.replaceAll('$', ''));
      if (rangesIntersect(covered, range)) return covered;
    } catch {
      // A malformed ref stays unmodelled; the worksheet patch preserves it.
    }
  }
  return null;
}

/**
 * A shared group is stored as a master plus followers, and every member is
 * read-only in v1. Clearing one would be refused by the save guard, so the
 * merge is blocked here before any cell patch is built.
 */
function findSharedFormulaBlock(sheet: Sheet, range: GridRange): boolean {
  for (const [key, cell] of sheet.cells) {
    const kind = cell.formulaSource?.kind;
    if (kind !== 'shared-master' && kind !== 'shared-follower') continue;
    const separator = key.indexOf(':');
    const row = Number(key.slice(0, separator));
    const col = Number(key.slice(separator + 1));
    if (
      row >= range.startRow &&
      row <= range.endRow &&
      col >= range.startCol &&
      col <= range.endCol
    ) {
      return true;
    }
  }
  return false;
}

function rangeContainsRange(outer: GridRange, inner: GridRange): boolean {
  return (
    inner.startRow >= outer.startRow &&
    inner.endRow <= outer.endRow &&
    inner.startCol >= outer.startCol &&
    inner.endCol <= outer.endCol
  );
}

function compareRanges(left: GridRange, right: GridRange): number {
  return (
    left.startRow - right.startRow ||
    left.startCol - right.startCol ||
    left.endRow - right.endRow ||
    left.endCol - right.endCol
  );
}

function rangeKey(range: GridRange): string {
  return `${range.startRow}:${range.startCol}:${range.endRow}:${range.endCol}`;
}

function blocked(code: MergeBlockCode, message: string, range: GridRange): MergePlanResult {
  return { kind: 'blocked', code, message, range };
}
