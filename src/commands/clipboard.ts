import {
  MAX_CLIPBOARD_CELLS,
  captureClipboard,
  clipboardCellCount,
  translateFormulaForPaste,
  type ClipboardCell,
  type ClipboardSnapshot,
} from '@model/clipboard';
import type { Cell, GridRange, Sheet, SheetId } from '@model/types';
import { cellStateEquals, cloneCell, getCell } from '@model/types';
import { BLANK, cloneCellValue } from '@model/value';
import type { CommandChange, CommandContext, CommandDefinition, CommandResult } from './types';

export interface ClipboardSourceArgs {
  readonly sheetId?: SheetId;
  readonly range?: GridRange;
}

export interface ClipboardPasteArgs {
  readonly sheetId?: SheetId;
  readonly target?: GridRange;
  readonly snapshot: ClipboardSnapshot;
}

const copyCommand: CommandDefinition<ClipboardSourceArgs> = {
  id: 'edit.copy',
  label: 'Copy',
  category: 'edit',
  keybindings: ['Ctrl+C'],
  isEnabled: hasSheet,
  run(context, args = {}) {
    return capture(context, args, 'copy');
  },
};

const cutCommand: CommandDefinition<ClipboardSourceArgs> = {
  id: 'edit.cut',
  label: 'Cut',
  category: 'edit',
  keybindings: ['Ctrl+X'],
  isEnabled: hasSheet,
  run(context, args = {}) {
    return capture(context, args, 'cut');
  },
};

/**
 * A paste is one transaction: the target cells, a cut's cleared source cells
 * and the merged-range patch all travel in one `CommandResult`. A copied
 * formula is shifted by the paste offset; a cut is a move, so its formula text
 * is kept unchanged.
 */
const pasteCommand: CommandDefinition<ClipboardPasteArgs> = {
  id: 'edit.paste',
  label: 'Paste',
  category: 'edit',
  keybindings: ['Ctrl+V'],
  isEnabled: hasSheet,
  run(context, args) {
    const sheetId = args.sheetId ?? context.activeSheetId;
    const sheet =
      sheetId === null
        ? undefined
        : context.workbook.sheets.find((candidate) => candidate.id === sheetId);
    if (sheet === undefined) return toast('The paste target sheet is not open.');
    const target = args.target ?? context.selection ?? rangeOfActiveCell(context);
    if (target === null || !validRange(target)) {
      return toast('Select a cell or range to paste into.');
    }
    const snapshot = args.snapshot;
    if (snapshot.cells.length === 0 || snapshot.cells[0]?.length === 0) {
      return toast('The clipboard is empty.');
    }
    const height = snapshot.range.endRow - snapshot.range.startRow + 1;
    const width = snapshot.range.endCol - snapshot.range.startCol + 1;
    if (snapshot.cells.length !== height || snapshot.cells.some((row) => row.length !== width)) {
      return toast('The clipboard snapshot is malformed and was not pasted.');
    }
    // A copy repeats across a larger selection. Each destination axis is at
    // least the copied size, so a 1x8 row pasted into A2:A10 fills A2:H10
    // rather than clipping to column A. When the selection is larger, the
    // pattern tiles by modulo in both directions, including sizes where the
    // repeat does not divide evenly. A single-cell target or an exactly
    // matching target keeps the copied size. The destination rectangle becomes
    // the paste area, so the fill is still one cell.set change and one undo
    // transaction. A cut is a move and never tiles.
    const targetHeight = target.endRow - target.startRow + 1;
    const targetWidth = target.endCol - target.startCol + 1;
    const tile = snapshot.mode === 'copy' && targetCells(target) > 1;
    const destination: GridRange = tile
      ? {
          startRow: target.startRow,
          startCol: target.startCol,
          endRow: target.startRow + Math.max(height, targetHeight) - 1,
          endCol: target.startCol + Math.max(width, targetWidth) - 1,
        }
      : {
          startRow: target.startRow,
          startCol: target.startCol,
          endRow: target.startRow + height - 1,
          endCol: target.startCol + width - 1,
        };
    if (clipboardCellCount(destination) > MAX_CLIPBOARD_CELLS) {
      return toast(
        `Select at most ${MAX_CLIPBOARD_CELLS.toLocaleString('en-US')} cells before pasting.`,
      );
    }
    if (destination.endRow >= 1_048_576 || destination.endCol >= 16_384) {
      return toast('The pasted range would leave the sheet; the paste was not applied.');
    }
    if (snapshot.mode === 'cut' && targetCells(target) > 1) {
      if (targetHeight !== height || targetWidth !== width) {
        return toast('To move cells, select a single cell or an area the same size as the cut.');
      }
    }
    const sourceSheet =
      snapshot.mode === 'cut'
        ? context.workbook.sheets.find((candidate) => candidate.id === snapshot.sheetId)
        : undefined;
    if (snapshot.mode === 'cut' && sourceSheet === undefined) {
      return toast('The sheet the cut came from is no longer open.');
    }
    const sameSheet = sourceSheet?.id === sheet.id;
    if (snapshot.mode === 'cut' && sameSheet && sameRange(destination, snapshot.range)) {
      return { changes: [], selectionAfter: destination };
    }

    const targetChange = buildTargetChange(sheet, snapshot, destination, snapshot.mode === 'copy');
    const sourceChange =
      snapshot.mode === 'cut' &&
      sourceSheet !== undefined &&
      (!sameSheet || !sameRange(destination, snapshot.range))
        ? buildSourceClearChange(sourceSheet, snapshot.range)
        : null;
    const mergeChanges = buildMergeChanges(sheet, sourceSheet, snapshot, destination, sameSheet);

    const changes: CommandChange[] = [];
    if (sourceChange !== null) changes.push(sourceChange);
    changes.push(...mergeChanges.source);
    if (targetChange !== null) changes.push(targetChange);
    changes.push(...mergeChanges.target);
    return { changes, selectionAfter: destination };
  },
};

export function clipboardCommands(): readonly CommandDefinition<never>[] {
  return [
    copyCommand as unknown as CommandDefinition<never>,
    cutCommand as unknown as CommandDefinition<never>,
    pasteCommand as unknown as CommandDefinition<never>,
  ];
}

function capture(
  context: CommandContext,
  args: ClipboardSourceArgs,
  mode: 'copy' | 'cut',
): CommandResult {
  const sheetId = args.sheetId ?? context.activeSheetId;
  const sheet =
    sheetId === null
      ? undefined
      : context.workbook.sheets.find((candidate) => candidate.id === sheetId);
  if (sheet === undefined) return toast('Open a worksheet before copying.');
  const range = args.range ?? context.selection ?? rangeOfActiveCell(context);
  if (range === null || !validRange(range)) return toast('Select a cell or range first.');
  const count = clipboardCellCount(range);
  if (count > MAX_CLIPBOARD_CELLS) {
    return toast(
      `Select at most ${MAX_CLIPBOARD_CELLS.toLocaleString('en-US')} cells before copying.`,
    );
  }
  return { changes: [], clipboard: captureClipboard(sheet, range, mode) };
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

function targetCells(range: GridRange): number {
  return (range.endRow - range.startRow + 1) * (range.endCol - range.startCol + 1);
}

function buildTargetChange(
  sheet: Sheet,
  snapshot: ClipboardSnapshot,
  destination: GridRange,
  adjustFormula: boolean,
): CommandChange | null {
  const patternHeight = snapshot.cells.length;
  const patternWidth = snapshot.cells[0]?.length ?? 0;
  const before: (Cell | null)[][] = [];
  const after: (Cell | null)[][] = [];
  let changed = false;
  for (let row = destination.startRow; row <= destination.endRow; row += 1) {
    const beforeRow: (Cell | null)[] = [];
    const afterRow: (Cell | null)[] = [];
    // The destination is at least the copied block. When it is larger, modulo
    // repeats the pattern in both directions from the target top-left.
    const rowOffset = (row - destination.startRow) % patternHeight;
    const snapshotRow = snapshot.cells[rowOffset] ?? [];
    for (let col = destination.startCol; col <= destination.endCol; col += 1) {
      const colOffset = (col - destination.startCol) % patternWidth;
      const existing = getCell(sheet, row, col);
      // Each source cell sits at its own position in the copied pattern, so a
      // relative formula shifts by the distance from that source cell to the
      // cell it is pasted into. A repeated tile therefore shifts again for
      // every row and column it fills.
      const pasted = pastedCell(
        snapshotRow[colOffset] ?? null,
        adjustFormula ? row - (snapshot.range.startRow + rowOffset) : 0,
        adjustFormula ? col - (snapshot.range.startCol + colOffset) : 0,
      );
      beforeRow.push(cloneCell(existing));
      afterRow.push(pasted);
      if (!cellStateEquals(existing, pasted)) changed = true;
    }
    before.push(beforeRow);
    after.push(afterRow);
  }
  return changed
    ? {
        type: 'cell.set',
        sheetId: sheet.id,
        range: destination,
        before,
        after,
      }
    : null;
}

function pastedCell(source: ClipboardCell | null, rowDelta: number, colDelta: number): Cell | null {
  if (source === null) return null;
  if (source.formula !== null) {
    return {
      value: BLANK,
      formula: translateFormulaForPaste(source.formula, rowDelta, colDelta),
      formulaDerived: false,
      savedResult: null,
      styleId: source.styleId,
      richText: null,
      stale: true,
    };
  }
  if (source.value.type === 'blank' && source.styleId === 0) return null;
  return {
    value: cloneCellValue(source.value),
    formula: null,
    formulaDerived: false,
    savedResult: null,
    styleId: source.styleId,
    richText: null,
    stale: false,
  };
}

function buildSourceClearChange(sheet: Sheet, range: GridRange): CommandChange | null {
  const before: (Cell | null)[][] = [];
  const after: (Cell | null)[][] = [];
  let changed = false;
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    const beforeRow: (Cell | null)[] = [];
    const afterRow: (Cell | null)[] = [];
    for (let col = range.startCol; col <= range.endCol; col += 1) {
      const existing = getCell(sheet, row, col);
      beforeRow.push(cloneCell(existing));
      afterRow.push(null);
      if (existing !== null) changed = true;
    }
    before.push(beforeRow);
    after.push(afterRow);
  }
  return changed
    ? {
        type: 'cell.set',
        sheetId: sheet.id,
        range,
        before,
        after,
      }
    : null;
}

function buildMergeChanges(
  targetSheet: Sheet,
  sourceSheet: Sheet | undefined,
  snapshot: ClipboardSnapshot,
  destination: GridRange,
  sameSheet: boolean,
): { readonly source: CommandChange[]; readonly target: CommandChange[] } {
  const source: CommandChange[] = [];
  const target: CommandChange[] = [];
  if (snapshot.mode === 'cut' && sourceSheet !== undefined && !sameSheet) {
    const cleared = buildMergeClearChange(sourceSheet, snapshot.range);
    if (cleared !== null) source.push(cleared);
  }
  const merged = buildMergeChange(
    targetSheet,
    snapshot,
    destination,
    sameSheet && snapshot.mode === 'cut',
  );
  if (merged !== null) target.push(merged);
  return { source, target };
}

function buildMergeChange(
  sheet: Sheet,
  snapshot: ClipboardSnapshot,
  destination: GridRange,
  removeSource: boolean,
): CommandChange | null {
  const before = sheet.merges.map((merge) => merge.range);
  const after = planMerges(before, snapshot, destination, removeSource);
  if (mergeListsEqual(before, after)) return null;
  return {
    type: 'merge.set',
    sheetId: sheet.id,
    before: before.map((range) => ({ ...range })),
    after,
  };
}

function buildMergeClearChange(sheet: Sheet, range: GridRange): CommandChange | null {
  const before = sheet.merges.map((merge) => merge.range);
  const after = before.filter((candidate) => !rangesIntersect(candidate, range));
  if (mergeListsEqual(before, after)) return null;
  return {
    type: 'merge.set',
    sheetId: sheet.id,
    before: before.map((candidate) => ({ ...candidate })),
    after,
  };
}

function planMerges(
  current: readonly GridRange[],
  snapshot: ClipboardSnapshot,
  destination: GridRange,
  removeSource: boolean,
): GridRange[] {
  const kept = current
    .filter((range) => !rangesIntersect(range, destination))
    .filter((range) => !removeSource || !rangesIntersect(range, snapshot.range));
  const rowDelta = destination.startRow - snapshot.range.startRow;
  const colDelta = destination.startCol - snapshot.range.startCol;
  for (const merge of snapshot.merges) {
    const shifted: GridRange = {
      startRow: merge.startRow + rowDelta,
      startCol: merge.startCol + colDelta,
      endRow: merge.endRow + rowDelta,
      endCol: merge.endCol + colDelta,
    };
    if (!kept.some((range) => sameRange(range, shifted))) kept.push(shifted);
  }
  return kept.sort(compareRanges);
}

function rangesIntersect(left: GridRange, right: GridRange): boolean {
  return !(
    left.endRow < right.startRow ||
    right.endRow < left.startRow ||
    left.endCol < right.startCol ||
    right.endCol < left.startCol
  );
}

function sameRange(left: GridRange, right: GridRange): boolean {
  return (
    left.startRow === right.startRow &&
    left.startCol === right.startCol &&
    left.endRow === right.endRow &&
    left.endCol === right.endCol
  );
}

function mergeListsEqual(left: readonly GridRange[], right: readonly GridRange[]): boolean {
  if (left.length !== right.length) return false;
  const leftKeys = left.map(rangeKey).sort();
  const rightKeys = right.map(rangeKey).sort();
  return leftKeys.every((key, index) => key === rightKeys[index]);
}

function rangeKey(range: GridRange): string {
  return `${range.startRow}:${range.startCol}:${range.endRow}:${range.endCol}`;
}

function compareRanges(left: GridRange, right: GridRange): number {
  return (
    left.startRow - right.startRow ||
    left.startCol - right.startCol ||
    left.endRow - right.endRow ||
    left.endCol - right.endCol
  );
}

function toast(message: string): CommandResult {
  return { changes: [], ui: { type: 'toast', message } };
}
