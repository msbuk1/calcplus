import type { Cell, CellAddress, SheetId, Workbook } from '@model/types';
import { cellStateEquals, cloneCell, getSheetById, setCellSnapshot } from '@model/types';
import type { CommandChange, CommandContext, CommandResult } from './types';
import { changeCells, inverseChange } from './types';
import { createCommandRegistry } from './registry';

export interface Transaction {
  readonly seq: number;
  readonly commandId: string;
  readonly changes: readonly CommandChange[];
  readonly inverseChanges: readonly CommandChange[];
  readonly estimatedBytes: number;
}

const MAX_TRANSACTIONS = 200;
const MAX_HISTORY_BYTES = 64 * 1024 * 1024;

export class DocumentSession {
  readonly workbook: Workbook;
  private readonly registry;
  private readonly undoStack: Transaction[] = [];
  private readonly redoStack: Transaction[] = [];
  private readonly dirtyCells = new Map<SheetId, Set<string>>();
  private nextSeq = 1;
  private documentVersion = 1;

  constructor(workbook: Workbook) {
    this.workbook = workbook;
    this.registry = createCommandRegistry();
  }

  get version(): number {
    return this.documentVersion;
  }

  get canUndo(): boolean {
    return this.undoStack.length > 0;
  }

  get canRedo(): boolean {
    return this.redoStack.length > 0;
  }

  get transactionCount(): number {
    return this.undoStack.length;
  }

  get historyBytes(): number {
    return [...this.undoStack, ...this.redoStack].reduce(
      (total, transaction) => total + transaction.estimatedBytes,
      0,
    );
  }

  get hasUnsavedChanges(): boolean {
    for (const cells of this.dirtyCells.values()) {
      if (cells.size > 0) return true;
    }
    return false;
  }

  dirtyCellMap(): ReadonlyMap<SheetId, ReadonlySet<string>> {
    return this.dirtyCells;
  }

  /**
   * Capture the current dirty cell state. Saves use this snapshot so edits
   * made while a file write is in flight stay dirty afterwards.
   */
  dirtyCellSnapshot(): Map<SheetId, Map<string, Cell | null>> {
    const snapshot = new Map<SheetId, Map<string, Cell | null>>();
    for (const [sheetId, keys] of this.dirtyCells) {
      if (keys.size === 0) continue;
      const sheet = getSheetById(this.workbook, sheetId);
      const cells = new Map<string, Cell | null>();
      for (const key of keys) cells.set(key, cloneCell(sheet.cells.get(key) ?? null));
      snapshot.set(sheetId, cells);
    }
    return snapshot;
  }

  dispatch<Args>(commandId: string, args: Args): CommandResult {
    if (commandId === 'edit.undo') {
      this.undo();
      return { changes: [] };
    }
    if (commandId === 'edit.redo') {
      this.redo();
      return { changes: [] };
    }
    const command = this.registry.require(commandId);
    const context: CommandContext = {
      workbook: this.workbook,
      selection: null,
      activeCell: null,
    };
    if (!command.isEnabled(context)) {
      throw new Error(`Command is disabled: ${commandId}`);
    }
    const result = (command.run as (context: CommandContext, args: unknown) => CommandResult)(
      context,
      args,
    );
    if (result.changes.length === 0) return result;
    this.applyChanges(result.changes, 'after');
    this.record(commandId, result.changes);
    this.documentVersion += 1;
    return result;
  }

  undo(): void {
    const transaction = this.undoStack.pop();
    if (transaction === undefined) return;
    this.applyChanges([...transaction.inverseChanges].reverse(), 'after');
    this.redoStack.push(transaction);
    this.trimHistory();
    this.markDirty(transaction.changes);
    this.documentVersion += 1;
  }

  redo(): void {
    const transaction = this.redoStack.pop();
    if (transaction === undefined) return;
    this.applyChanges(transaction.changes, 'after');
    this.undoStack.push(transaction);
    this.trimHistory();
    this.markDirty(transaction.changes);
    this.documentVersion += 1;
  }

  /**
   * Clear the cells that were saved. When a snapshot is supplied, a cell that
   * changed while the save was awaiting the file write is compared by state
   * and stays dirty.
   */
  markSaved(saved?: ReadonlyMap<SheetId, ReadonlyMap<string, Cell | null>>): void {
    if (saved === undefined) {
      this.dirtyCells.clear();
      this.documentVersion += 1;
      return;
    }
    for (const [sheetId, savedCells] of saved) {
      const dirty = this.dirtyCells.get(sheetId);
      if (dirty === undefined) continue;
      const sheet = getSheetById(this.workbook, sheetId);
      for (const [key, savedCell] of savedCells) {
        const current = sheet.cells.get(key) ?? null;
        if (cellStateEquals(current, savedCell)) dirty.delete(key);
      }
      if (dirty.size === 0) this.dirtyCells.delete(sheetId);
    }
    this.documentVersion += 1;
  }

  private record(commandId: string, changes: readonly CommandChange[]): void {
    const transaction: Transaction = {
      seq: this.nextSeq,
      commandId,
      changes,
      inverseChanges: changes.map(inverseChange),
      estimatedBytes: estimateBytes(changes),
    };
    this.nextSeq += 1;
    this.undoStack.push(transaction);
    this.redoStack.length = 0;
    this.trimHistory();
    this.markDirty(changes);
  }

  private trimHistory(): void {
    while (this.undoStack.length > MAX_TRANSACTIONS) this.undoStack.shift();
    while (this.redoStack.length > MAX_TRANSACTIONS) this.redoStack.shift();
    while (
      this.undoStack.length + this.redoStack.length > 1 &&
      this.historyBytes > MAX_HISTORY_BYTES
    ) {
      if (this.redoStack.length > 0) this.redoStack.shift();
      else this.undoStack.shift();
    }
  }

  private applyChanges(changes: readonly CommandChange[], direction: 'before' | 'after'): void {
    for (const change of changes) {
      const sheet = getSheetById(this.workbook, change.sheetId);
      const snapshots = direction === 'after' ? change.after : change.before;
      for (let rowOffset = 0; rowOffset < snapshots.length; rowOffset += 1) {
        const row = snapshots[rowOffset];
        if (row === undefined) continue;
        for (let colOffset = 0; colOffset < row.length; colOffset += 1) {
          const snapshot = row[colOffset] ?? null;
          setCellSnapshot(
            sheet,
            change.range.startRow + rowOffset,
            change.range.startCol + colOffset,
            cloneCell(snapshot),
          );
        }
      }
    }
  }

  private markDirty(changes: readonly CommandChange[]): void {
    for (const change of changes) {
      const cells = this.dirtyCells.get(change.sheetId) ?? new Set<string>();
      for (const cell of changeCells(change)) {
        cells.add(`${cell.row}:${cell.col}`);
      }
      this.dirtyCells.set(change.sheetId, cells);
    }
  }
}

function estimateBytes(changes: readonly CommandChange[]): number {
  let total = 0;
  for (const change of changes) {
    for (const row of change.before) {
      for (const cell of row) total += estimateCell(cell) + 48;
    }
    for (const row of change.after) {
      for (const cell of row) total += estimateCell(cell) + 48;
    }
  }
  return total;
}

function estimateCell(cell: Cell | null): number {
  if (cell === null) return 8;
  let bytes = 64;
  if (cell.formula !== null) bytes += cell.formula.length * 2;
  if (cell.richText !== null) bytes += cell.richText.xml.length * 2;
  if (cell.value.type === 'text') bytes += cell.value.value.length * 2;
  if (cell.savedResult !== null && cell.savedResult.type === 'text') {
    bytes += cell.savedResult.value.length * 2;
  }
  return bytes;
}

export function singleCellAddress(row: number, col: number): CellAddress {
  return { row, col };
}
