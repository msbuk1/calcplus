import type {
  Cell,
  CellAddress,
  ColumnDimension,
  GridRange,
  RowDimension,
  SheetId,
  StyleId,
  Workbook,
} from '@model/types';
import {
  cellIsEmpty,
  cellStateEquals,
  cloneCell,
  getSheetById,
  setCellSnapshot,
} from '@model/types';
import { BLANK, cellValueEquals, errorValue, type CellValue } from '@model/value';
import type { CalculationScheduler, CommandChange, CommandContext, CommandResult } from './types';
import { changeCells, inverseChange } from './types';
import { createCommandRegistry } from './registry';

/** A derived result leaving the formula engine. Formula text never travels with it. */
export interface CalculationCellUpdate {
  readonly sheetId: SheetId;
  readonly row: number;
  readonly col: number;
  readonly value: CellValue;
  readonly status: 'value' | 'circular';
  /**
   * `true` when the value came from an input the engine did not evaluate.
   * The cached result is preserved and the cell is marked not recalculated.
   */
  readonly stale: boolean;
}

export interface Transaction {
  readonly seq: number;
  readonly commandId: string;
  readonly changes: readonly CommandChange[];
  readonly inverseChanges: readonly CommandChange[];
  readonly estimatedBytes: number;
}

export interface DispatchOptions {
  readonly activeSheetId?: SheetId | null;
  readonly selection?: GridRange | null;
  readonly activeCell?: CellAddress | null;
}

const MAX_TRANSACTIONS = 200;
const MAX_HISTORY_BYTES = 64 * 1024 * 1024;

/** Dimension records captured before a save, so concurrent edits stay dirty. */
export interface DirtyDimensionSnapshot {
  readonly rows: ReadonlyMap<number, RowDimension | null>;
  readonly columns: ReadonlyMap<number, ColumnDimension | null>;
}

export class DocumentSession {
  readonly workbook: Workbook;
  private readonly registry;
  private readonly undoStack: Transaction[] = [];
  private readonly redoStack: Transaction[] = [];
  private readonly dirtyCells = new Map<SheetId, Set<string>>();
  private readonly dirtyDimensions = new Map<
    SheetId,
    { readonly rows: Set<number>; readonly columns: Set<number> }
  >();
  private dirtyWorkbook = false;
  private calculationScheduler: CalculationScheduler | null = null;
  private nextSeq = 1;
  private documentVersion = 1;
  private editVersionCounter = 0;

  constructor(workbook: Workbook) {
    this.workbook = workbook;
    this.registry = createCommandRegistry();
  }

  get version(): number {
    return this.documentVersion;
  }

  /** Bumped only by commands, undo and redo; calculation results leave it alone. */
  get editVersion(): number {
    return this.editVersionCounter;
  }

  setCalculationScheduler(scheduler: CalculationScheduler | null): void {
    this.calculationScheduler = scheduler;
  }

  clearCalculationScheduler(scheduler: CalculationScheduler): void {
    if (this.calculationScheduler === scheduler) this.calculationScheduler = null;
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
    if (this.dirtyWorkbook) return true;
    for (const cells of this.dirtyCells.values()) {
      if (cells.size > 0) return true;
    }
    for (const dimensions of this.dirtyDimensions.values()) {
      if (dimensions.rows.size > 0 || dimensions.columns.size > 0) return true;
    }
    return false;
  }

  dirtyCellMap(): ReadonlyMap<SheetId, ReadonlySet<string>> {
    return this.dirtyCells;
  }

  /**
   * Capture the dimension state before a save. Rows and columns are kept
   * separately because the two record shapes have different fields.
   */
  dirtyDimensionSnapshot(): Map<SheetId, DirtyDimensionSnapshot> {
    const snapshot = new Map<SheetId, DirtyDimensionSnapshot>();
    for (const [sheetId, keys] of this.dirtyDimensions) {
      if (keys.rows.size === 0 && keys.columns.size === 0) continue;
      const sheet = getSheetById(this.workbook, sheetId);
      const rows = new Map<number, RowDimension | null>();
      for (const row of keys.rows) {
        const dimension = sheet.rowDimensions.get(row);
        rows.set(row, dimension === undefined ? null : { ...dimension });
      }
      const columns = new Map<number, ColumnDimension | null>();
      for (const column of keys.columns) {
        const dimension = sheet.columnDimensions.get(column);
        columns.set(column, dimension === undefined ? null : { ...dimension });
      }
      snapshot.set(sheetId, { rows, columns });
    }
    return snapshot;
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

  dispatch<Args>(commandId: string, args: Args, options: DispatchOptions = {}): CommandResult {
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
      activeSheetId: options.activeSheetId ?? this.workbook.sheets[0]?.id ?? null,
      selection: options.selection ?? null,
      activeCell: options.activeCell ?? null,
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
    this.editVersionCounter += 1;
    this.calculationScheduler?.afterChanges(result.changes);
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
    this.editVersionCounter += 1;
    this.calculationScheduler?.afterChanges(transaction.inverseChanges);
  }

  redo(): void {
    const transaction = this.redoStack.pop();
    if (transaction === undefined) return;
    this.applyChanges(transaction.changes, 'after');
    this.undoStack.push(transaction);
    this.trimHistory();
    this.markDirty(transaction.changes);
    this.documentVersion += 1;
    this.editVersionCounter += 1;
    this.calculationScheduler?.afterChanges(transaction.changes);
  }

  /**
   * Clear the cells and dimensions that were saved. When a snapshot is
   * supplied, an item that changed while the save was awaiting the file write
   * is compared by state and stays dirty.
   */
  markSaved(
    saved?: ReadonlyMap<SheetId, ReadonlyMap<string, Cell | null>>,
    savedDimensions?: ReadonlyMap<SheetId, DirtyDimensionSnapshot>,
  ): void {
    this.dirtyWorkbook = false;
    if (saved === undefined) {
      this.dirtyCells.clear();
    } else {
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
    }
    if (savedDimensions === undefined) {
      this.dirtyDimensions.clear();
    } else {
      for (const [sheetId, savedDims] of savedDimensions) {
        const dirty = this.dirtyDimensions.get(sheetId);
        if (dirty === undefined) continue;
        const sheet = getSheetById(this.workbook, sheetId);
        for (const [row, savedRow] of savedDims.rows) {
          if (dimensionEquals(sheet.rowDimensions.get(row) ?? null, savedRow))
            dirty.rows.delete(row);
        }
        for (const [column, savedColumn] of savedDims.columns) {
          if (dimensionEquals(sheet.columnDimensions.get(column) ?? null, savedColumn)) {
            dirty.columns.delete(column);
          }
        }
        if (dirty.rows.size === 0 && dirty.columns.size === 0) this.dirtyDimensions.delete(sheetId);
      }
    }
    this.documentVersion += 1;
  }

  /**
   * Apply derived engine results. They are not undoable transactions: the
   * authored text and the command history stay authoritative (technical design
   * §4.4). Circular results only change the in-app value and never the cached
   * result that can reach the file.
   *
   * A stale result is displayed but keeps the cached result untouched. A
   * fresh result becomes both the displayed value and the cached result
   * (file-compat §6.1).
   */
  applyCalculationResults(
    updates: readonly CalculationCellUpdate[],
    options: { readonly markDirty?: boolean } = {},
  ): number {
    let changed = 0;
    for (const update of updates) {
      const sheet = this.workbook.sheets.find((candidate) => candidate.id === update.sheetId);
      if (sheet === undefined) continue;
      const key = `${update.row}:${update.col}`;
      const cell = sheet.cells.get(key);
      if (cell === undefined) continue;
      if (cell.formula === null && cell.formulaSource === undefined) continue;

      if (cell.formulaSource !== undefined || cell.richText !== null) {
        // Shared groups, legacy arrays and rich text are never recalculated.
        // Only the derived stale marker changes; their stored value and cached
        // result stay exactly as imported.
        if (!update.stale || cell.stale === true) continue;
        setCellSnapshot(sheet, update.row, update.col, { ...cell, stale: true });
        changed += 1;
        continue;
      }

      if (update.status === 'circular') {
        const circular = errorValue('#CIRC!');
        if (cellValueEquals(cell.value, circular) && cell.stale === true) continue;
        setCellSnapshot(sheet, update.row, update.col, {
          ...cell,
          value: circular,
          stale: true,
        });
        changed += 1;
        continue;
      }

      if (update.value.type === 'error' && update.value.value === '#CIRC!') continue;

      if (update.stale) {
        if (cellValueEquals(cell.value, update.value) && cell.stale === true) continue;
        setCellSnapshot(sheet, update.row, update.col, {
          ...cell,
          value: update.value,
          stale: true,
        });
        changed += 1;
        if (options.markDirty === true) this.markCellDirty(update.sheetId, update.row, update.col);
        continue;
      }

      const next: Cell = {
        ...cell,
        value: update.value,
        savedResult: update.value,
        stale: false,
      };
      if (
        cellValueEquals(cell.value, next.value) &&
        nullableCellValueEquals(cell.savedResult, next.savedResult) &&
        cell.stale === false
      ) {
        continue;
      }
      setCellSnapshot(sheet, update.row, update.col, next);
      changed += 1;
      if (options.markDirty === true) this.markCellDirty(update.sheetId, update.row, update.col);
    }
    if (changed > 0) this.documentVersion += 1;
    return changed;
  }

  /** Cells whose displayed value is not a fully evaluated result. */
  staleCellCount(): number {
    let count = 0;
    for (const sheet of this.workbook.sheets) {
      for (const cell of sheet.cells.values()) {
        if (cell.stale === true) count += 1;
      }
    }
    return count;
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
      if (change.type === 'workbook.dateSystem') {
        this.workbook.date1904 = direction === 'after' ? change.after : change.before;
        continue;
      }
      const sheet = getSheetById(this.workbook, change.sheetId);
      if (change.type === 'dimension.set') {
        const snapshot = direction === 'after' ? change.after : change.before;
        if (change.axis === 'row') {
          if (snapshot === null) sheet.rowDimensions.delete(change.index);
          else sheet.rowDimensions.set(change.index, { ...snapshot });
        } else if (snapshot === null) {
          sheet.columnDimensions.delete(change.index);
        } else {
          sheet.columnDimensions.set(change.index, { ...snapshot });
        }
        continue;
      }
      if (change.type === 'cell.set') {
        this.applyCellSnapshots(
          sheet,
          change.range,
          direction === 'after' ? change.after : change.before,
        );
        continue;
      }
      this.applyStyleSnapshots(
        sheet,
        change.range,
        direction === 'after' ? change.after : change.before,
      );
    }
  }

  private applyCellSnapshots(
    sheet: ReturnType<typeof getSheetById>,
    range: GridRange,
    snapshots: readonly (readonly (Cell | null)[])[],
  ): void {
    for (let rowOffset = 0; rowOffset < snapshots.length; rowOffset += 1) {
      const row = snapshots[rowOffset];
      if (row === undefined) continue;
      for (let colOffset = 0; colOffset < row.length; colOffset += 1) {
        const snapshot = row[colOffset] ?? null;
        setCellSnapshot(
          sheet,
          range.startRow + rowOffset,
          range.startCol + colOffset,
          cloneCell(snapshot),
        );
      }
    }
  }

  private applyStyleSnapshots(
    sheet: ReturnType<typeof getSheetById>,
    range: GridRange,
    snapshots: readonly (readonly (StyleId | null)[])[],
  ): void {
    for (let rowOffset = 0; rowOffset < snapshots.length; rowOffset += 1) {
      const row = snapshots[rowOffset];
      if (row === undefined) continue;
      for (let colOffset = 0; colOffset < row.length; colOffset += 1) {
        const styleId = row[colOffset] ?? null;
        const rowIndex = range.startRow + rowOffset;
        const colIndex = range.startCol + colOffset;
        if (styleId === null) {
          setCellSnapshot(sheet, rowIndex, colIndex, null);
          continue;
        }
        const key = `${rowIndex}:${colIndex}`;
        const existing = sheet.cells.get(key);
        if (existing === undefined) {
          if (styleId !== 0) {
            sheet.cells.set(key, {
              value: BLANK,
              formula: null,
              formulaDerived: false,
              savedResult: null,
              styleId,
              richText: null,
              stale: false,
            });
          }
          continue;
        }
        const next: Cell = { ...existing, styleId };
        if (styleId === 0 && cellIsEmpty(next)) {
          sheet.cells.delete(key);
        } else {
          sheet.cells.set(key, next);
        }
      }
    }
  }

  private markDirty(changes: readonly CommandChange[]): void {
    for (const change of changes) {
      if (change.type === 'workbook.dateSystem') {
        this.dirtyWorkbook = true;
        continue;
      }
      if (change.type === 'dimension.set') {
        const dimensions = this.dirtyDimensions.get(change.sheetId) ?? {
          rows: new Set<number>(),
          columns: new Set<number>(),
        };
        if (change.axis === 'row') dimensions.rows.add(change.index);
        else dimensions.columns.add(change.index);
        this.dirtyDimensions.set(change.sheetId, dimensions);
        continue;
      }
      const cells = this.dirtyCells.get(change.sheetId) ?? new Set<string>();
      for (const cell of changeCells(change)) {
        cells.add(`${cell.row}:${cell.col}`);
      }
      this.dirtyCells.set(change.sheetId, cells);
    }
  }

  private markCellDirty(sheetId: SheetId, row: number, col: number): void {
    const cells = this.dirtyCells.get(sheetId) ?? new Set<string>();
    cells.add(`${row}:${col}`);
    this.dirtyCells.set(sheetId, cells);
  }
}

function dimensionEquals(
  left: RowDimension | ColumnDimension | null,
  right: RowDimension | ColumnDimension | null,
): boolean {
  if (left === null || right === null) return left === right;
  const leftRecord = left as Record<string, unknown>;
  const rightRecord = right as Record<string, unknown>;
  for (const key of new Set([...Object.keys(leftRecord), ...Object.keys(rightRecord)])) {
    if (leftRecord[key] !== rightRecord[key]) return false;
  }
  return true;
}

function nullableCellValueEquals(left: CellValue | null, right: CellValue | null): boolean {
  if (left === null || right === null) return left === right;
  return cellValueEquals(left, right);
}

function estimateBytes(changes: readonly CommandChange[]): number {
  let total = 0;
  for (const change of changes) {
    if (change.type === 'workbook.dateSystem') {
      total += 8;
      continue;
    }
    if (change.type === 'dimension.set') {
      total += 64;
      continue;
    }
    if (change.type === 'style.set') {
      total += 24;
      total += change.before.reduce((count, row) => count + row.length, 0) * 8;
      total += change.after.reduce((count, row) => count + row.length, 0) * 8;
      continue;
    }
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
