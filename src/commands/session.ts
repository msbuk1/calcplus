import type {
  Cell,
  CellAddress,
  ColumnDimension,
  GridRange,
  RowDimension,
  Sheet,
  SheetId,
  StyleId,
  Workbook,
} from '@model/types';
import { cellStateEquals, cloneCell, getSheetById, setCellSnapshot } from '@model/types';
import { findSheet } from '@model/workbook';
import { BLANK, cellValueEquals, errorValue, type CellValue } from '@model/value';
import { isSheetVisible } from '@view/sheet-visibility';
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

/** One command inside a batch. Commands run in order and share one context. */
export interface DispatchBatchEntry {
  readonly id: string;
  readonly args: unknown;
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
  private readonly dirtyMerges = new Set<SheetId>();
  private dirtyWorkbook = false;
  private calculationScheduler: CalculationScheduler | null = null;
  private nextSeq = 1;
  private documentVersion = 1;
  private editVersionCounter = 0;
  /** Bumped by workbook-level structural changes: sheets, names and date system. */
  private structuralVersionCounter = 0;

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

  /**
   * Bumped by every workbook-level structural change — sheet insert, rename or
   * delete, a defined-name change or the date system — through a command,
   * undo or redo. A save compares this before and after its write await so a
   * structural edit in flight stays dirty.
   */
  get structuralVersion(): number {
    return this.structuralVersionCounter;
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
    return this.dirtyMerges.size > 0;
  }

  dirtyCellMap(): ReadonlyMap<SheetId, ReadonlySet<string>> {
    return this.dirtyCells;
  }

  /** Sheets whose merged ranges changed since the last save. */
  dirtyMergeSheets(): ReadonlySet<SheetId> {
    return this.dirtyMerges;
  }

  /** Capture the current merge lists before a save, for concurrent edits. */
  dirtyMergeSnapshot(): Map<SheetId, readonly GridRange[]> {
    const snapshot = new Map<SheetId, readonly GridRange[]>();
    for (const sheetId of this.dirtyMerges) {
      const sheet = findSheet(this.workbook, sheetId);
      if (sheet !== undefined) {
        snapshot.set(
          sheetId,
          sheet.merges.map((merge) => ({ ...merge.range })),
        );
      }
    }
    return snapshot;
  }

  /**
   * Capture the dimension state before a save. Rows and columns are kept
   * separately because the two record shapes have different fields.
   */
  dirtyDimensionSnapshot(): Map<SheetId, DirtyDimensionSnapshot> {
    const snapshot = new Map<SheetId, DirtyDimensionSnapshot>();
    for (const [sheetId, keys] of this.dirtyDimensions) {
      if (keys.rows.size === 0 && keys.columns.size === 0) continue;
      const sheet = findSheet(this.workbook, sheetId);
      if (sheet === undefined) continue;
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
      const sheet = findSheet(this.workbook, sheetId);
      if (sheet === undefined) continue;
      const cells = new Map<string, Cell | null>();
      for (const key of keys) cells.set(key, cloneCell(sheet.cells.get(key) ?? null));
      snapshot.set(sheetId, cells);
    }
    return snapshot;
  }

  dispatch<Args>(commandId: string, args: Args, options: DispatchOptions = {}): CommandResult {
    if (commandId === 'edit.undo') return this.undo();
    if (commandId === 'edit.redo') return this.redo();
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
    if (command.history === 'terminal') {
      // Excel makes a sheet delete permanent: the change is applied but no
      // undo entry exists, and the rest of the history can no longer be
      // replayed against a workbook that lost a sheet.
      this.undoStack.length = 0;
      this.redoStack.length = 0;
      this.markDirty(result.changes);
    } else {
      this.record(commandId, result.changes);
    }
    this.documentVersion += 1;
    this.editVersionCounter += 1;
    this.calculationScheduler?.afterChanges(result.changes);
    return result;
  }

  /**
   * Run several commands as one undo transaction. Each command runs in order
   * and its changes are applied before the next command runs, so dependent
   * patches compose exactly as they do through separate dispatches. If a
   * command rejects with a toast or throws, the changes already applied in this
   * batch are rolled back and nothing is recorded. The whole batch therefore
   * reaches the document atomically, with one transaction and one history
   * entry, whatever the number of commands.
   */
  dispatchBatch(
    entries: readonly DispatchBatchEntry[],
    options: DispatchOptions = {},
    commandId = 'batch',
  ): CommandResult {
    if (entries.length === 0) return { changes: [] };
    const context: CommandContext = {
      workbook: this.workbook,
      activeSheetId: options.activeSheetId ?? this.workbook.sheets[0]?.id ?? null,
      selection: options.selection ?? null,
      activeCell: options.activeCell ?? null,
    };
    const applied: CommandChange[] = [];
    let terminal = false;
    const rollback = (): void => {
      if (applied.length === 0) return;
      this.applyChanges([...applied].reverse().map(inverseChange), 'after');
      applied.length = 0;
    };
    try {
      for (const entry of entries) {
        const command = this.registry.require(entry.id);
        if (!command.isEnabled(context)) {
          throw new Error(`Command is disabled: ${entry.id}`);
        }
        const result = (command.run as (context: CommandContext, args: unknown) => CommandResult)(
          context,
          entry.args,
        );
        if (result.ui?.message !== undefined) {
          // A rejection means the whole action failed; roll the earlier
          // commands back so no tab is left half-applied.
          rollback();
          return { ...result, changes: [] };
        }
        if (result.changes.length === 0) continue;
        this.applyChanges(result.changes, 'after');
        applied.push(...result.changes);
        if (command.history === 'terminal') terminal = true;
      }
    } catch (error) {
      rollback();
      throw error;
    }
    if (applied.length === 0) return { changes: [] };
    if (terminal) {
      this.undoStack.length = 0;
      this.redoStack.length = 0;
      this.markDirty(applied);
    } else {
      this.record(commandId, applied);
    }
    this.documentVersion += 1;
    this.editVersionCounter += 1;
    this.calculationScheduler?.afterChanges(applied);
    return { changes: applied };
  }

  undo(): CommandResult {
    const transaction = this.undoStack.pop();
    if (transaction === undefined) return { changes: [] };
    this.applyChanges([...transaction.inverseChanges].reverse(), 'after');
    this.redoStack.push(transaction);
    this.trimHistory();
    this.markDirty(transaction.changes);
    this.documentVersion += 1;
    this.editVersionCounter += 1;
    this.calculationScheduler?.afterChanges(transaction.inverseChanges);
    const activeSheetId = activeSheetHintAfterUndo(this.workbook, transaction.changes);
    return activeSheetId === null ? { changes: [] } : { changes: [], activeSheetId };
  }

  redo(): CommandResult {
    const transaction = this.redoStack.pop();
    if (transaction === undefined) return { changes: [] };
    this.applyChanges(transaction.changes, 'after');
    this.undoStack.push(transaction);
    this.trimHistory();
    this.markDirty(transaction.changes);
    this.documentVersion += 1;
    this.editVersionCounter += 1;
    this.calculationScheduler?.afterChanges(transaction.changes);
    const activeSheetId = activeSheetHintAfterRedo(transaction.changes);
    return activeSheetId === null ? { changes: [] } : { changes: [], activeSheetId };
  }

  /**
   * Clear the cells and dimensions that were saved. When a snapshot is
   * supplied, an item that changed while the save was awaiting the file write
   * is compared by state and stays dirty. The structural version captured
   * before the write keeps the workbook dirty when a sheet command ran during
   * the await; that command is not in the saved bytes.
   */
  markSaved(
    saved?: ReadonlyMap<SheetId, ReadonlyMap<string, Cell | null>>,
    savedDimensions?: ReadonlyMap<SheetId, DirtyDimensionSnapshot>,
    savedStructuralVersion?: number,
    savedMerges?: ReadonlyMap<SheetId, readonly GridRange[]>,
  ): void {
    if (
      savedStructuralVersion === undefined ||
      savedStructuralVersion === this.structuralVersionCounter
    ) {
      this.dirtyWorkbook = false;
    }
    if (saved === undefined) {
      this.dirtyCells.clear();
    } else {
      for (const [sheetId, savedCells] of saved) {
        const dirty = this.dirtyCells.get(sheetId);
        if (dirty === undefined) continue;
        const sheet = findSheet(this.workbook, sheetId);
        if (sheet === undefined) continue;
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
        const sheet = findSheet(this.workbook, sheetId);
        if (sheet === undefined) continue;
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
    if (savedMerges === undefined) {
      this.dirtyMerges.clear();
    } else {
      for (const [sheetId, saved] of savedMerges) {
        const sheet = findSheet(this.workbook, sheetId);
        if (sheet === undefined) {
          this.dirtyMerges.delete(sheetId);
          continue;
        }
        const current = sheet.merges.map((merge) => merge.range);
        if (mergeRangeListsEqual(current, saved)) this.dirtyMerges.delete(sheetId);
      }
    }
    // A sheet delete is terminal, so a removed sheet's dirty markers can never
    // be written again. Drop them rather than keeping a save-clean state hostage.
    this.dropDetachedDirtyMarkers();
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
      if (change.type === 'sheet.add') {
        if (direction === 'after') this.insertSheetSnapshot(change.after, change.index);
        else this.removeSheetSnapshot(change.after.id);
        continue;
      }
      if (change.type === 'sheet.delete') {
        if (direction === 'after') this.removeSheetSnapshot(change.before.id);
        else this.insertSheetSnapshot(change.before, change.index);
        continue;
      }
      if (change.type === 'sheet.rename') {
        const sheet = findSheet(this.workbook, change.sheetId);
        if (sheet !== undefined) sheet.name = direction === 'after' ? change.after : change.before;
        continue;
      }
      if (change.type === 'name.set') {
        this.applyNameFormula(change, direction);
        continue;
      }
      const sheet = getSheetById(this.workbook, change.sheetId);
      if (change.type === 'merge.set') {
        const merged = direction === 'after' ? change.after : change.before;
        sheet.merges.splice(
          0,
          sheet.merges.length,
          ...merged.map((range) => ({ range: { ...range } })),
        );
        continue;
      }
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

  private insertSheetSnapshot(sheet: Sheet, index: number): void {
    if (findSheet(this.workbook, sheet.id) !== undefined) return;
    const position = Math.min(Math.max(0, index), this.workbook.sheets.length);
    this.workbook.sheets.splice(position, 0, sheet);
  }

  private removeSheetSnapshot(sheetId: SheetId): void {
    const index = this.workbook.sheets.findIndex((sheet) => sheet.id === sheetId);
    if (index >= 0) this.workbook.sheets.splice(index, 1);
  }

  private applyNameFormula(
    change: {
      readonly scope: 'workbook' | 'sheet';
      readonly sheetId: SheetId | null;
      readonly name: string;
      readonly before: string;
      readonly after: string;
    },
    direction: 'before' | 'after',
  ): void {
    const names =
      change.scope === 'workbook'
        ? this.workbook.names
        : change.sheetId === null
          ? undefined
          : findSheet(this.workbook, change.sheetId)?.names;
    if (names === undefined) return;
    const index = names.findIndex((candidate) => candidate.name === change.name);
    if (index < 0) return;
    names[index] = {
      ...names[index]!,
      formula: direction === 'after' ? change.after : change.before,
    };
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
        sheet.cells.set(key, next);
      }
    }
  }

  private markDirty(changes: readonly CommandChange[]): void {
    for (const change of changes) {
      if (
        change.type === 'workbook.dateSystem' ||
        change.type === 'sheet.add' ||
        change.type === 'sheet.delete' ||
        change.type === 'sheet.rename' ||
        change.type === 'name.set'
      ) {
        this.dirtyWorkbook = true;
        this.structuralVersionCounter += 1;
        if (change.type === 'sheet.delete') this.dropSheetDirtyMarkers(change.before.id);
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
      if (change.type === 'merge.set') {
        this.dirtyMerges.add(change.sheetId);
        continue;
      }
      const cells = this.dirtyCells.get(change.sheetId) ?? new Set<string>();
      for (const cell of changeCells(change)) {
        cells.add(`${cell.row}:${cell.col}`);
      }
      this.dirtyCells.set(change.sheetId, cells);
    }
  }

  /**
   * A deleted sheet is gone for good; its retained cell and dimension markers
   * can never reach a save and must not keep `hasUnsavedChanges` true.
   */
  private dropSheetDirtyMarkers(sheetId: SheetId): void {
    this.dirtyCells.delete(sheetId);
    this.dirtyDimensions.delete(sheetId);
    this.dirtyMerges.delete(sheetId);
  }

  /**
   * Safety net for a snapshot that still names a sheet the terminal delete
   * removed; no sheet outside the workbook can be written again.
   */
  private dropDetachedDirtyMarkers(): void {
    for (const sheetId of [...this.dirtyCells.keys()]) {
      if (findSheet(this.workbook, sheetId) === undefined) this.dirtyCells.delete(sheetId);
    }
    for (const sheetId of [...this.dirtyDimensions.keys()]) {
      if (findSheet(this.workbook, sheetId) === undefined) this.dirtyDimensions.delete(sheetId);
    }
    for (const sheetId of [...this.dirtyMerges]) {
      if (findSheet(this.workbook, sheetId) === undefined) this.dirtyMerges.delete(sheetId);
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

function mergeRangeListsEqual(left: readonly GridRange[], right: readonly GridRange[]): boolean {
  if (left.length !== right.length) return false;
  return left.every(
    (range, index) =>
      range.startRow === right[index]?.startRow &&
      range.startCol === right[index]?.startCol &&
      range.endRow === right[index]?.endRow &&
      range.endCol === right[index]?.endCol,
  );
}

function estimateBytes(changes: readonly CommandChange[]): number {
  let total = 0;
  for (const change of changes) {
    if (change.type === 'workbook.dateSystem') {
      total += 8;
      continue;
    }
    if (change.type === 'dimension.set' || change.type === 'sheet.rename') {
      total += 64;
      continue;
    }
    if (change.type === 'merge.set') {
      total += 32 + (change.before.length + change.after.length) * 32;
      continue;
    }
    if (change.type === 'name.set') {
      total += 64 + change.before.length * 2 + change.after.length * 2;
      continue;
    }
    if (change.type === 'sheet.add' || change.type === 'sheet.delete') {
      total += estimateSheet(change.before ?? change.after);
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

function estimateSheet(sheet: Sheet): number {
  let bytes = 256 + sheet.name.length * 2 + sheet.partPath.length * 2;
  for (const cell of sheet.cells.values()) bytes += estimateCell(cell) + 48;
  bytes += sheet.merges.length * 32;
  bytes += sheet.rowDimensions.size * 48 + sheet.columnDimensions.size * 48;
  bytes += sheet.names.reduce((total, name) => total + name.formula.length * 2, 0);
  return bytes;
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

function activeSheetHintAfterUndo(
  workbook: Workbook,
  changes: readonly CommandChange[],
): SheetId | null {
  for (const change of changes) {
    if (change.type === 'sheet.add') return neighbourSheetId(workbook, change.index);
  }
  for (const change of changes) {
    if (change.type === 'sheet.rename') return change.sheetId;
  }
  return null;
}

function activeSheetHintAfterRedo(changes: readonly CommandChange[]): SheetId | null {
  for (const change of changes) {
    if (change.type === 'sheet.add') return change.after.id;
  }
  for (const change of changes) {
    if (change.type === 'sheet.rename') return change.sheetId;
  }
  return null;
}

function neighbourSheetId(workbook: Workbook, index: number): SheetId | null {
  if (workbook.sheets.length === 0) return null;
  const position = Math.min(Math.max(0, index), workbook.sheets.length - 1);
  // A hidden sheet is never activated; skip to the nearest visible tab.
  for (let offset = 0; offset < workbook.sheets.length; offset += 1) {
    const after = workbook.sheets[position + offset];
    if (after !== undefined && isSheetVisible(workbook, after.id)) return after.id;
    const before = workbook.sheets[position - offset - 1];
    if (before !== undefined && isSheetVisible(workbook, before.id)) return before.id;
  }
  return null;
}
