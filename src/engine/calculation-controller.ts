import type { CalculationScheduler, CommandChange } from '@commands/types';
import { type CalculationCellUpdate, DocumentSession } from '@commands/session';
import type { Cell, Workbook } from '@model/types';
import { cellKey, parseCellKey } from '@model/types';
import { BLANK, type CellValue, errorValue } from '@model/value';
import { FormulaEngineClient } from './formula-engine-client';
import { prepareFormula } from './formula-text';
import type {
  EngineBatchResult,
  EngineCellInput,
  EngineCellRef,
  EngineEdit,
  EngineWorkbookInput,
} from './port';

export interface CalculationCallbacks {
  /** A derived value changed; the view should repaint. */
  readonly onModelChanged: () => void;
  /** Called whenever the set of cells in a cycle changes. */
  readonly onCircular: (circular: readonly EngineCellRef[], sheetNames: readonly string[]) => void;
  readonly onError: (message: string) => void;
}

export interface CalculationState {
  readonly ready: boolean;
  readonly failed: boolean;
  readonly queuedEdits: number;
  readonly lastElapsedMs: number;
  readonly lastCellCount: number;
  readonly circularCount: number;
  readonly staleCount: number;
  readonly unsupportedCount: number;
}

export interface CalculationControllerOptions {
  readonly client?: FormulaEngineClient;
}

/**
 * Glue between the workbook model and the worker engine.
 *
 * The model owns authored text and cached results. This controller sends only
 * engine inputs and applies only typed values back, so HyperFormula's
 * normalised formula text never reaches the model or the save path.
 */
export class CalculationController implements CalculationScheduler {
  private readonly client: FormulaEngineClient;
  private readonly editsByVersion = new Map<number, Set<string>>();
  private disposed = false;
  private started = false;
  private lastElapsedMs = 0;
  private lastCellCount = 0;
  private circularSignature = '';
  private circularCount = 0;
  private staleCount = 0;
  private unsupportedCount = 0;

  constructor(
    private readonly session: DocumentSession,
    private readonly callbacks: CalculationCallbacks,
    options: CalculationControllerOptions = {},
  ) {
    this.client =
      options.client ??
      new FormulaEngineClient({
        onResult: (result, baseEditVersion, phase) =>
          this.handleResult(result, baseEditVersion, phase),
        onError: (message) => {
          if (!this.disposed) this.callbacks.onError(message);
        },
      });
    // Scheduling starts before the first edit can land, but the snapshot build
    // is deferred so the grid can paint from cached results first.
    this.session.setCalculationScheduler(this);
  }

  /** Build the engine from the current workbook. Open does not wait for it. */
  start(): void {
    if (this.disposed || this.started) return;
    this.started = true;
    setTimeout(() => {
      if (this.disposed) return;
      try {
        const snapshot = snapshotWorkbook(this.session.workbook);
        this.unsupportedCount = snapshot.unsupportedCount;
        this.staleCount = this.session.staleCellCount();
        this.client.initialize(snapshot.workbook, snapshot.cells, this.session.editVersion);
      } catch (error) {
        this.callbacks.onError(error instanceof Error ? error.message : String(error));
      }
    }, 0);
  }

  afterChanges(changes: readonly CommandChange[]): void {
    if (this.disposed || changes.length === 0) return;
    const edits = editsForChanges(this.session.workbook, changes);
    if (edits.length === 0) return;
    const version = this.session.editVersion;
    const keys = this.editsByVersion.get(version) ?? new Set<string>();
    for (const edit of edits) keys.add(engineKey(edit));
    this.editsByVersion.set(version, keys);
    this.client.applyEdits(edits, version);
  }

  dispose(): void {
    if (this.disposed) return;
    this.disposed = true;
    this.editsByVersion.clear();
    this.session.clearCalculationScheduler(this);
    this.client.dispose();
  }

  state(): CalculationState {
    return {
      ready: this.client.isReady,
      failed: this.client.hasFailed,
      queuedEdits: this.client.queuedEditCount,
      lastElapsedMs: this.lastElapsedMs,
      lastCellCount: this.lastCellCount,
      circularCount: this.circularCount,
      staleCount: this.staleCount,
      unsupportedCount: this.unsupportedCount,
    };
  }

  private handleResult(
    result: EngineBatchResult,
    baseEditVersion: number,
    phase: 'load' | 'apply',
  ): void {
    if (this.disposed) return;
    this.lastElapsedMs = result.elapsedMs;
    this.lastCellCount = result.cells.length;

    // Results arrive in edit order. Apply what this batch settled, but never
    // overwrite a cell edited by a later batch that is still in flight.
    const superseded = new Set<string>();
    for (const [version, keys] of this.editsByVersion) {
      if (version <= baseEditVersion) continue;
      for (const key of keys) superseded.add(key);
    }

    const updates: CalculationCellUpdate[] = [];
    for (const cell of result.cells) {
      if (superseded.has(engineKey(cell))) continue;
      const sheet = this.session.workbook.sheets[cell.sheet];
      if (sheet === undefined) continue;
      updates.push({
        sheetId: sheet.id,
        row: cell.row,
        col: cell.col,
        value: cell.value,
        status: cell.status,
        stale: cell.stale,
      });
    }

    for (const ref of result.skipped) {
      if (superseded.has(engineKey(ref))) continue;
      const sheet = this.session.workbook.sheets[ref.sheet];
      if (sheet === undefined) continue;
      const cell = sheet.cells.get(cellKey(ref.row, ref.col));
      if (cell === undefined) continue;
      if (cell.formula === null && cell.formulaSource === undefined) continue;
      updates.push({
        sheetId: sheet.id,
        row: ref.row,
        col: ref.col,
        value: skippedDisplayValue(cell),
        status: 'value',
        stale: true,
      });
    }

    const changed = this.session.applyCalculationResults(updates, {
      markDirty: phase === 'apply',
    });
    if (changed > 0) this.callbacks.onModelChanged();

    for (const version of [...this.editsByVersion.keys()]) {
      if (version <= baseEditVersion) this.editsByVersion.delete(version);
    }

    this.refreshCounts();
    this.circularCount = result.circular.length;
    const signature = result.circular
      .map((ref) => `${ref.sheet}:${ref.row}:${ref.col}`)
      .sort()
      .join('|');
    if (signature !== this.circularSignature) {
      this.circularSignature = signature;
      this.callbacks.onCircular(
        result.circular,
        this.session.workbook.sheets.map((sheet) => sheet.name),
      );
    }
  }

  private refreshCounts(): void {
    let stale = 0;
    let unsupported = 0;
    for (const sheet of this.session.workbook.sheets) {
      for (const cell of sheet.cells.values()) {
        if (cell.stale === true) stale += 1;
        if (cell.formula !== null && isNotRecalculated(cell)) unsupported += 1;
      }
    }
    this.staleCount = stale;
    this.unsupportedCount = unsupported;
  }
}

interface WorkbookSnapshot {
  readonly workbook: EngineWorkbookInput;
  readonly cells: readonly EngineCellInput[];
  readonly unsupportedCount: number;
}

function snapshotWorkbook(workbook: Workbook): WorkbookSnapshot {
  const cells: EngineCellInput[] = [];
  let unsupportedCount = 0;
  const sheetNames = workbook.sheets.map((sheet) => sheet.name);

  for (const [sheetIndex, sheet] of workbook.sheets.entries()) {
    for (const [key, cell] of sheet.cells) {
      if (
        cell.value.type === 'blank' &&
        cell.formula === null &&
        cell.formulaSource === undefined
      ) {
        continue;
      }
      const { row, col } = parseCellKey(key);
      const input = engineInputFor(cell);
      if (cell.formula !== null && input.formula === null) unsupportedCount += 1;
      cells.push({
        sheet: sheetIndex,
        row,
        col,
        formula: input.formula,
        value: sourceValueFor(cell),
        staleInput: input.staleInput,
      });
    }
  }

  return {
    workbook: { date1904: workbook.date1904, sheetNames },
    cells,
    unsupportedCount,
  };
}

function editsForChanges(
  workbook: Workbook,
  changes: readonly CommandChange[],
): readonly EngineEdit[] {
  const edits = new Map<string, EngineEdit>();
  for (const change of changes) {
    const sheetIndex = workbook.sheets.findIndex((sheet) => sheet.id === change.sheetId);
    if (sheetIndex < 0) continue;
    const sheet = workbook.sheets[sheetIndex]!;
    for (let row = change.range.startRow; row <= change.range.endRow; row += 1) {
      for (let col = change.range.startCol; col <= change.range.endCol; col += 1) {
        const key = `${sheetIndex}:${row}:${col}`;
        if (edits.has(key)) continue;
        const cell = sheet.cells.get(cellKey(row, col)) ?? null;
        if (cell === null) {
          edits.set(key, {
            sheet: sheetIndex,
            row,
            col,
            formula: null,
            value: BLANK,
            staleInput: false,
          });
          continue;
        }
        const input = engineInputFor(cell);
        edits.set(key, {
          sheet: sheetIndex,
          row,
          col,
          formula: input.formula,
          value: sourceValueFor(cell),
          staleInput: input.staleInput,
        });
      }
    }
  }
  return [...edits.values()];
}

/**
 * Only plain authored formulas are handed over as formulas. Shared groups,
 * legacy arrays and rich text keep their authored text and cached value; the
 * engine uses that cached value and marks dependants stale (formula semantics
 * §6 and §9). Unsupported functions and unparseable text are screened here so
 * their authored text never even enters the worker; the adapter screens again
 * for direct use.
 */
function engineInputFor(cell: Cell): { formula: string | null; staleInput: boolean } {
  if (cell.formula === null || cell.richText !== null || cell.formulaSource !== undefined) {
    return {
      formula: null,
      staleInput: cell.formula !== null || cell.formulaSource !== undefined,
    };
  }
  if (prepareFormula(cell.formula).text === null) {
    return { formula: null, staleInput: true };
  }
  return { formula: cell.formula, staleInput: false };
}

function sourceValueFor(cell: Cell): CellValue {
  if (cell.formula === null) return cell.value;
  if (cell.savedResult !== null) return cell.savedResult;
  if (cell.formulaSource !== undefined || cell.richText !== null) return cell.value;
  const prepared = prepareFormula(cell.formula);
  // An unsupported function without a cached result shows #NAME?. Feeding that
  // error through the engine lets dependants propagate it exactly as Excel
  // would; they are marked stale by the worker's dependency walk.
  if (prepared.skip === 'unsupported-function') return errorValue('#NAME?');
  return BLANK;
}

function skippedDisplayValue(cell: Cell): CellValue {
  if (cell.savedResult !== null) return cell.savedResult;
  if (cell.formulaSource !== undefined || cell.richText !== null) return cell.value;
  const prepared = cell.formula === null ? null : prepareFormula(cell.formula);
  return prepared?.skip === 'unsupported-function' ? errorValue('#NAME?') : BLANK;
}

function isNotRecalculated(cell: Cell): boolean {
  if (cell.formulaSource !== undefined || cell.richText !== null) return true;
  const prepared = prepareFormula(cell.formula ?? '');
  return prepared.skip !== null;
}

function engineKey(ref: EngineCellRef): string {
  return `${ref.sheet}:${ref.row}:${ref.col}`;
}
