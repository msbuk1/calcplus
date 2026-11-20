import type { CalculationScheduler, CommandChange } from '@commands/types';
import { type CalculationCellUpdate, DocumentSession } from '@commands/session';
import { absolutizeFormulaReferences } from '@model/references';
import type { Cell, Sheet, Workbook } from '@model/types';
import { cellKey, parseCellKey } from '@model/types';
import { BLANK, type CellValue, errorValue } from '@model/value';
import { FormulaEngineClient, type EngineWorkerFactory } from './formula-engine-client';
import { prepareFormula } from './formula-text';
import type {
  EngineBatchResult,
  EngineCellInput,
  EngineCellRef,
  EngineEdit,
  EngineNamedExpression,
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
  /** True while a result batch is being applied in slices. */
  readonly applying: boolean;
  readonly lastElapsedMs: number;
  readonly lastCellCount: number;
  readonly circularCount: number;
  readonly staleCount: number;
  readonly unsupportedCount: number;
}

export interface CalculationControllerOptions {
  readonly client?: FormulaEngineClient;
  /** Worker factory for the built-in client; a Node test supplies an inline one. */
  readonly workerFactory?: EngineWorkerFactory;
  /**
   * The longest synchronous slice of a result batch, in milliseconds.
   * Performance §3.2 forbids a recalculation from blocking the main thread for
   * more than 50 ms; 8 ms leaves room for the browser and the repaint.
   */
  readonly resultSliceBudgetMs?: number;
  /** A hard cap on updates per slice, so the slice always ends. */
  readonly resultChunkSize?: number;
  /** The longest synchronous slice while the workbook snapshot is built. */
  readonly snapshotSliceBudgetMs?: number;
  /** Touched cells per snapshot slice. */
  readonly snapshotChunkSize?: number;
  /** Yields between slices. Defaults to a macrotask so paint can run. */
  readonly yieldToUi?: () => Promise<void>;
  readonly now?: () => number;
}

const DEFAULT_RESULT_SLICE_BUDGET_MS = 8;
const DEFAULT_RESULT_CHUNK_SIZE = 2_000;
const DEFAULT_SNAPSHOT_SLICE_BUDGET_MS = 8;
const DEFAULT_SNAPSHOT_CHUNK_SIZE = 5_000;

/**
 * Glue between the workbook model and the worker engine.
 *
 * The model owns authored text and cached results. This controller sends only
 * engine inputs and applies only typed values back, so HyperFormula's
 * normalised formula text never reaches the model or the save path.
 *
 * Applying a result is sliced: a 50,000-cell fan-out arrives as one worker
 * batch, but the model is updated in bounded slices with a yield between them.
 * The workbook snapshot is read the same way before a rebuild, so a structural
 * edit cannot freeze the grid while 250,000 cells are copied. A newer edit made
 * while the slices run wins, and a rebuild discards the rest of the old batch.
 * The controller also keeps the stale and unsupported counts incrementally;
 * rescanning every cell on each keystroke would itself break the 50 ms
 * main-thread rule.
 */
export class CalculationController implements CalculationScheduler {
  private readonly client: FormulaEngineClient;
  private readonly editsByVersion = new Map<number, Set<string>>();
  private readonly resultSliceBudgetMs: number;
  private readonly resultChunkSize: number;
  private readonly snapshotSliceBudgetMs: number;
  private readonly snapshotChunkSize: number;
  private readonly yieldToUi: () => Promise<void>;
  private readonly clock: () => number;
  private readonly unsupportedKeys = new Set<string>();
  private disposed = false;
  private started = false;
  private applying = false;
  private buildGeneration = 0;
  private lastElapsedMs = 0;
  private lastCellCount = 0;
  private circularSignature = '';
  private circularCount = 0;
  private staleCount = 0;
  /** Edit version of the last structural rebuild; older results are ignored. */
  private lastRebuildEditVersion = 0;

  constructor(
    private readonly session: DocumentSession,
    private readonly callbacks: CalculationCallbacks,
    options: CalculationControllerOptions = {},
  ) {
    this.resultSliceBudgetMs = Math.max(
      0,
      options.resultSliceBudgetMs ?? DEFAULT_RESULT_SLICE_BUDGET_MS,
    );
    this.resultChunkSize = Math.max(1, options.resultChunkSize ?? DEFAULT_RESULT_CHUNK_SIZE);
    this.snapshotSliceBudgetMs = Math.max(
      0,
      options.snapshotSliceBudgetMs ?? DEFAULT_SNAPSHOT_SLICE_BUDGET_MS,
    );
    this.snapshotChunkSize = Math.max(1, options.snapshotChunkSize ?? DEFAULT_SNAPSHOT_CHUNK_SIZE);
    this.yieldToUi = options.yieldToUi ?? defaultYieldToUi;
    this.clock = options.now ?? (() => performance.now());
    this.client =
      options.client ??
      new FormulaEngineClient(
        {
          onResult: (result, baseEditVersion, phase) =>
            this.handleResult(result, baseEditVersion, phase),
          onError: (message) => {
            if (!this.disposed) this.callbacks.onError(message);
          },
        },
        options.workerFactory === undefined ? {} : { workerFactory: options.workerFactory },
      );
    // Scheduling starts before the first edit can land, but the snapshot build
    // is deferred so the grid can paint from cached results first.
    this.session.setCalculationScheduler(this);
  }

  /** Build the engine from the current workbook. Open does not wait for it. */
  start(): void {
    if (this.disposed || this.started) return;
    this.started = true;
    setTimeout(() => {
      if (this.disposed || this.client.isReady || this.buildGeneration > 0) return;
      this.buildFromWorkbook();
    }, 0);
  }

  /**
   * Rebuild the derived dependency graph after a workbook-level change such
   * as the 1900/1904 switch or a structural edit. The model stays untouched.
   */
  rebuild(): void {
    if (this.disposed) return;
    this.started = true;
    this.lastRebuildEditVersion = this.session.editVersion;
    this.editsByVersion.clear();
    this.buildFromWorkbook();
  }

  afterChanges(changes: readonly CommandChange[]): void {
    if (this.disposed || changes.length === 0) return;
    if (changesRequireRebuild(changes)) {
      // Sheet structure and defined-name text change how the engine resolves
      // every formula, so the derived graph is rebuilt rather than patched
      // cell by cell (workbook-model §5). Results from a batch started before
      // the rebuild address the old sheet indices and are discarded.
      this.rebuild();
      return;
    }
    this.refreshCountsFromChanges(changes);
    const cellChanges = changes.filter((change) => change.type === 'cell.set');
    const edits = editsForChanges(this.session.workbook, cellChanges);
    if (edits.length === 0) return;
    const version = this.session.editVersion;
    const keys = this.editsByVersion.get(version) ?? new Set<string>();
    for (const edit of edits) keys.add(engineKey(edit));
    this.editsByVersion.set(version, keys);
    this.client.applyEdits(edits, version);
  }

  private buildFromWorkbook(): void {
    const generation = this.buildGeneration + 1;
    this.buildGeneration = generation;
    void this.buildSnapshot(generation).catch((error: unknown) => {
      if (!this.disposed && generation === this.buildGeneration) {
        this.callbacks.onError(error instanceof Error ? error.message : String(error));
      }
    });
  }

  /**
   * Build the snapshot in bounded slices, then hand it to the worker. A newer
   * rebuild (or dispose) bumps the generation and discards this one. The edit
   * version captured at the start is the base version: an edit that lands while
   * the snapshot is being read stays queued in the client and supersedes the
   * load result, so a stale snapshot value can never overwrite it.
   */
  private async buildSnapshot(generation: number): Promise<void> {
    const baseEditVersion = this.session.editVersion;
    const snapshot = await snapshotWorkbook(this.session.workbook, {
      isCurrent: () => !this.disposed && generation === this.buildGeneration,
      yieldToUi: this.yieldToUi,
      clock: this.clock,
      sliceBudgetMs: this.snapshotSliceBudgetMs,
      chunkSize: this.snapshotChunkSize,
    });
    if (snapshot === null || this.disposed || generation !== this.buildGeneration) return;
    this.unsupportedKeys.clear();
    for (const key of snapshot.unsupportedKeys) this.unsupportedKeys.add(key);
    this.staleCount = snapshot.staleCount;
    // Edits made before the snapshot started are already in it; re-sending
    // them after a structural edit would address their old coordinates.
    this.client.discardPendingBefore(baseEditVersion);
    this.client.initialize(snapshot.workbook, snapshot.cells, baseEditVersion);
  }

  dispose(): void {
    if (this.disposed) return;
    this.disposed = true;
    this.buildGeneration += 1;
    this.editsByVersion.clear();
    this.session.clearCalculationScheduler(this);
    this.client.dispose();
  }

  state(): CalculationState {
    return {
      ready: this.client.isReady,
      failed: this.client.hasFailed,
      queuedEdits: this.client.queuedEditCount,
      applying: this.applying,
      lastElapsedMs: this.lastElapsedMs,
      lastCellCount: this.lastCellCount,
      circularCount: this.circularCount,
      staleCount: this.staleCount,
      unsupportedCount: this.unsupportedKeys.size,
    };
  }

  private handleResult(
    result: EngineBatchResult,
    baseEditVersion: number,
    phase: 'load' | 'apply',
  ): Promise<void> | void {
    if (this.disposed) return;
    if (baseEditVersion < this.lastRebuildEditVersion) return;
    this.lastElapsedMs = result.elapsedMs;
    this.lastCellCount = result.cells.length;
    this.applying = true;
    return this.applyBatch(result, baseEditVersion, phase).finally(() => {
      this.applying = false;
    });
  }

  /**
   * Apply one engine batch in bounded slices. Results arrive in edit order; the
   * client waits for this promise, so a second result cannot overtake it. An
   * edit that lands between slices is seen by `supersededKeys` and wins over
   * the older derived value.
   */
  private async applyBatch(
    result: EngineBatchResult,
    baseEditVersion: number,
    phase: 'load' | 'apply',
  ): Promise<void> {
    if (this.disposed) return;

    const sheetsById = new Map(this.session.workbook.sheets.map((sheet) => [sheet.id, sheet]));
    const updates: PendingUpdate[] = [];
    for (const cell of result.cells) {
      const sheet = this.session.workbook.sheets[cell.sheet];
      if (sheet === undefined) continue;
      updates.push({
        key: engineKey(cell),
        update: {
          sheetId: sheet.id,
          row: cell.row,
          col: cell.col,
          value: cell.value,
          status: cell.status,
          stale: cell.stale,
        },
      });
    }

    for (const ref of result.skipped) {
      const sheet = this.session.workbook.sheets[ref.sheet];
      if (sheet === undefined) continue;
      const cell = sheet.cells.get(cellKey(ref.row, ref.col));
      if (cell === undefined) continue;
      if (cell.formula === null && cell.formulaSource === undefined) continue;
      updates.push({
        key: engineKey(ref),
        update: {
          sheetId: sheet.id,
          row: ref.row,
          col: ref.col,
          value: skippedDisplayValue(cell),
          status: 'value',
          stale: true,
        },
      });
    }

    let index = 0;
    while (index < updates.length) {
      if (this.disposed) return;
      if (baseEditVersion < this.lastRebuildEditVersion) return;
      const superseded = this.supersededKeys(baseEditVersion);
      const slice: PendingUpdate[] = [];
      const started = this.clock();
      let scanned = 0;
      while (index < updates.length) {
        const pending = updates[index]!;
        index += 1;
        scanned += 1;
        if (!superseded.has(pending.key)) slice.push(pending);
        if (scanned >= this.resultChunkSize) break;
        if (this.clock() - started >= this.resultSliceBudgetMs) break;
      }
      if (slice.length > 0) this.applySlice(slice, sheetsById, phase);
      if (index < updates.length) await this.yieldToUi();
      if (this.disposed) return;
    }

    for (const version of [...this.editsByVersion.keys()]) {
      if (version <= baseEditVersion) this.editsByVersion.delete(version);
    }

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

  private applySlice(
    slice: readonly PendingUpdate[],
    sheetsById: ReadonlyMap<string, Sheet>,
    phase: 'load' | 'apply',
  ): void {
    const wasStale = new Map<string, boolean>();
    for (const pending of slice) {
      const cell = sheetsById
        .get(pending.update.sheetId)
        ?.cells.get(cellKey(pending.update.row, pending.update.col));
      wasStale.set(pending.key, cell?.stale === true);
    }

    const changed = this.session.applyCalculationResults(
      slice.map((pending) => pending.update),
      { markDirty: phase === 'apply' },
    );

    let delta = 0;
    for (const pending of slice) {
      const cell = sheetsById
        .get(pending.update.sheetId)
        ?.cells.get(cellKey(pending.update.row, pending.update.col));
      const nowStale = cell?.stale === true;
      const before = wasStale.get(pending.key) === true;
      if (nowStale && !before) delta += 1;
      else if (!nowStale && before) delta -= 1;
    }
    this.staleCount = Math.max(0, this.staleCount + delta);

    if (changed > 0) this.callbacks.onModelChanged();
  }

  /** Cells edited by a batch newer than `baseEditVersion`, keyed like the engine. */
  private supersededKeys(baseEditVersion: number): Set<string> {
    const superseded = new Set<string>();
    for (const [version, keys] of this.editsByVersion) {
      if (version <= baseEditVersion) continue;
      for (const key of keys) superseded.add(key);
    }
    return superseded;
  }

  /**
   * Keep the stale and unsupported counts up to date from the edited ranges
   * only. The `cell.set` change carries its before/after snapshots, so the
   * stale delta is exact and no cell has to be rescanned. The authored formula
   * text is what decides "unsupported", so calculation results never change it.
   */
  private refreshCountsFromChanges(changes: readonly CommandChange[]): void {
    let staleDelta = 0;
    for (const change of changes) {
      if (change.type !== 'cell.set') continue;
      if (!this.session.workbook.sheets.some((candidate) => candidate.id === change.sheetId)) {
        continue;
      }
      for (let offsetRow = 0; offsetRow < change.after.length; offsetRow += 1) {
        const afterRow = change.after[offsetRow] ?? [];
        const beforeRow = change.before[offsetRow] ?? [];
        for (let offsetCol = 0; offsetCol < afterRow.length; offsetCol += 1) {
          const before = beforeRow[offsetCol] ?? null;
          const after = afterRow[offsetCol] ?? null;
          const wasStale = before?.stale === true;
          const isStale = after?.stale === true;
          if (isStale && !wasStale) staleDelta += 1;
          else if (!isStale && wasStale) staleDelta -= 1;

          const row = change.range.startRow + offsetRow;
          const col = change.range.startCol + offsetCol;
          const key = `${change.sheetId}:${row}:${col}`;
          if (after !== null && after.formula !== null && isNotRecalculated(after)) {
            this.unsupportedKeys.add(key);
          } else {
            this.unsupportedKeys.delete(key);
          }
        }
      }
    }
    this.staleCount = Math.max(0, this.staleCount + staleDelta);
  }
}

interface PendingUpdate {
  /** Engine key, matching `editsByVersion`. */
  readonly key: string;
  readonly update: CalculationCellUpdate;
}

function defaultYieldToUi(): Promise<void> {
  return new Promise((resolve) => {
    setTimeout(resolve, 0);
  });
}

/**
 * A sheet insert/delete/rename or a defined-name text change alters the
 * engine's sheet list or name table, so the dependency graph is rebuilt.
 */
export function changesRequireRebuild(changes: readonly CommandChange[]): boolean {
  return changes.some(
    (change) =>
      change.type === 'workbook.dateSystem' ||
      change.type === 'name.set' ||
      change.type === 'sheet.add' ||
      change.type === 'sheet.delete' ||
      change.type === 'sheet.rename' ||
      change.type === 'structure.apply',
  );
}

interface WorkbookSnapshot {
  readonly workbook: EngineWorkbookInput;
  readonly cells: readonly EngineCellInput[];
  readonly unsupportedKeys: readonly string[];
  readonly staleCount: number;
}

interface SnapshotSliceOptions {
  readonly isCurrent: () => boolean;
  readonly yieldToUi: () => Promise<void>;
  readonly clock: () => number;
  readonly sliceBudgetMs: number;
  readonly chunkSize: number;
}

/**
 * Read the model into engine inputs in bounded slices. It returns `null` when a
 * newer rebuild or a dispose made this snapshot obsolete. Blank cells still
 * count toward the stale total, exactly like `DocumentSession.staleCellCount`,
 * but they are not sent to the engine.
 */
async function snapshotWorkbook(
  workbook: Workbook,
  options: SnapshotSliceOptions,
): Promise<WorkbookSnapshot | null> {
  const cells: EngineCellInput[] = [];
  const unsupportedKeys: string[] = [];
  const sheetNames = workbook.sheets.map((sheet) => sheet.name);
  let staleCount = 0;

  let sliceStarted = options.clock();
  let inSlice = 0;
  const takeSlice = async (): Promise<boolean> => {
    inSlice += 1;
    if (inSlice < options.chunkSize && options.clock() - sliceStarted < options.sliceBudgetMs) {
      return true;
    }
    inSlice = 0;
    await options.yieldToUi();
    sliceStarted = options.clock();
    return options.isCurrent();
  };

  for (const [sheetIndex, sheet] of workbook.sheets.entries()) {
    for (const [key, cell] of sheet.cells) {
      if (cell.stale === true) staleCount += 1;
      if (
        cell.value.type === 'blank' &&
        cell.formula === null &&
        cell.formulaSource === undefined
      ) {
        if (!(await takeSlice())) return null;
        continue;
      }
      const { row, col } = parseCellKey(key);
      const input = engineInputFor(cell);
      if (cell.formula !== null && input.formula === null) {
        unsupportedKeys.push(`${sheet.id}:${row}:${col}`);
      }
      cells.push({
        sheet: sheetIndex,
        row,
        col,
        formula: input.formula,
        value: sourceValueFor(cell),
        staleInput: input.staleInput,
      });
      if (!(await takeSlice())) return null;
    }
  }

  return {
    workbook: {
      date1904: workbook.date1904,
      sheetNames,
      namedExpressions: namedExpressionsFor(workbook),
    },
    cells,
    unsupportedKeys,
    staleCount,
  };
}

/**
 * The engine-only form of defined names. The model keeps the authored text;
 * this copy is screened against the 52-function registry and absolutised
 * because HyperFormula rejects relative references in named expressions.
 */
function namedExpressionsFor(workbook: Workbook): readonly EngineNamedExpression[] {
  const expressions: EngineNamedExpression[] = [];
  const add = (name: string, formula: string, sheet: number | null): void => {
    const text = formula.trim().replace(/^=/, '');
    if (text === '' || name.trim() === '') return;
    const prepared = prepareFormula(text);
    if (prepared.text === null) return;
    const expression = absolutizeFormulaReferences(prepared.text);
    if (expression.trim() === '') return;
    expressions.push({ name, expression: `=${expression}`, sheet });
  };
  for (const name of workbook.names) add(name.name, name.formula, null);
  for (const [sheetIndex, sheet] of workbook.sheets.entries()) {
    for (const name of sheet.names) add(name.name, name.formula, sheetIndex);
  }
  return expressions;
}

function editsForChanges(
  workbook: Workbook,
  changes: readonly CommandChange[],
): readonly EngineEdit[] {
  const edits = new Map<string, EngineEdit>();
  for (const change of changes) {
    if (change.type !== 'cell.set') continue;
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
