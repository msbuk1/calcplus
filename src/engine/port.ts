import type { CellValue } from '@model/value';

/**
 * The narrow formula-engine port (technical design §3.5).
 *
 * HyperFormula sits behind this interface inside a Web Worker. The port
 * returns values and statuses only: authored formula text and saved results
 * never travel back from the engine into the model or the file.
 */

/** Zero-based engine coordinates. Sheet indices follow the workbook tab order. */
export interface EngineCellRef {
  readonly sheet: number;
  readonly row: number;
  readonly col: number;
}

/**
 * One cell handed to the engine. `formula` is the authored text without the
 * leading equals sign, or `null` when the cell is a literal. The engine gates
 * the text against the published function registry and keeps the authored
 * form out of its own result.
 *
 * `staleInput` marks a cell whose authored formula must not be recalculated
 * (shared groups, legacy arrays, rich text). Its cached result is used by
 * dependants, which are marked stale in turn (formula semantics §9).
 */
export interface EngineCellInput extends EngineCellRef {
  readonly formula: string | null;
  readonly value: CellValue;
  readonly staleInput?: boolean;
}

/** One edited cell. A `null` formula replaces the cell with the literal value. */
export interface EngineEdit extends EngineCellRef {
  readonly formula: string | null;
  readonly value: CellValue;
  readonly staleInput?: boolean;
}

export type EngineCellStatus = 'value' | 'circular';

/**
 * A calculated result. Only formula cells produce one. `stale` marks a result
 * computed from an input the engine did not evaluate; such a value is shown
 * but never written to the file (file-compat §6.1).
 */
export interface EngineResultCell extends EngineCellRef {
  readonly value: CellValue;
  readonly status: EngineCellStatus;
  readonly stale: boolean;
}

/**
 * The dependency graph is owned by HyperFormula. `cells` contains only the
 * formula cells whose values changed in this batch, so callers apply an
 * incremental delta instead of rebuilding a sheet.
 */
export interface EngineBatchResult {
  readonly cells: readonly EngineResultCell[];
  /** Every formula cell currently in a cycle. */
  readonly circular: readonly EngineCellRef[];
  /** Formula cells the engine did not evaluate; their cached result stays. */
  readonly skipped: readonly EngineCellRef[];
  readonly elapsedMs: number;
}

export interface EngineWorkbookInput {
  readonly date1904: boolean;
  readonly sheetNames: readonly string[];
}

/**
 * The synchronous adapter surface implemented in the worker. The main thread
 * talks to it through `FormulaEngineClient`, which adds the async worker
 * boundary and request coalescing.
 */
export interface FormulaEngine {
  start(workbook: EngineWorkbookInput): void;
  load(cells: readonly EngineCellInput[]): void;
  finishLoad(): EngineBatchResult;
  apply(edits: readonly EngineEdit[]): EngineBatchResult;
  dispose(): void;
}
