import {
  DetailedCellError,
  ErrorType,
  ExportedCellChange,
  HyperFormula,
  type CellValue as HyperFormulaCellValue,
  type ConfigParams,
  type ExportedChange,
  type RawCellContent,
  type SimpleCellAddress,
  type SimpleCellRange,
} from 'hyperformula';
import type { CellValue } from '@model/value';
import {
  BLANK,
  booleanValue,
  errorValue,
  numberValue,
  textValue,
  type CellErrorCode,
} from '@model/value';
import { prepareFormula } from './formula-text';
import { engineLanguage, registerCalcplusLanguage, supportsNullError } from './engine-language';
import { registerCalcplusFunctions } from './shim-plugin';
import type {
  EngineBatchResult,
  EngineCellInput,
  EngineCellRef,
  EngineEdit,
  EngineResultCell,
  EngineWorkbookInput,
  FormulaEngine,
} from './port';

const MAX_ROWS = 1_048_576;
const MAX_COLUMNS = 16_384;

/** Distinct formulas screened before the validation cache clears. */
const VALIDATE_CACHE_LIMIT = 4096;

/** One input cell after screening: the engine content and the decision made. */
interface InputPlan {
  readonly key: string;
  readonly content: RawCellContent;
  /** True when the engine accepts the rewritten formula for this cell. */
  readonly accepted: boolean;
}

/** A planned engine write, addressed for block writing. */
interface PlannedWrite {
  readonly address: SimpleCellAddress;
  readonly content: RawCellContent;
}

/** A hole-free block of planned writes: full column run, adjacent rows. */
interface CellBand {
  readonly firstRow: number;
  lastRow: number;
  readonly firstCol: number;
  readonly lastCol: number;
}

/**
 * HyperFormula GPLv3 adapter. It runs inside the formula worker; nothing in
 * this module is imported by the main thread (ADR 0004).
 *
 * The adapter never returns formula text. It reads values and error types from
 * the engine and maps them to the model's typed values. Authored text is
 * screened and rewritten for the engine only (formula semantics §9).
 */
export class HyperFormulaEngine implements FormulaEngine {
  private hf: HyperFormula | null = null;
  private readonly formulaCells = new Map<string, EngineCellRef>();
  private readonly circularCells = new Map<string, EngineCellRef>();
  private readonly skippedCells = new Map<string, EngineCellRef>();
  private readonly validateCache = new Map<string, boolean>();
  private suspended = false;

  start(workbook: EngineWorkbookInput): void {
    this.dispose();
    registerCalcplusLanguage();
    registerCalcplusFunctions();
    const engine = HyperFormula.buildEmpty(engineConfig(workbook.date1904));
    this.hf = engine;
    for (const [index, name] of workbook.sheetNames.entries()) {
      try {
        engine.addSheet(name);
      } catch {
        // A malformed workbook can carry duplicate or illegal sheet names.
        // Keep the sheet index stable so calculations still run for the rest.
        try {
          engine.addSheet(`calcplus_${index + 1}`);
        } catch {
          // Nothing left to do; formulas referencing the sheet stay unresolved.
        }
      }
    }
    // Defined names are read-only in v1, but formulas may use them. The
    // calculation controller has already screened and absolutised the
    // authored name formula; registration is best-effort so one malformed
    // name cannot stop calculation for the workbook.
    for (const named of workbook.namedExpressions ?? []) {
      const scope = named.sheet ?? undefined;
      try {
        if (engine.isItPossibleToAddNamedExpression(named.name, named.expression, scope)) {
          engine.addNamedExpression(named.name, named.expression, scope);
        }
      } catch {
        // Preserve the authored name and let formulas that use it evaluate as
        // #NAME? rather than failing the whole engine build.
      }
    }
    // Bare TRUE and FALSE are named expressions, not function calls
    // (technical design §3.5).
    addBooleanName(engine, 'TRUE');
    addBooleanName(engine, 'FALSE');
    try {
      engine.suspendEvaluation();
      this.suspended = true;
    } catch {
      this.suspended = false;
    }
  }

  /**
   * Bulk load. The worker sends the snapshot in chunks of 10,000 cells and
   * each chunk lands here, so this writes exactly the listed cells and no
   * others. Cells are screened first (one plan per cell), then written in
   * rectangular blocks: a large sheet must not pay the per-call engine cost
   * for every cell. A block write that throws falls back to the original
   * per-cell path, so a malformed cell behaves exactly as it always has.
   */
  load(cells: readonly EngineCellInput[]): void {
    const engine = this.requireEngine();
    const writes = new Map<string, PlannedWrite>();
    for (const cell of cells) {
      const plan = this.planInput(engine, cell);
      writes.set(plan.key, { address: addressOf(cell), content: plan.content });
    }
    try {
      writePlannedCells(engine, writes);
    } catch {
      for (const cell of cells) this.setInput(engine, cell);
    }
  }

  finishLoad(): EngineBatchResult {
    const engine = this.requireEngine();
    const started = now();
    if (this.suspended) {
      try {
        engine.resumeEvaluation();
      } finally {
        this.suspended = false;
      }
    }

    const staleKeys = this.staleDependantKeys(engine);
    const cells: EngineResultCell[] = [];
    this.circularCells.clear();
    for (const [key, ref] of this.formulaCells) {
      const result = this.resultFor(engine, ref, staleKeys.has(key));
      cells.push(result);
      if (result.status === 'circular') this.circularCells.set(key, ref);
    }
    return this.batchResult(cells, started);
  }

  apply(edits: readonly EngineEdit[]): EngineBatchResult {
    const engine = this.requireEngine();
    const started = now();
    const editedFormulaKeys = new Set<string>();
    const changed = new Map<string, EngineCellRef>();

    // HyperFormula marks TODAY/NOW as volatile in its dependency graph, so
    // every setCellContents pass also recomputes them and their dependants.
    for (const edit of edits) {
      const address = addressOf(edit);
      const key = keyOf(edit);
      let exported: ExportedChange[] = [];
      let acceptedFormula = false;

      if (edit.staleInput === true) {
        this.skipFormula(key, refOf(edit));
      } else if (edit.formula !== null) {
        const prepared = prepareFormula(edit.formula);
        if (prepared.text !== null) {
          const candidate = `=${prepared.text}`;
          if (this.validate(engine, candidate)) {
            try {
              exported = engine.setCellContents(address, candidate);
              this.formulaCells.set(key, refOf(edit));
              this.skippedCells.delete(key);
              acceptedFormula = true;
            } catch {
              acceptedFormula = false;
            }
          }
        }
      }

      if (!acceptedFormula) {
        this.formulaCells.delete(key);
        this.circularCells.delete(key);
        if (edit.staleInput === true || edit.formula !== null) {
          this.skipFormula(key, refOf(edit));
        } else {
          this.skippedCells.delete(key);
        }
        exported = this.setLiteral(engine, address, edit.value);
      }

      if (acceptedFormula) editedFormulaKeys.add(key);
      this.collectChanged(exported, changed);
    }

    // The edited formula cell itself must be reported even when the engine
    // considers its value unchanged, so the model refreshes its cached result.
    for (const key of editedFormulaKeys) {
      const ref = this.formulaCells.get(key);
      if (ref !== undefined) changed.set(key, ref);
    }

    const staleKeys = this.staleDependantKeys(engine);
    const cells: EngineResultCell[] = [];
    for (const [key, ref] of changed) {
      const result = this.resultFor(engine, ref, staleKeys.has(key));
      cells.push(result);
      if (result.status === 'circular') this.circularCells.set(key, ref);
      else this.circularCells.delete(key);
    }
    return this.batchResult(cells, started);
  }

  dispose(): void {
    const engine = this.hf;
    this.hf = null;
    this.formulaCells.clear();
    this.circularCells.clear();
    this.skippedCells.clear();
    this.validateCache.clear();
    this.suspended = false;
    if (engine !== null) {
      try {
        engine.destroy();
      } catch {
        // Destroying an already broken engine must not mask the original error.
      }
    }
  }

  /**
   * Screen one input cell, record its bookkeeping state and return the engine
   * content for it: the rewritten formula when the engine accepts it, the
   * cached literal otherwise (formula semantics §9).
   */
  private planInput(engine: HyperFormula, cell: EngineCellInput): InputPlan {
    const key = keyOf(cell);

    if (cell.staleInput !== true && cell.formula !== null) {
      const prepared = prepareFormula(cell.formula);
      if (prepared.text !== null) {
        const candidate = `=${prepared.text}`;
        if (this.validate(engine, candidate)) {
          this.formulaCells.set(key, refOf(cell));
          this.skippedCells.delete(key);
          return { key, content: candidate, accepted: true };
        }
      }
    }

    this.formulaCells.delete(key);
    this.circularCells.delete(key);
    if (cell.staleInput === true || cell.formula !== null) {
      this.skipFormula(key, refOf(cell));
    } else {
      this.skippedCells.delete(key);
    }
    return { key, content: rawContentFor(cell.value), accepted: false };
  }

  private setInput(engine: HyperFormula, cell: EngineCellInput): void {
    const plan = this.planInput(engine, cell);
    const address = addressOf(cell);

    if (!plan.accepted) {
      // The plan already recorded the cached-value state.
      this.setLiteral(engine, address, cell.value);
      return;
    }

    try {
      engine.setCellContents(address, plan.content);
    } catch {
      // The engine refused the formula: fall back to the cached-value
      // literal, as a single rejected cell always has.
      this.formulaCells.delete(plan.key);
      this.circularCells.delete(plan.key);
      this.skipFormula(plan.key, refOf(cell));
      this.setLiteral(engine, address, cell.value);
    }
  }

  private skipFormula(key: string, ref: EngineCellRef): void {
    this.formulaCells.delete(key);
    this.circularCells.delete(key);
    this.skippedCells.set(key, ref);
  }

  private setLiteral(
    engine: HyperFormula,
    address: SimpleCellAddress,
    value: CellValue,
  ): ExportedChange[] {
    try {
      return engine.setCellContents(address, rawContentFor(value));
    } catch {
      return [];
    }
  }

  private collectChanged(
    exported: readonly ExportedChange[],
    changed: Map<string, EngineCellRef>,
  ): void {
    for (const change of exported) {
      if (!(change instanceof ExportedCellChange)) continue;
      const ref: EngineCellRef = { sheet: change.sheet, row: change.row, col: change.col };
      const key = keyOf(ref);
      if (this.formulaCells.has(key)) changed.set(key, ref);
    }
  }

  /**
   * A formula that depends on an unevaluated input is displayed but never
   * written from that input (file-compat §6.1). Walk HyperFormula's own
   * dependency graph from every skipped cell, including range vertices so a
   * whole-column reference does not need materialising.
   */
  private staleDependantKeys(engine: HyperFormula): Set<string> {
    const stale = new Set<string>();
    if (this.skippedCells.size === 0) return stale;
    const queue: (SimpleCellAddress | SimpleCellRange)[] = [...this.skippedCells.values()];
    const visited = new Set<string>();
    while (queue.length > 0) {
      const node = queue.pop();
      if (node === undefined) break;
      const key = nodeKey(node);
      if (visited.has(key)) continue;
      visited.add(key);
      let dependants: (SimpleCellAddress | SimpleCellRange)[];
      try {
        dependants = engine.getCellDependents(node);
      } catch {
        continue;
      }
      for (const dependant of dependants) {
        if (isRangeNode(dependant)) {
          queue.push(dependant);
          continue;
        }
        const dependantKey = keyOf(dependant);
        if (this.formulaCells.has(dependantKey)) {
          stale.add(dependantKey);
          queue.push(dependant);
        }
      }
    }
    return stale;
  }

  private resultFor(engine: HyperFormula, ref: EngineCellRef, stale: boolean): EngineResultCell {
    const raw = engine.getCellValue(addressOf(ref));
    let status: EngineResultCell['status'] = 'value';
    if (raw instanceof DetailedCellError) {
      if (raw.type === ErrorType.CYCLE) status = 'circular';
    }
    return {
      ...ref,
      value: mapEngineValue(raw),
      status,
      stale: stale || status === 'circular',
    };
  }

  private batchResult(cells: readonly EngineResultCell[], started: number): EngineBatchResult {
    return {
      cells,
      circular: [...this.circularCells.values()],
      skipped: [...this.skippedCells.values()],
      elapsedMs: Math.max(0, now() - started),
    };
  }

  /**
   * Parse-check one engine formula. The engine parses the same text to the
   * same verdict within one workbook build, and a filled-down sheet asks
   * about the same text thousands of times, so one check per distinct
   * formula is enough.
   */
  private validate(engine: HyperFormula, formula: string): boolean {
    const cached = this.validateCache.get(formula);
    if (cached !== undefined) return cached;
    let valid: boolean;
    try {
      valid = engine.validateFormula(formula);
    } catch {
      valid = false;
    }
    if (this.validateCache.size >= VALIDATE_CACHE_LIMIT) this.validateCache.clear();
    this.validateCache.set(formula, valid);
    return valid;
  }

  private requireEngine(): HyperFormula {
    if (this.hf === null) throw new Error('The formula engine has not been started');
    return this.hf;
  }
}

function engineConfig(date1904: boolean): Partial<ConfigParams> {
  const language = engineLanguage();
  return {
    licenseKey: 'gpl-v3',
    ...(language === undefined ? {} : { language }),
    leapYear1900: !date1904,
    maxRows: MAX_ROWS,
    maxColumns: MAX_COLUMNS,
    nullDate: date1904 ? { year: 1904, month: 1, day: 1 } : { year: 1899, month: 12, day: 31 },
    useArrayArithmetic: false,
    // Undo and redo live in the model's command history; the engine port
    // exposes neither. Disabling HyperFormula's undo skips its per-write
    // bookkeeping, which dominated bulk loads.
    undoLimit: 0,
    ignoreWhiteSpace: 'any',
    evaluateNullToZero: true,
    caseSensitive: false,
    useWildcards: true,
    matchWholeCell: true,
    functionArgSeparator: ',',
    decimalSeparator: '.',
  };
}

function addBooleanName(engine: HyperFormula, name: string): void {
  try {
    engine.addNamedExpression(name, `=${name}()`);
  } catch {
    // A workbook may one day carry a conflicting defined name (M3C);
    // the engine still has TRUE()/FALSE() as functions.
  }
}

function addressOf(ref: EngineCellRef): SimpleCellAddress {
  return { sheet: ref.sheet, row: ref.row, col: ref.col };
}

function refOf(ref: EngineCellRef): EngineCellRef {
  return { sheet: ref.sheet, row: ref.row, col: ref.col };
}

function keyOf(ref: EngineCellRef): string {
  return `${ref.sheet}:${ref.row}:${ref.col}`;
}

function isRangeNode(node: SimpleCellAddress | SimpleCellRange): node is SimpleCellRange {
  return 'start' in node;
}

function nodeKey(node: SimpleCellAddress | SimpleCellRange): string {
  if (isRangeNode(node)) {
    return `r:${node.start.sheet}:${node.start.row}:${node.start.col}:${node.end.row}:${node.end.col}`;
  }
  return `c:${node.sheet}:${node.row}:${node.col}`;
}

/**
 * HyperFormula parses raw strings as formulas, numbers, booleans, errors and
 * dates. A leading apostrophe forces the literal-text path and is stripped by
 * the engine, so text cells keep their text semantics for dependants.
 */
function rawContentFor(value: CellValue): RawCellContent {
  switch (value.type) {
    case 'number':
      return Number.isFinite(value.value) ? value.value : '#NUM!';
    case 'text':
      return `'${value.value}`;
    case 'boolean':
      return value.value;
    case 'error':
      // #CIRC! is an in-app-only error; it is never a valid engine input.
      if (value.value === '#CIRC!') return null;
      // The derived language teaches HyperFormula the #NULL! literal. Without
      // it, degrade to #VALUE! rather than treat #NULL! as text.
      if (value.value === '#NULL!' && !supportsNullError()) return '#VALUE!';
      return value.value;
    case 'blank':
      return null;
  }
}

/**
 * Write planned contents in rectangular blocks. Only planned addresses are
 * written: `load` arrives in chunks, so a block must never touch a cell its
 * chunk did not list. Each block is a hole-free band — a contiguous column
 * run extended over adjacent rows — so every cell inside a block rectangle
 * carries planned content and no empty cell is written anywhere.
 */
function writePlannedCells(engine: HyperFormula, writes: Map<string, PlannedWrite>): void {
  const sheets = new Map<number, Map<number, Map<number, RawCellContent>>>();
  for (const write of writes.values()) {
    const { sheet, row, col } = write.address;
    let rows = sheets.get(sheet);
    if (rows === undefined) {
      rows = new Map();
      sheets.set(sheet, rows);
    }
    let cols = rows.get(row);
    if (cols === undefined) {
      cols = new Map();
      rows.set(row, cols);
    }
    cols.set(col, write.content);
  }
  for (const [sheet, rows] of sheets) {
    for (const band of planBands(rows)) {
      const contents: RawCellContent[][] = [];
      for (let row = band.firstRow; row <= band.lastRow; row += 1) {
        const cols = rows.get(row);
        // Bands cover only rows that carry planned cells (see planBands), so
        // anything else is a packing bug; fail the block and let the caller
        // fall back to the per-cell path.
        if (cols === undefined) throw new Error('The write band covers an unplanned row');
        const line: RawCellContent[] = [];
        for (let col = band.firstCol; col <= band.lastCol; col += 1) {
          const content = cols.get(col);
          if (content === undefined) throw new Error('The write band covers an unplanned cell');
          line.push(content);
        }
        contents.push(line);
      }
      engine.setCellContents({ sheet, row: band.firstRow, col: band.firstCol }, contents);
    }
  }
}

/**
 * Pack one sheet's planned writes into hole-free bands. For each row the
 * contiguous column runs are found; a run over the same columns on the next
 * row extends its band. Rows and columns are visited in ascending order, so
 * every address inside a band rectangle is present in `rows`.
 */
function planBands(rows: Map<number, Map<number, RawCellContent>>): CellBand[] {
  const bands: CellBand[] = [];
  let open = new Map<string, CellBand>();
  const rowNumbers = [...rows.keys()].sort((left, right) => left - right);
  for (const row of rowNumbers) {
    const cols = rows.get(row);
    if (cols === undefined) continue;
    const nextOpen = new Map<string, CellBand>();
    for (const [firstCol, lastCol] of columnRuns(cols)) {
      const key = `${firstCol}:${lastCol}`;
      const band = open.get(key);
      if (band !== undefined && band.lastRow + 1 === row) {
        band.lastRow = row;
        nextOpen.set(key, band);
      } else {
        const fresh: CellBand = { firstRow: row, lastRow: row, firstCol, lastCol };
        bands.push(fresh);
        nextOpen.set(key, fresh);
      }
    }
    open = nextOpen;
  }
  return bands;
}

/** Contiguous column runs of one row: ascending and disjoint. */
function columnRuns(cols: Map<number, RawCellContent>): [number, number][] {
  const runs: [number, number][] = [];
  const colNumbers = [...cols.keys()].sort((left, right) => left - right);
  for (const col of colNumbers) {
    const last = runs[runs.length - 1];
    if (last !== undefined && col === last[1] + 1) last[1] = col;
    else runs.push([col, col]);
  }
  return runs;
}

function mapEngineValue(value: HyperFormulaCellValue): CellValue {
  if (value === null) return BLANK;
  switch (typeof value) {
    case 'number':
      return Number.isFinite(value) ? numberValue(value) : errorValue('#NUM!');
    case 'string':
      return textValue(value);
    case 'boolean':
      return booleanValue(value);
    case 'object':
      if (value instanceof DetailedCellError) return errorValue(errorCodeFor(value.type));
      return errorValue('#VALUE!');
    default:
      return BLANK;
  }
}

function errorCodeFor(type: ErrorType): CellErrorCode {
  // The derived language maps this custom type to the Excel error literal.
  if ((type as string) === 'NULL') return '#NULL!';
  switch (type) {
    case ErrorType.DIV_BY_ZERO:
      return '#DIV/0!';
    case ErrorType.NAME:
      return '#NAME?';
    case ErrorType.VALUE:
      return '#VALUE!';
    case ErrorType.NUM:
      return '#NUM!';
    case ErrorType.NA:
      return '#N/A';
    case ErrorType.CYCLE:
      return '#CIRC!';
    case ErrorType.REF:
      return '#REF!';
    case ErrorType.SPILL:
    case ErrorType.LIC:
    case ErrorType.ERROR:
      // Technical design §3.5: #SPILL!/#LIC!/#ERROR! never reach the model.
      return type === ErrorType.SPILL ? '#VALUE!' : '#NAME?';
  }
}

function now(): number {
  return typeof performance !== 'undefined' ? performance.now() : Date.now();
}
