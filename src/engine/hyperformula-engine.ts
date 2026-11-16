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

  load(cells: readonly EngineCellInput[]): void {
    const engine = this.requireEngine();
    for (const cell of cells) this.setInput(engine, cell);
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
    this.suspended = false;
    if (engine !== null) {
      try {
        engine.destroy();
      } catch {
        // Destroying an already broken engine must not mask the original error.
      }
    }
  }

  private setInput(engine: HyperFormula, cell: EngineCellInput): void {
    const address = addressOf(cell);
    const key = keyOf(cell);

    if (cell.staleInput !== true && cell.formula !== null) {
      const prepared = prepareFormula(cell.formula);
      if (prepared.text !== null) {
        const candidate = `=${prepared.text}`;
        if (this.validate(engine, candidate)) {
          try {
            engine.setCellContents(address, candidate);
            this.formulaCells.set(key, refOf(cell));
            this.skippedCells.delete(key);
            return;
          } catch {
            // Fall through to the cached-value literal.
          }
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
    this.setLiteral(engine, address, cell.value);
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

  private validate(engine: HyperFormula, formula: string): boolean {
    try {
      return engine.validateFormula(formula);
    } catch {
      return false;
    }
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
