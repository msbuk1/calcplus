import { HyperFormulaEngine } from '@engine/hyperformula-engine';
import type { EngineCellInput } from '@engine/port';
import { BLANK, numberValue } from '@model/value';
import {
  type BudgetGate,
  type MachineCalibration,
  type MachineFingerprint,
  machineCalibration,
  machineFingerprint,
  runBudgetProbe,
} from './machine';

/**
 * Focused probe for performance §3.2: structural edit, graph rebuild, 250k
 * touched cells.
 *
 * It measures the same path `CalculationController.rebuild()` uses after an
 * insert/delete: build the engine, load the workbook snapshot and finish the
 * batch. Shape generation is outside the timed section. Run it from the
 * focused gate test:
 *
 *   npm run test -- tests/m3c-structural-rebuild.test.ts
 *
 * The budget is a contract with the reference machine (performance §1). The
 * verdict says whether this run may enforce it: on a declared reference
 * machine the median must be at or under 1.5 s, and a miss gets the one
 * immediate re-run performance §7 allows. Everywhere else the number, the
 * fingerprint and the measured machine speed are reported and the budget is a
 * liveness ceiling, never a pass/fail claim the hardware cannot support.
 */

export const STRUCTURAL_REBUILD_BUDGET_MS = 1500;

export interface StructuralRebuildProbeOptions {
  readonly rows?: number;
  readonly columns?: number;
  readonly formulaColumns?: number;
  readonly warmupRuns?: number;
  readonly runs?: number;
  readonly date1904?: boolean;
  readonly fingerprint?: MachineFingerprint;
  readonly calibration?: MachineCalibration;
  readonly declaredReference?: boolean;
}

export interface StructuralRebuildProbeResult {
  readonly rows: number;
  readonly columns: number;
  readonly touchedCells: number;
  readonly formulas: number;
  readonly runsMs: readonly number[];
  readonly bestMs: number;
  readonly medianMs: number;
  readonly budgetMs: number;
  readonly fingerprint: MachineFingerprint;
  readonly calibration: MachineCalibration;
  readonly verdict: BudgetGate;
}

export function measureStructuralRebuildGraph(
  options: StructuralRebuildProbeOptions = {},
): StructuralRebuildProbeResult {
  const rows = options.rows ?? 5_000;
  const columns = options.columns ?? 50;
  const formulaColumns = options.formulaColumns ?? 10;
  const warmupRuns = Math.max(0, options.warmupRuns ?? 1);
  const runs = Math.max(1, options.runs ?? 3);
  const date1904 = options.date1904 ?? false;
  if (rows < 1 || columns < 2 || formulaColumns < 1 || formulaColumns >= columns) {
    throw new RangeError('The structural rebuild probe shape is invalid');
  }

  const cells = buildShape(rows, columns, formulaColumns);
  const fingerprint = options.fingerprint ?? machineFingerprint();
  const calibration = options.calibration ?? machineCalibration();

  const verdict = runBudgetProbe(
    {
      case: 'structural-rebuild-250k',
      budgetMs: STRUCTURAL_REBUILD_BUDGET_MS,
      fingerprint,
      calibration,
      declaredReference: options.declaredReference,
    },
    () => {
      for (let warmup = 0; warmup < warmupRuns; warmup += 1) {
        runRebuild(cells, date1904);
      }
      const measured: number[] = [];
      for (let run = 0; run < runs; run += 1) {
        measured.push(runRebuild(cells, date1904));
      }
      return measured;
    },
  );

  const runsMs = verdict.runsMs;
  return {
    rows,
    columns,
    touchedCells: rows * columns,
    formulas: rows * formulaColumns,
    runsMs,
    bestMs: Math.min(...runsMs),
    medianMs: verdict.measuredMs,
    budgetMs: STRUCTURAL_REBUILD_BUDGET_MS,
    fingerprint,
    calibration,
    verdict,
  };
}

function runRebuild(cells: readonly EngineCellInput[], date1904: boolean): number {
  const engine = new HyperFormulaEngine();
  const started = now();
  try {
    engine.start({ date1904, sheetNames: ['Sheet1'] });
    engine.load(cells);
    engine.finishLoad();
    return now() - started;
  } finally {
    engine.dispose();
  }
}

function buildShape(
  rows: number,
  columns: number,
  formulaColumns: number,
): readonly EngineCellInput[] {
  const cells: EngineCellInput[] = [];
  for (let row = 0; row < rows; row += 1) {
    for (let col = 0; col < columns; col += 1) {
      if (col > 0 && col <= formulaColumns) {
        cells.push({
          sheet: 0,
          row,
          col,
          formula: `A${row + 1}*2`,
          value: BLANK,
        });
        continue;
      }
      cells.push({
        sheet: 0,
        row,
        col,
        formula: null,
        value: numberValue(col === 0 ? row + 1 : (row + col) % 1000),
      });
    }
  }
  return cells;
}

function now(): number {
  return typeof performance !== 'undefined' ? performance.now() : Date.now();
}
