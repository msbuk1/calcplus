import { HyperFormulaEngine } from '@engine/hyperformula-engine';
import type { EngineCellInput } from '@engine/port';
import { BLANK, numberValue } from '@model/value';

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
 * The test asserts the guaranteed envelope budget (1.5 s) on the reference
 * machine. The returned run list lets a benchmark report the distribution
 * instead of only pass/fail.
 */

export const STRUCTURAL_REBUILD_BUDGET_MS = 1500;

export interface StructuralRebuildProbeOptions {
  readonly rows?: number;
  readonly columns?: number;
  readonly formulaColumns?: number;
  readonly warmupRuns?: number;
  readonly runs?: number;
  readonly date1904?: boolean;
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
  for (let run = 0; run < warmupRuns; run += 1) {
    runRebuild(cells, date1904);
  }
  const runsMs: number[] = [];
  for (let run = 0; run < runs; run += 1) {
    runsMs.push(runRebuild(cells, date1904));
  }
  const sorted = [...runsMs].sort((left, right) => left - right);
  return {
    rows,
    columns,
    touchedCells: rows * columns,
    formulas: rows * formulaColumns,
    runsMs,
    bestMs: sorted[0] ?? 0,
    medianMs: sorted[Math.floor(sorted.length / 2)] ?? 0,
    budgetMs: STRUCTURAL_REBUILD_BUDGET_MS,
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
