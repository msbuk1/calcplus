import { describe, expect, it } from 'vitest';
import { HyperFormulaEngine } from '@engine/hyperformula-engine';
import type { EngineCellInput } from '@engine/port';
import { numberValue } from '@model/value';
import {
  CONTRACT_CORPUS,
  type CorpusShape,
  formulaCells,
  generateShape,
  touchedCells,
} from '../corpus/generate';
import {
  type BudgetGate,
  calibrateMachineSpeed,
  describeCalibration,
  describeFingerprint,
  describeGate,
  machineFingerprint,
  referenceDeclared,
  runBudgetProbe,
} from '@testing/machine';
import { buildCorpusWorkbook, type CorpusWorkbook } from './corpus';
import { type RegressionCheck, readHistory, recordMedian, writeReport } from './history';

/**
 * The Node-side benchmark suite (performance §6, layer 1).
 *
 *   npm run bench:model
 *   CALCPLUS_BENCH_SHAPES=tiny,fan-out npm run bench:model
 *
 * Every corpus shape is built from the committed generator with a fixed seed.
 * Each case records the build (start, load, finish), the incremental edit of
 * A1 and the steady heap, plus the machine fingerprint and the measured machine
 * speed. Warm-up runs are discarded; three runs are recorded and the median is
 * judged by the shared reference-machine gate: the absolute budget is enforced
 * only on a declared reference machine, and the per-change rule elsewhere is
 * the 20% regression check against the rolling median of the last ten local
 * runs (performance §7). A contended machine is reported but does not enter the
 * history.
 */

const CALIBRATION = calibrateMachineSpeed();
const FINGERPRINT = machineFingerprint();
const DECLARED_REFERENCE = referenceDeclared();
const HEAP_LIMIT_BYTES = 1.2 * 1024 ** 3;
const CONTENDED_RATIO = 1.5;

interface CaseResult {
  readonly shape: string;
  readonly rows: number;
  readonly cols: number;
  readonly sheets: number;
  readonly touchedCells: number;
  readonly formulas: number;
  readonly modelOnly: readonly string[];
  readonly buildMs: readonly number[];
  readonly buildMedianMs: number;
  readonly buildGate: BudgetGate;
  readonly incrementalMs: readonly number[] | null;
  readonly incrementalMedianMs: number | null;
  readonly incrementalGate: BudgetGate | null;
  readonly heapUsedBytes: number | null;
  readonly rssBytes: number | null;
  readonly gcForced: boolean;
  readonly buildRegression: RegressionCheck | null;
  readonly incrementalRegression: RegressionCheck | null;
}

const selected = selectedShapes();

describe('M7 corpus microbenchmarks', () => {
  for (const shape of selected) {
    it(`${shape.name}: build, incremental recalc and heap`, () => {
      const corpus = buildCorpusWorkbook(shape);
      const buildRuns: number[] = [];
      const warmups = 1;
      for (let warmup = 0; warmup < warmups; warmup += 1) runBuild(corpus);

      const buildGate = runBudgetProbe(
        {
          case: `${shape.name}-build`,
          budgetMs: buildBudget(shape),
          fingerprint: FINGERPRINT,
          calibration: CALIBRATION,
          declaredReference: DECLARED_REFERENCE,
        },
        () => {
          buildRuns.length = 0;
          for (let run = 0; run < 3; run += 1) buildRuns.push(runBuild(corpus));
          return buildRuns;
        },
      );

      const hasFormulas = corpus.formulas > 0;
      let incrementalRuns: number[] | null = null;
      let incrementalGate: BudgetGate | null = null;
      if (hasFormulas) {
        let measured: number[] = [];
        incrementalGate = runBudgetProbe(
          {
            case: `${shape.name}-incremental`,
            budgetMs: incrementalBudget(shape),
            fingerprint: FINGERPRINT,
            calibration: CALIBRATION,
            declaredReference: DECLARED_REFERENCE,
          },
          () => {
            measured = measureIncremental(corpus);
            return measured;
          },
        );
        incrementalRuns = measured;
      }

      const heap = measureHeap(corpus);

      const contended = CALIBRATION.contentionRatio > CONTENDED_RATIO;
      const buildRegression = recordMedian(
        FINGERPRINT,
        `${shape.name}-build`,
        buildGate.measuredMs,
        {
          contended,
        },
      );
      const incrementalRegression =
        incrementalGate === null
          ? null
          : recordMedian(FINGERPRINT, `${shape.name}-incremental`, incrementalGate.measuredMs, {
              contended,
            });

      const result: CaseResult = {
        shape: shape.name,
        rows: shape.rows,
        cols: shape.cols,
        sheets: shape.sheets ?? 1,
        touchedCells: corpus.touchedCells,
        formulas: corpus.formulas,
        modelOnly: corpus.modelOnly,
        buildMs: buildGate.runsMs,
        buildMedianMs: buildGate.measuredMs,
        buildGate,
        incrementalMs: incrementalRuns,
        incrementalMedianMs: incrementalGate?.measuredMs ?? null,
        incrementalGate,
        heapUsedBytes: heap.bytes,
        rssBytes: heap.rssBytes,
        gcForced: heap.gcForced,
        buildRegression,
        incrementalRegression,
      };
      report(result);

      expect(buildGate.passed, describeGate(buildGate)).toBe(true);
      if (incrementalGate !== null) {
        expect(incrementalGate.passed, describeGate(incrementalGate)).toBe(true);
      }
      if (heap.bytes !== null) {
        expect(heap.bytes, `${shape.name} heap ${formatBytes(heap.bytes)}`).toBeLessThan(
          HEAP_LIMIT_BYTES,
        );
      }
      expect(buildRegression?.blocked, buildRegression?.reason).not.toBe(true);
      expect(incrementalRegression?.blocked, incrementalRegression?.reason).not.toBe(true);
    }, 600_000);
  }

  it('records the corpus definition and the machine fingerprint', () => {
    const report = {
      harness: 'bench/model/run.bench.ts',
      milestone: 'M7',
      generatedAt: new Date().toISOString(),
      machine: FINGERPRINT,
      calibration: CALIBRATION,
      declaredReference: DECLARED_REFERENCE,
      shapes: CONTRACT_CORPUS.map((shape) => ({
        name: shape.name,
        rows: shape.rows,
        cols: shape.cols,
        sheets: shape.sheets ?? 1,
        touchedCells: touchedCells(shape),
        formulas: formulaCells(shape),
        pattern: shape.formulaPattern ?? 'none',
        cfRules: shape.cfRules ?? 0,
        catches: shape.catches,
      })),
    };
    const path = writeReport('corpus-definition.json', report);
    process.stdout.write(
      `\nCorpus: ${CONTRACT_CORPUS.length} shapes, ${describeFingerprint(FINGERPRINT)}\n` +
        `Machine speed: ${describeCalibration(CALIBRATION)}\n` +
        `Reference machine: ${DECLARED_REFERENCE ? 'declared' : 'not declared'}\n` +
        `Wrote ${path}\n`,
    );
    expect(report.shapes).toHaveLength(10);
  });
});

function runBuild(corpus: CorpusWorkbook): number {
  const engine = new HyperFormulaEngine();
  const started = performance.now();
  try {
    engine.start(corpus.workbook);
    engine.load(corpus.cells as readonly EngineCellInput[]);
    engine.finishLoad();
    return performance.now() - started;
  } finally {
    engine.dispose();
  }
}

/**
 * Time the incremental recalc after one edit of A1. The edit alternates values
 * so HyperFormula cannot short-circuit an unchanged input. The fan-out and
 * whole-column shapes make this the documented heavy case.
 */
function measureIncremental(corpus: CorpusWorkbook): number[] {
  const engine = new HyperFormulaEngine();
  try {
    engine.start(corpus.workbook);
    engine.load(corpus.cells);
    engine.finishLoad();
    engine.apply([literalEdit(2)]);
    const runs: number[] = [];
    for (let run = 0; run < 3; run += 1) {
      const started = performance.now();
      engine.apply([literalEdit(run % 2 === 0 ? 3 : 4)]);
      runs.push(performance.now() - started);
    }
    return runs;
  } finally {
    engine.dispose();
  }
}

function measureHeap(corpus: CorpusWorkbook): {
  bytes: number | null;
  rssBytes: number | null;
  gcForced: boolean;
} {
  const engine = new HyperFormulaEngine();
  try {
    engine.start(corpus.workbook);
    engine.load(corpus.cells);
    engine.finishLoad();
    const gc = (globalThis as { gc?: () => void }).gc;
    gc?.();
    const memory = process.memoryUsage();
    // The working set is the JS heap plus typed arrays; the resident set is a
    // coarse upper bound that also sees allocator overhead.
    return {
      bytes: memory.heapUsed + memory.arrayBuffers,
      rssBytes: memory.rss,
      gcForced: typeof gc === 'function',
    };
  } finally {
    engine.dispose();
  }
}

function literalEdit(value: number) {
  return { sheet: 0, row: 0, col: 0, formula: null, value: numberValue(value) };
}

function buildBudget(shape: CorpusShape): number {
  // Performance §3.2 budgets the full recalculation of the 250k guaranteed
  // shape at 1.5 s and the 1M degraded shape at 5 s. The cold build of a
  // fan-out shape is measured, not budgeted: only the pathological ceiling has
  // a "must finish" bound so a hang still fails.
  if (shape.name === 'pathological') return 60_000;
  if (shape.name === 'degraded-1m') return 5_000;
  return 1_500;
}

function incrementalBudget(shape: CorpusShape): number {
  if (shape.name === 'pathological') return 5_000;
  if (shape.formulaPattern === 'fan-out-single' || shape.formulaPattern === 'mixed') {
    return formulaCells(shape) > 50_000 ? 800 : 400;
  }
  if (shape.formulaPattern === 'whole-column') return 5_000;
  return 400;
}

function selectedShapes(): readonly CorpusShape[] {
  const requested = process.env.CALCPLUS_BENCH_SHAPES;
  if (requested === undefined || requested.trim() === '') return CONTRACT_CORPUS;
  return requested
    .split(',')
    .map((name) => name.trim())
    .filter((name) => name !== '')
    .map((name) => generateShape(name));
}

function report(result: CaseResult): void {
  const heap = result.heapUsedBytes === null ? 'n/a' : formatBytes(result.heapUsedBytes);
  const rss = result.rssBytes === null ? 'n/a' : formatBytes(result.rssBytes);
  const incremental =
    result.incrementalMedianMs === null ? '-' : `${result.incrementalMedianMs.toFixed(1)} ms`;
  process.stdout.write(
    [
      `${result.shape.padEnd(14)}`,
      `touched ${String(result.touchedCells).padStart(9)}`,
      `formulas ${String(result.formulas).padStart(8)}`,
      `build ${result.buildMedianMs.toFixed(1).padStart(8)} ms (${result.buildGate.mode})`,
      `edit ${incremental.padStart(11)}${result.incrementalGate === null ? '' : ` (${result.incrementalGate.mode})`}`,
      `heap ${heap.padStart(9)}${result.gcForced ? '' : ' (no forced GC)'}`,
      `rss ${rss.padStart(9)}`,
      result.modelOnly.length === 0 ? '' : `model-only: ${result.modelOnly.join(', ')}`,
    ]
      .filter((part) => part !== '')
      .join('  ') + '\n',
  );
  const path = writeReport(`latest-${result.shape}.json`, {
    ...result,
    machine: FINGERPRINT,
    calibration: CALIBRATION,
    declaredReference: DECLARED_REFERENCE,
    rollingHistory: readHistory(FINGERPRINT)[`${result.shape}-build`] ?? null,
  });
  process.stdout.write(`  wrote ${path}\n`);
}

function formatBytes(bytes: number): string {
  return `${(bytes / 1024 ** 2).toFixed(0)} MB`;
}
