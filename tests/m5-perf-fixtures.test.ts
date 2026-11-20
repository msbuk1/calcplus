import { describe, expect, it } from 'vitest';
import { ADVISORY_LIVENESS_FACTOR, describeGate } from '@testing/machine';
import {
  CF_RANGE_STATISTICS_100K_BUDGET_MS,
  CF_VISIBLE_EVALUATION_BUDGET_MS,
  FILTER_100K_BUDGET_MS,
  SORT_100K_BUDGET_MS,
  describeFingerprint,
  machineFingerprint,
  measureCfStatistics100k,
  measureCfVisible100k,
  measureFilter100k,
  measureSort100k,
} from './helpers/perf-fixtures';

/**
 * M5 performance fixtures F17 and CF18.
 *
 *   npm run test -- tests/m5-perf-fixtures.test.ts
 *
 * F17 `filter-100k` (sort-autofilter §4): a 100k-row sort completes within
 * 2.0 s end to end (plan, apply and one transaction), and filter apply, change
 * and clear each complete within 1.0 s. CF18 `cf-100k`
 * (conditional-formatting §9): range statistics over 100k cells complete within
 * 200 ms and a visible frame evaluates within 4 ms.
 *
 * The shapes are generated from a fixed seed in `tests/helpers/perf-fixtures.ts`,
 * never committed, and every measurement follows the shared reference-machine
 * gate: one warm-up run is discarded, three runs are recorded, and the verdict
 * says whether this run may enforce the absolute budget. On a declared
 * reference machine a miss gets the one immediate re-run performance §7
 * allows; elsewhere the fingerprint and machine speed are reported and the
 * number is advisory (performance §6), with 5x the budget as the liveness
 * ceiling so a stopped implementation still fails.
 *
 * Baseline recorded 2026-11-18 on an AMD Ryzen 5 3600 (6 cores / 12 threads),
 * 31.3 GiB RAM, Linux, Node 24.15.0: sort 927 ms; filter apply/change/clear
 * 88/127/50 ms; CF range statistics 8-22 ms; visible frame 1.52 ms. Each run
 * recomputes and reports its own fingerprint, so the note is context, not a
 * gate.
 */
describe('F17 filter-100k', () => {
  it('sorts 100k rows within 2.0 seconds end to end', () => {
    const result = measureSort100k({ warmupRuns: 1, runs: 3 });
    const fingerprint = describeFingerprint(result.fingerprint);
    expect(result.rows).toBe(100_000);
    expect(result.columns).toBe(2);
    expect(result.touchedCells).toBe(200_002);
    expect(result.runsMs).toHaveLength(3);
    expect(result.transactionsPerRun).toEqual([1, 1, 1]);
    expect(
      result.verdict.passed,
      `sort median ${result.medianMs.toFixed(1)} ms (runs ${result.runsMs
        .map((run) => run.toFixed(1))
        .join(', ')} ms) against ${SORT_100K_BUDGET_MS} ms on ${fingerprint}. ${describeGate(
        result.verdict,
      )}`,
    ).toBe(true);
    if (result.verdict.mode === 'enforced') {
      expect(result.medianMs).toBeLessThanOrEqual(SORT_100K_BUDGET_MS);
    }
  }, 300_000);

  it('applies, changes and clears a filter on 100k rows within 1.0 second each', () => {
    const result = measureFilter100k({ warmupRuns: 1, runs: 3 });
    const fingerprint = describeFingerprint(result.fingerprint);
    expect(result.rows).toBe(100_000);
    expect(result.actions.map((action) => action.action)).toEqual(['apply', 'change', 'clear']);
    for (const action of result.actions) {
      expect(action.runsMs, action.action).toHaveLength(3);
      expect(action.transactionsPerRun, action.action).toEqual([1, 1, 1]);
      expect(
        action.verdict.passed,
        `filter ${action.action} median ${action.medianMs.toFixed(1)} ms (runs ${action.runsMs
          .map((run) => run.toFixed(1))
          .join(', ')} ms) against ${FILTER_100K_BUDGET_MS} ms on ${fingerprint}. ${describeGate(
          action.verdict,
        )}`,
      ).toBe(true);
      if (action.verdict.mode === 'enforced') {
        expect(action.medianMs).toBeLessThanOrEqual(FILTER_100K_BUDGET_MS);
      }
    }
  }, 300_000);
});

describe('CF18 cf-100k', () => {
  it('computes 100k-cell range statistics within 200 ms per rule kind', () => {
    const result = measureCfStatistics100k({ warmupRuns: 1, runs: 3 });
    const fingerprint = describeFingerprint(result.fingerprint);
    expect(result.cells).toBe(100_000);
    expect(result.statistics.map((entry) => entry.kind)).toEqual([
      'top10',
      'duplicate',
      'colorScale',
      'dataBar',
      'iconSet',
    ]);
    for (const statistic of result.statistics) {
      expect(statistic.runsMs, statistic.kind).toHaveLength(3);
      expect(
        statistic.verdict.passed,
        `${statistic.kind} median ${statistic.medianMs.toFixed(1)} ms (runs ${statistic.runsMs
          .map((run) => run.toFixed(1))
          .join(
            ', ',
          )} ms) against ${CF_RANGE_STATISTICS_100K_BUDGET_MS} ms on ${fingerprint}. ${describeGate(
          statistic.verdict,
        )}`,
      ).toBe(true);
      if (statistic.verdict.mode === 'enforced') {
        expect(statistic.medianMs).toBeLessThanOrEqual(CF_RANGE_STATISTICS_100K_BUDGET_MS);
      }
    }
    expect(
      result.worstMedianMs,
      `worst range-statistic median ${result.worstMedianMs.toFixed(1)} ms on ${fingerprint}`,
    ).toBeLessThanOrEqual(CF_RANGE_STATISTICS_100K_BUDGET_MS * ADVISORY_LIVENESS_FACTOR);
  }, 300_000);

  it('evaluates a visible frame of the 50-rule sheet within 4 ms', () => {
    const result = measureCfVisible100k({ warmupRuns: 1, runs: 3 });
    const fingerprint = describeFingerprint(result.fingerprint);
    expect(result.visibleCells).toBe(800);
    expect(result.visibleRulesPerCell).toBe(2);
    expect(result.paintedCells).toBeGreaterThan(0);
    expect(result.runsMs).toHaveLength(3);
    expect(
      result.verdict.passed,
      `visible-frame median ${result.medianMs.toFixed(2)} ms (runs ${result.runsMs
        .map((run) => run.toFixed(2))
        .join(
          ', ',
        )} ms) over ${result.visibleCells} cells and ${result.paintedCells} painted, against ${CF_VISIBLE_EVALUATION_BUDGET_MS} ms on ${fingerprint}. ${describeGate(
        result.verdict,
      )}`,
    ).toBe(true);
    if (result.verdict.mode === 'enforced') {
      expect(result.medianMs).toBeLessThanOrEqual(CF_VISIBLE_EVALUATION_BUDGET_MS);
    }
  }, 300_000);
});

describe('machine fingerprint', () => {
  it('records the CPU, core counts, RAM, OS and Node version', () => {
    const fingerprint = machineFingerprint();
    expect(fingerprint.cpuModel).not.toBe('');
    expect(fingerprint.logicalCpus).toBeGreaterThan(0);
    expect(fingerprint.totalMemoryBytes).toBeGreaterThan(0);
    expect(fingerprint.platform).not.toBe('');
    expect(fingerprint.osName).not.toBe('');
    expect(fingerprint.nodeVersion).toMatch(/^v\d+/);
  });
});
