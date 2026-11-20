import { describe, expect, it } from 'vitest';
import { describeGate } from '@testing/machine';
import {
  STRUCTURAL_REBUILD_BUDGET_MS,
  measureStructuralRebuildGraph,
} from '@testing/structural-rebuild';

/**
 * Gate measurement for performance §3.2:
 *
 *   npm run test -- tests/m3c-structural-rebuild.test.ts
 *
 * This is the focused probe, not the full suite. It times the same engine
 * rebuild path used after a structural edit (build, load, finish) for 250,000
 * touched cells and 50,000 formulas. One warm-up run is discarded; three runs
 * are recorded and the median is asserted against the guaranteed envelope.
 *
 * The numbers are only a contract on the reference machine (performance §1,
 * §6). On every other machine the probe records the fingerprint and the
 * machine-speed sample and treats 1.5 s as a liveness ceiling: that keeps the
 * suite honest on a slow or busy laptop instead of flaking, and a declared
 * reference run still enforces the budget with the one immediate re-run that
 * performance §7 allows.
 */
describe('M3C structural-edit graph rebuild budget', () => {
  it('rebuilds a 250k-cell, 50k-formula dependency graph within 1.5 seconds', () => {
    const result = measureStructuralRebuildGraph({ warmupRuns: 1, runs: 3 });
    if (process.env.CALCPLUS_PERF_VERBOSE === '1') console.log(describeGate(result.verdict));
    expect(result.touchedCells).toBe(250_000);
    expect(result.formulas).toBe(50_000);
    expect(result.runsMs).toHaveLength(3);
    expect(result.verdict.case).toBe('structural-rebuild-250k');
    expect(result.verdict.budgetMs).toBe(STRUCTURAL_REBUILD_BUDGET_MS);
    // The gate fails only when the contract actually applies to this run.
    expect(result.verdict.passed, describeGate(result.verdict)).toBe(true);
    if (result.verdict.mode === 'enforced') {
      expect(result.medianMs).toBeLessThanOrEqual(STRUCTURAL_REBUILD_BUDGET_MS);
    } else {
      // Performance §6: a non-reference result is advisory, but it still has
      // to finish inside the documented liveness ceiling.
      expect(result.verdict.reference.matches).toBe(false);
      expect(result.medianMs).toBeLessThanOrEqual(result.verdict.livenessCeilingMs);
    }
  }, 300_000);
});
