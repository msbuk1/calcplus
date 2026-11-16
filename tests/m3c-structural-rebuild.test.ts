import { describe, expect, it } from 'vitest';
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
 */
describe('M3C structural-edit graph rebuild budget', () => {
  it('rebuilds a 250k-cell, 50k-formula dependency graph within 1.5 seconds', () => {
    const result = measureStructuralRebuildGraph({ warmupRuns: 1, runs: 3 });
    expect(result.touchedCells).toBe(250_000);
    expect(result.formulas).toBe(50_000);
    expect(result.runsMs).toHaveLength(3);
    expect(result.medianMs).toBeLessThanOrEqual(STRUCTURAL_REBUILD_BUDGET_MS);
  }, 120_000);
});
