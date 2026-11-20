import { describe, expect, it } from 'vitest';
import {
  DEGRADED_FORMULAS,
  DEGRADED_TOUCHED_CELLS,
  GUARANTEED_FORMULAS,
  GUARANTEED_TOUCHED_CELLS,
  GUARANTEED_WORKING_SET_BYTES,
  MEMORY_REFUSE_BYTES,
  MEMORY_WARNING_BYTES,
  classifyEnvelope,
  degradationPlan,
  describeDegradation,
  shouldConfirmLargeOperation,
  shouldRefuseLargeOperation,
} from '@engine/degradation';

/**
 * The degradation ladder boundaries (performance §8).
 *
 * Rendering degrades before correctness does. Every plan must keep
 * `valuesPreserved` true, and a refused workbook must say plainly that the file
 * was untouched. The memory thresholds are exact: 500 MB is the guaranteed
 * envelope, 1 GB warns, 1.2 GB refuses new large operations.
 */
describe('size envelopes', () => {
  it('keeps exactly 250,000 touched cells and 50,000 formulas guaranteed', () => {
    const plan = degradationPlan({
      touchedCells: GUARANTEED_TOUCHED_CELLS,
      formulas: GUARANTEED_FORMULAS,
    });
    expect(plan.envelope).toBe('guaranteed');
    expect(plan.banner).toBeNull();
    expect(plan.conditionalFormatScope).toBe('full');
    expect(plan.scrollOverscan).toBe('full');
    expect(plan.volatileRecalculation).toBe('immediate');
    expect(plan.notices).toEqual([]);
  });

  it('degrades one touched cell past the guarantee, in the documented order', () => {
    const plan = degradationPlan({ touchedCells: GUARANTEED_TOUCHED_CELLS + 1, formulas: 0 });
    expect(plan.envelope).toBe('degraded');
    expect(plan.banner).toMatchObject({ level: 'info' });
    expect(plan.banner?.message).toContain('Large workbook');
    expect(plan.notices[0]).toContain('Large workbook');
    expect(plan.conditionalFormatScope).toBe('visible');
    expect(plan.scrollOverscan).toBe('reduced');
    expect(plan.volatileRecalculation).toBe('idle');
    expect(plan.memoryWarning).toBe(false);
    expect(plan.refuseLargeOperations).toBe(false);
    expect(plan.valuesPreserved).toBe(true);
  });

  it('degrades one formula past the guarantee', () => {
    const plan = degradationPlan({ touchedCells: 1, formulas: GUARANTEED_FORMULAS + 1 });
    expect(plan.envelope).toBe('degraded');
    expect(plan.conditionalFormatScope).toBe('visible');
  });

  it('keeps the degraded envelope up to the documented ceiling', () => {
    expect(classifyEnvelope(DEGRADED_TOUCHED_CELLS, DEGRADED_FORMULAS)).toBe('degraded');
    expect(classifyEnvelope(DEGRADED_TOUCHED_CELLS - 1, DEGRADED_FORMULAS - 1)).toBe('degraded');
    expect(degradationPlan({ touchedCells: 500_000, formulas: 100_000 }).envelope).toBe('degraded');
  });

  it('refuses one touched cell past one million without touching the file', () => {
    const plan = degradationPlan({ touchedCells: DEGRADED_TOUCHED_CELLS + 1, formulas: 0 });
    expect(plan.envelope).toBe('refused');
    expect(plan.refusesWorkbook).toBe(true);
    expect(plan.refusesOpen).toBe(true);
    expect(plan.banner).toMatchObject({ level: 'critical' });
    expect(plan.banner?.message).toContain('1,000,000');
    expect(plan.banner?.message).toContain('was not modified');
    expect(plan.valuesPreserved).toBe(true);
    // Refusal is not a rendering step: no promise of a reduced mode.
    expect(plan.conditionalFormatScope).toBe('visible');
  });

  it('refuses one formula past 200,000', () => {
    const plan = degradationPlan({ touchedCells: 0, formulas: DEGRADED_FORMULAS + 1 });
    expect(plan.envelope).toBe('refused');
    expect(plan.refusesWorkbook).toBe(true);
    expect(plan.banner?.message).toContain('200,000');
  });
});

describe('memory ladder', () => {
  it('stays quiet below one gigabyte', () => {
    const plan = degradationPlan({
      touchedCells: 1_000,
      formulas: 0,
      estimatedWorkingSetBytes: MEMORY_WARNING_BYTES - 1,
    });
    expect(plan.memoryWarning).toBe(false);
    expect(plan.refuseLargeOperations).toBe(false);
    expect(plan.promptToClose).toBe(false);
    expect(plan.banner).toBeNull();
    expect(shouldConfirmLargeOperation(plan)).toBe(false);
  });

  it('warns at exactly one gigabyte and asks before a large operation', () => {
    const plan = degradationPlan({
      touchedCells: 1_000,
      formulas: 0,
      estimatedWorkingSetBytes: MEMORY_WARNING_BYTES,
    });
    expect(plan.memoryWarning).toBe(true);
    expect(plan.refuseLargeOperations).toBe(false);
    expect(plan.banner).toMatchObject({ level: 'warning' });
    expect(plan.banner?.message).toContain('1 GB');
    expect(shouldConfirmLargeOperation(plan)).toBe(true);
  });

  it('refuses new large operations at exactly 1.2 gigabytes', () => {
    const plan = degradationPlan({
      touchedCells: 1_000,
      formulas: 0,
      estimatedWorkingSetBytes: MEMORY_REFUSE_BYTES,
    });
    expect(plan.memoryWarning).toBe(true);
    expect(plan.refuseLargeOperations).toBe(true);
    expect(plan.promptToClose).toBe(true);
    expect(plan.refusesOpen).toBe(false);
    expect(plan.banner).toMatchObject({ level: 'critical' });
    expect(plan.banner?.message).toContain('1.2 GB');
    expect(shouldRefuseLargeOperation(plan)).toBe(true);
    expect(shouldConfirmLargeOperation(plan)).toBe(true);
  });

  it('refuses an open above 1.2 gigabytes before any import', () => {
    const plan = degradationPlan({
      touchedCells: 1_000,
      formulas: 0,
      estimatedWorkingSetBytes: MEMORY_REFUSE_BYTES + 1,
      phase: 'open',
    });
    expect(plan.refusesOpen).toBe(true);
    expect(plan.refusesWorkbook).toBe(false);
    expect(plan.banner?.message).toContain('1.2 GB');
  });

  it('warns but does not refuse a runtime workbook at exactly 1.2 gigabytes', () => {
    const plan = degradationPlan({
      touchedCells: 1_000,
      formulas: 0,
      estimatedWorkingSetBytes: MEMORY_REFUSE_BYTES,
      phase: 'runtime',
    });
    expect(plan.refusesOpen).toBe(false);
    expect(plan.refuseLargeOperations).toBe(true);
  });

  it('puts the size notice first and the memory notice after it', () => {
    const plan = degradationPlan({
      touchedCells: 300_000,
      formulas: 10_000,
      estimatedWorkingSetBytes: MEMORY_WARNING_BYTES,
    });
    expect(plan.notices).toHaveLength(2);
    expect(plan.notices[0]).toContain('Large workbook');
    expect(plan.notices[1]).toContain('1 GB');
    expect(plan.banner?.level).toBe('warning');
  });

  it('treats the guaranteed working set as below the warning line', () => {
    const plan = degradationPlan({
      touchedCells: GUARANTEED_TOUCHED_CELLS,
      formulas: GUARANTEED_FORMULAS,
      estimatedWorkingSetBytes: GUARANTEED_WORKING_SET_BYTES,
    });
    expect(plan.memoryWarning).toBe(false);
    expect(plan.envelope).toBe('guaranteed');
  });
});

describe('the ladder never claims lost values', () => {
  it('keeps valuesPreserved true at every step', () => {
    const steps = [
      degradationPlan({ touchedCells: 1, formulas: 0 }),
      degradationPlan({ touchedCells: 300_000, formulas: 0 }),
      degradationPlan({ touchedCells: 900_000, formulas: 0 }),
      degradationPlan({
        touchedCells: 1_000,
        formulas: 0,
        estimatedWorkingSetBytes: MEMORY_REFUSE_BYTES,
      }),
      degradationPlan({ touchedCells: DEGRADED_TOUCHED_CELLS + 1, formulas: 0 }),
    ];
    for (const plan of steps) {
      expect(plan.valuesPreserved).toBe(true);
      expect(describeDegradation(plan)).not.toContain('dropped');
    }
  });

  it('describes the reduced extras without claiming a different result', () => {
    const plan = degradationPlan({ touchedCells: 300_000, formulas: 0 });
    expect(describeDegradation(plan)).toBe(
      'envelope degraded, conditional formatting visible-only, reduced overscan, volatile recalculation on idle, values preserved',
    );
  });
});
