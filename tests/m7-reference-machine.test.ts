import { describe, expect, it } from 'vitest';
import {
  ADVISORY_LIVENESS_FACTOR,
  MEDIAN_REGRESSION_FACTOR,
  type MachineCalibration,
  type MachineFingerprint,
  calibrateMachineSpeed,
  describeGate,
  referenceStatus,
  runBudgetProbe,
} from '@testing/machine';

/**
 * The reference-machine gate (performance §1, §6, §7).
 *
 * These tests use synthetic fingerprints and measurements, so they prove the
 * policy itself: the absolute budget is enforced only on a declared reference
 * machine; every other run is advisory but still has a liveness ceiling; and an
 * enforced miss gets exactly one immediate re-run.
 */
const REFERENCE: MachineFingerprint = {
  cpuModel: 'Mid-range 2023 laptop CPU',
  logicalCpus: 8,
  physicalCores: 4,
  totalMemoryBytes: 16 * 1024 ** 3,
  platform: 'linux',
  osName: 'Ubuntu',
  osVersion: '24.04',
  nodeVersion: 'v24.15.0',
};

const CALIBRATION: MachineCalibration = {
  units: 20_000_000,
  samplesMs: [100, 101, 102],
  medianMs: 101,
  bestMs: 100,
  unitsPerMs: 198_019,
  bestUnitsPerMs: 200_000,
  contentionRatio: 1.01,
};

describe('reference machine class', () => {
  it('matches a declared mid-range laptop with four physical cores', () => {
    const status = referenceStatus(REFERENCE, true);
    expect(status.matches).toBe(true);
    expect(status.reasons).toEqual([]);
  });

  it('does not match an undeclared run even on the right hardware', () => {
    const status = referenceStatus(REFERENCE, false);
    expect(status.matches).toBe(false);
    expect(status.reasons.join(' ')).toContain('CALCPLUS_REFERENCE_MACHINE=1');
  });

  it('rejects too few physical cores, too little RAM and an old OS', () => {
    expect(referenceStatus({ ...REFERENCE, physicalCores: 2 }, true).reasons.join(' ')).toContain(
      'physical cores',
    );
    expect(
      referenceStatus({ ...REFERENCE, totalMemoryBytes: 8 * 1024 ** 3 }, true).reasons.join(' '),
    ).toContain('GiB of RAM');
    expect(referenceStatus({ ...REFERENCE, osVersion: '20.04' }, true).reasons.join(' ')).toContain(
      'Ubuntu 22.04',
    );
    expect(
      referenceStatus(
        { ...REFERENCE, platform: 'win32', osName: 'Windows 10', osVersion: '10.0.19045' },
        true,
      ).reasons.join(' '),
    ).toContain('Windows 11');
  });
});

describe('budget gate', () => {
  it('enforces the absolute budget on the declared reference machine', () => {
    const gate = runBudgetProbe(
      {
        case: 'synthetic',
        budgetMs: 100,
        fingerprint: REFERENCE,
        calibration: CALIBRATION,
        declaredReference: true,
      },
      () => [90, 95, 100],
    );
    expect(gate.mode).toBe('enforced');
    expect(gate.passed).toBe(true);
    expect(gate.retried).toBe(false);
    expect(gate.measuredMs).toBe(95);
  });

  it('re-runs one miss and accepts the second measurement', () => {
    let runs = 0;
    const gate = runBudgetProbe(
      {
        case: 'synthetic',
        budgetMs: 100,
        fingerprint: REFERENCE,
        calibration: CALIBRATION,
        declaredReference: true,
      },
      () => {
        runs += 1;
        return runs === 1 ? [140, 150, 160] : [80, 90, 95];
      },
    );
    expect(runs).toBe(2);
    expect(gate.retried).toBe(true);
    expect(gate.passed).toBe(true);
    expect(gate.measuredMs).toBe(90);
  });

  it('fails a second enforced miss', () => {
    let runs = 0;
    const gate = runBudgetProbe(
      {
        case: 'synthetic',
        budgetMs: 100,
        fingerprint: REFERENCE,
        calibration: CALIBRATION,
        declaredReference: true,
      },
      () => {
        runs += 1;
        return [140, 150, 160];
      },
    );
    expect(runs).toBe(2);
    expect(gate.retried).toBe(true);
    expect(gate.passed).toBe(false);
    expect(gate.reason).toContain('after one immediate re-run');
  });

  it('keeps an undeclared run advisory while it is inside the liveness ceiling', () => {
    const gate = runBudgetProbe(
      {
        case: 'synthetic',
        budgetMs: 100,
        fingerprint: REFERENCE,
        calibration: CALIBRATION,
        declaredReference: false,
      },
      () => [280, 300, 320],
    );
    expect(gate.mode).toBe('advisory');
    expect(gate.withinBudget).toBe(false);
    expect(gate.passed).toBe(true);
    expect(gate.livenessCeilingMs).toBe(100 * ADVISORY_LIVENESS_FACTOR);
    expect(describeGate(gate)).toContain('advisory');
  });

  it('fails an undeclared run beyond the liveness ceiling', () => {
    const gate = runBudgetProbe(
      {
        case: 'synthetic',
        budgetMs: 100,
        fingerprint: REFERENCE,
        calibration: CALIBRATION,
        declaredReference: false,
      },
      () => [2_000, 2_100, 2_200],
    );
    expect(gate.mode).toBe('advisory');
    expect(gate.passed).toBe(false);
    expect(describeGate(gate)).toContain('liveness ceiling');
  });

  it('treats a two-core machine as advisory even when declared', () => {
    const gate = runBudgetProbe(
      {
        case: 'synthetic',
        budgetMs: 100,
        fingerprint: { ...REFERENCE, physicalCores: 2 },
        calibration: CALIBRATION,
        declaredReference: true,
      },
      () => [2_000, 2_000, 2_000],
    );
    expect(gate.mode).toBe('advisory');
    expect(gate.reference.matches).toBe(false);
    expect(gate.passed).toBe(false);
  });

  it('reports the machine speed with every gate', () => {
    const gate = runBudgetProbe(
      {
        case: 'synthetic',
        budgetMs: 100,
        fingerprint: REFERENCE,
        calibration: CALIBRATION,
        declaredReference: true,
      },
      () => [90, 95, 100],
    );
    expect(gate.calibration.unitsPerMs).toBeGreaterThan(0);
    expect(gate.fingerprint.cpuModel).toBe(REFERENCE.cpuModel);
    expect(describeGate(gate)).toContain('units/ms');
  });
});

describe('machine-speed sample', () => {
  it('measures a positive throughput and a contention ratio', () => {
    const calibration = calibrateMachineSpeed({ samples: 2, units: 200_000 });
    expect(calibration.samplesMs).toHaveLength(2);
    expect(calibration.medianMs).toBeGreaterThan(0);
    expect(calibration.bestMs).toBeGreaterThan(0);
    expect(calibration.unitsPerMs).toBeGreaterThan(0);
    expect(calibration.contentionRatio).toBeGreaterThanOrEqual(1);
  });

  it('exposes the 20% per-change regression ratio', () => {
    expect(MEDIAN_REGRESSION_FACTOR).toBe(1.2);
  });
});
