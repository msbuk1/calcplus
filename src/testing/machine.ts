import { execFileSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { cpus, release, totalmem } from 'node:os';

/**
 * Machine facts and the reference-machine gate (performance §1, §6, §7).
 *
 * The absolute budgets in performance §3 are a contract with the reference
 * machine: a mid-range 2023 laptop with four physical cores and 16 GB of RAM,
 * running Windows 11 or Ubuntu 22.04 or newer. Performance §6 is explicit that
 * a result from a different machine is advisory, not gating, so a focused probe
 * enforces the absolute budget only when the maintainer declares the run with
 * `CALCPLUS_REFERENCE_MACHINE=1` and the hardware matches. Every other run
 * records the fingerprint and the measured machine speed and uses the budget as
 * a liveness ceiling instead.
 *
 * Performance §7 gives the per-change rule for other machines: block on a 20%
 * regression against the rolling median of the last ten local runs on the same
 * machine. The benchmark runner (`bench/model`) owns that comparison; a unit
 * test must never fail because a sibling process was using the CPU.
 */

export const REFERENCE_MIN_PHYSICAL_CORES = 4;
/** 16 GB of RAM is about 14.9 GiB. */
export const REFERENCE_MIN_MEMORY_BYTES = 15 * 1024 ** 3;

/**
 * The liveness ceiling for a run that is not on the reference machine. It is
 * not a budget, only a detector for an engine or grid that stopped working:
 * ten times the contractual number is far outside machine-speed variation (the
 * floor tier is allowed to be 2x slower, and a loaded laptop can be 3-4x) but
 * still catches a hang. Performance §6 keeps non-reference results advisory.
 */
export const ADVISORY_LIVENESS_FACTOR = 10;

/**
 * Performance §7: a per-change run blocks on a 20% regression against the
 * rolling median of the last ten local runs on the same machine.
 */
export const MEDIAN_REGRESSION_FACTOR = 1.2;

export interface MachineFingerprint {
  readonly cpuModel: string;
  readonly logicalCpus: number;
  /** Physical cores, or `null` when they could not be verified. */
  readonly physicalCores: number | null;
  readonly totalMemoryBytes: number;
  readonly platform: string;
  readonly osName: string;
  readonly osVersion: string;
  readonly nodeVersion: string;
}

export function machineFingerprint(): MachineFingerprint {
  const cpu = cpus()[0];
  return {
    cpuModel: cpu?.model.trim() ?? 'unknown',
    logicalCpus: cpus().length,
    physicalCores: physicalCoreCount(),
    totalMemoryBytes: totalmem(),
    platform: process.platform,
    osName: operatingSystemName(),
    osVersion: operatingSystemVersion(),
    nodeVersion: process.version,
  };
}

export function describeFingerprint(fingerprint: MachineFingerprint): string {
  const gib = (fingerprint.totalMemoryBytes / 1024 ** 3).toFixed(1);
  const cores =
    fingerprint.physicalCores === null
      ? `${fingerprint.logicalCpus} logical CPUs`
      : `${fingerprint.physicalCores} physical / ${fingerprint.logicalCpus} logical CPUs`;
  return `${fingerprint.cpuModel}, ${cores}, ${gib} GiB RAM, ${fingerprint.osName} ${fingerprint.osVersion}, Node ${fingerprint.nodeVersion}`;
}

export interface ReferenceStatus {
  /** True when `CALCPLUS_REFERENCE_MACHINE=1` was set for this run. */
  readonly declared: boolean;
  /** True when the run may claim the absolute budget. */
  readonly matches: boolean;
  /** Every failed check, for a report that must say what is unproven. */
  readonly reasons: readonly string[];
}

/** Performance §1: the machine class the absolute budgets apply to. */
export function referenceStatus(
  fingerprint: MachineFingerprint,
  declared: boolean = referenceDeclared(),
): ReferenceStatus {
  const reasons: string[] = [];
  if (!declared) {
    reasons.push(
      'the run was not declared with CALCPLUS_REFERENCE_MACHINE=1 (performance §6 makes other runs advisory)',
    );
  }
  if (fingerprint.physicalCores === null) {
    reasons.push('the physical core count could not be verified');
  } else if (fingerprint.physicalCores < REFERENCE_MIN_PHYSICAL_CORES) {
    reasons.push(
      `only ${fingerprint.physicalCores} physical cores, not ${REFERENCE_MIN_PHYSICAL_CORES}`,
    );
  }
  if (fingerprint.totalMemoryBytes < REFERENCE_MIN_MEMORY_BYTES) {
    reasons.push(
      `only ${Math.round(fingerprint.totalMemoryBytes / 1024 ** 3)} GiB of RAM, not 16 GB`,
    );
  }
  if (!isWindows11OrNewer(fingerprint) && !isUbuntu2204OrNewer(fingerprint)) {
    reasons.push(
      `${fingerprint.osName} ${fingerprint.osVersion} is not Windows 11 or Ubuntu 22.04 or newer`,
    );
  }
  return { declared, matches: reasons.length === 0, reasons };
}

export function referenceDeclared(): boolean {
  return process.env.CALCPLUS_REFERENCE_MACHINE === '1';
}

// ---------------------------------------------------------------------------
// Machine-speed sample.
// ---------------------------------------------------------------------------

/**
 * A fixed arithmetic workload used to report how fast the machine running the
 * probe actually is. It is deliberately integer-heavy and allocation-free so
 * that its throughput tracks the single-thread speed the engine build depends
 * on. It is reported with every result; it does not change a budget by itself.
 */
export const CALIBRATION_UNITS = 20_000_000;

/** Run the fixed workload once and return a checksum that keeps it observable. */
export function calibrationWorkload(units: number = CALIBRATION_UNITS): number {
  let accumulator = 0x9e3779b9;
  let sink = 0;
  for (let index = 0; index < units; index += 1) {
    accumulator ^= index + (accumulator << 6) + (accumulator >>> 2);
    accumulator = Math.imul(accumulator, 2_654_435_761) >>> 0;
    if ((accumulator & 0x3ff) === 0) sink = (sink + Math.sqrt(accumulator)) % 1_000_000;
  }
  return (accumulator + (sink >= 0 ? 0 : 1)) >>> 0;
}

export interface MachineCalibration {
  readonly units: number;
  readonly samplesMs: readonly number[];
  readonly medianMs: number;
  readonly bestMs: number;
  /** Median throughput: the speed number reported alongside a result. */
  readonly unitsPerMs: number;
  readonly bestUnitsPerMs: number;
  /**
   * Median divided by the best sample. A value well above 1 means the process
   * was competing for the CPU during the run.
   */
  readonly contentionRatio: number;
}

export function calibrateMachineSpeed(
  options: { samples?: number; units?: number } = {},
): MachineCalibration {
  const units = Math.max(1, options.units ?? CALIBRATION_UNITS);
  const samples = Math.max(1, options.samples ?? 3);
  const samplesMs: number[] = [];
  for (let sample = 0; sample < samples; sample += 1) {
    const started = now();
    calibrationWorkload(units);
    samplesMs.push(now() - started);
  }
  const sorted = [...samplesMs].sort((left, right) => left - right);
  const medianMs = sorted[Math.floor(sorted.length / 2)] ?? 0;
  const bestMs = sorted[0] ?? 0;
  return {
    units,
    samplesMs,
    medianMs,
    bestMs,
    unitsPerMs: medianMs === 0 ? 0 : units / medianMs,
    bestUnitsPerMs: bestMs === 0 ? 0 : units / bestMs,
    contentionRatio: bestMs === 0 ? 1 : medianMs / bestMs,
  };
}

let cachedCalibration: MachineCalibration | null = null;

/** The process-wide calibration sample, measured once and reused. */
export function machineCalibration(): MachineCalibration {
  cachedCalibration ??= calibrateMachineSpeed();
  return cachedCalibration;
}

/** Test seam: a probe can force a fresh sample instead of the cached one. */
export function resetMachineCalibration(): void {
  cachedCalibration = null;
}

export function describeCalibration(calibration: MachineCalibration): string {
  return `${Math.round(calibration.unitsPerMs).toLocaleString('en-US')} units/ms (best ${Math.round(
    calibration.bestUnitsPerMs,
  ).toLocaleString('en-US')}, contention ${calibration.contentionRatio.toFixed(2)}×)`;
}

// ---------------------------------------------------------------------------
// The budget gate.
// ---------------------------------------------------------------------------

export type BudgetMode = 'enforced' | 'advisory';

export interface BudgetGateOptions {
  readonly case: string;
  readonly budgetMs: number;
  readonly fingerprint?: MachineFingerprint;
  readonly calibration?: MachineCalibration;
  readonly declaredReference?: boolean;
}

export interface BudgetGate {
  readonly case: string;
  readonly budgetMs: number;
  readonly measuredMs: number;
  readonly runsMs: readonly number[];
  readonly withinBudget: boolean;
  /** `enforced` only on a declared, matching reference machine. */
  readonly mode: BudgetMode;
  /** False only when the contract actually gates and the budget is missed. */
  readonly passed: boolean;
  readonly livenessCeilingMs: number;
  /** True when the contract's "one immediate re-run" filtered the first miss. */
  readonly retried: boolean;
  readonly reference: ReferenceStatus;
  readonly fingerprint: MachineFingerprint;
  readonly calibration: MachineCalibration;
  readonly reason: string;
}

/**
 * Run a probe and decide what its numbers mean.
 *
 * On the reference machine the median must be at or under the budget; a miss
 * gets the one immediate re-run performance §7 allows. Anywhere else the
 * fingerprint and machine speed are reported and the budget becomes the
 * liveness ceiling: the run still catches a stopped engine, but it does not
 * make a performance claim the machine cannot support.
 */
export function runBudgetProbe(
  options: BudgetGateOptions,
  measure: () => readonly number[],
): BudgetGate {
  const fingerprint = options.fingerprint ?? machineFingerprint();
  const calibration = options.calibration ?? machineCalibration();
  const reference = referenceStatus(fingerprint, options.declaredReference);

  let runsMs = measure();
  let retried = false;
  if (reference.matches && median(runsMs) > options.budgetMs) {
    // Performance §7: one immediate re-run filters noise; a second miss blocks.
    runsMs = measure();
    retried = true;
  }

  const measuredMs = median(runsMs);
  const withinBudget = measuredMs <= options.budgetMs;
  const mode: BudgetMode = reference.matches ? 'enforced' : 'advisory';
  const livenessCeilingMs = options.budgetMs * ADVISORY_LIVENESS_FACTOR;
  const passed = mode === 'enforced' ? withinBudget : measuredMs <= livenessCeilingMs;

  return {
    case: options.case,
    budgetMs: options.budgetMs,
    measuredMs,
    runsMs,
    withinBudget,
    mode,
    passed,
    livenessCeilingMs,
    retried,
    reference,
    fingerprint,
    calibration,
    reason: gateReason(options.case, options.budgetMs, measuredMs, mode, passed, reference),
  };
}

function gateReason(
  name: string,
  budgetMs: number,
  measuredMs: number,
  mode: BudgetMode,
  passed: boolean,
  reference: ReferenceStatus,
): string {
  const measured = `${measuredMs.toFixed(1)} ms against ${budgetMs} ms`;
  if (mode === 'enforced') {
    return passed
      ? `${name}: ${measured} on the declared reference machine`
      : `${name}: ${measured} on the declared reference machine, after one immediate re-run`;
  }
  const advisory = `${name}: ${measured} is advisory on this machine (${reference.reasons.join('; ')})`;
  return passed
    ? `${advisory}; it is inside the liveness ceiling, but performance §6 makes the absolute budget non-gating here`
    : `${advisory}; it also exceeds the liveness ceiling, so the probe did not finish sanely`;
}

export function describeGate(gate: BudgetGate): string {
  return `${gate.reason}. Machine: ${describeFingerprint(gate.fingerprint)}. Speed: ${describeCalibration(
    gate.calibration,
  )}.${gate.retried ? ' The first measurement missed and was re-run once (performance §7).' : ''}`;
}

// ---------------------------------------------------------------------------
// Small helpers.
// ---------------------------------------------------------------------------

export function median(values: readonly number[]): number {
  if (values.length === 0) return 0;
  const sorted = [...values].sort((left, right) => left - right);
  return sorted[Math.floor(sorted.length / 2)] ?? 0;
}

function now(): number {
  return globalThis.performance.now();
}

function isWindows11OrNewer(fingerprint: MachineFingerprint): boolean {
  if (fingerprint.platform !== 'win32') return false;
  const build = Number(fingerprint.osVersion.split('.')[2] ?? '0');
  return Number.isFinite(build) && build >= 22_000;
}

function isUbuntu2204OrNewer(fingerprint: MachineFingerprint): boolean {
  if (fingerprint.platform !== 'linux') return false;
  if (fingerprint.osName !== 'Ubuntu') return false;
  const [major = 0, minor = 0] = fingerprint.osVersion.split('.').map((part) => Number(part));
  return major > 22 || (major === 22 && minor >= 4);
}

function operatingSystemName(): string {
  if (process.platform === 'win32') {
    const build = Number(release().split('.')[2] ?? '0');
    return build >= 22_000 ? 'Windows 11' : 'Windows 10';
  }
  if (process.platform === 'linux') {
    const releaseValues = readOsRelease();
    const name = releaseValues?.get('NAME');
    if (name !== undefined) return name.replace(/"/g, '');
    return 'Linux';
  }
  return process.platform;
}

function operatingSystemVersion(): string {
  if (process.platform === 'win32') return release();
  if (process.platform === 'linux') {
    const version = readOsRelease()?.get('VERSION_ID');
    if (version !== undefined) return version.replace(/"/g, '');
  }
  return release();
}

function readOsRelease(): Map<string, string> | null {
  try {
    const text = readFileSync('/etc/os-release', 'utf8');
    const entries = new Map<string, string>();
    for (const line of text.split('\n')) {
      const separator = line.indexOf('=');
      if (separator < 0) continue;
      entries.set(line.slice(0, separator), line.slice(separator + 1).replace(/"/g, ''));
    }
    return entries;
  } catch {
    return null;
  }
}

/** Physical cores, not logical CPUs: a four-core machine with SMT reports eight. */
export function physicalCoreCount(): number | null {
  if (process.platform === 'linux') return linuxPhysicalCores();
  if (process.platform === 'win32') return windowsPhysicalCores();
  return null;
}

function windowsPhysicalCores(): number | null {
  try {
    const output = execFileSync(
      'powershell.exe',
      [
        '-NoProfile',
        '-NonInteractive',
        '-Command',
        '(Get-CimInstance Win32_Processor | Measure-Object -Property NumberOfCores -Sum).Sum',
      ],
      { encoding: 'utf8', timeout: 30_000, stdio: ['ignore', 'pipe', 'ignore'] },
    );
    const cores = Number(output.trim());
    return Number.isFinite(cores) && cores > 0 ? cores : null;
  } catch {
    return null;
  }
}

function linuxPhysicalCores(): number | null {
  try {
    const text = readFileSync('/proc/cpuinfo', 'utf8');
    const cores = new Set<string>();
    let physicalId = '';
    let coreId = '';
    for (const line of text.split('\n')) {
      const separator = line.indexOf(':');
      if (separator < 0) {
        if (line.trim() === '' && physicalId !== '' && coreId !== '') {
          cores.add(`${physicalId}/${coreId}`);
        }
        physicalId = '';
        coreId = '';
        continue;
      }
      const key = line.slice(0, separator).trim();
      const value = line.slice(separator + 1).trim();
      if (key === 'physical id') physicalId = value;
      else if (key === 'core id') coreId = value;
    }
    if (physicalId !== '' && coreId !== '') cores.add(`${physicalId}/${coreId}`);
    return cores.size > 0 ? cores.size : null;
  } catch {
    return null;
  }
}
