import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import type { MachineFingerprint } from '@testing/machine';
import { MEDIAN_REGRESSION_FACTOR } from '@testing/machine';

/**
 * Rolling local benchmark history (performance §7).
 *
 * The per-change gate blocks on a 20% regression against the rolling median of
 * the last ten local runs on the same machine. `bench/results/` is gitignored:
 * the history is local, like the numbers. A run whose calibration shows the
 * machine was competing for the CPU is reported but does not enter the history,
 * so a busy laptop cannot poison the baseline.
 */
export const REGRESSION_BLOCK_RATIO = MEDIAN_REGRESSION_FACTOR;
export const HISTORY_LENGTH = 10;
/**
 * A 20% change below this median is measurement noise, not a regression. Cases
 * faster than one frame (16.7 ms) or a typical command (50 ms) are reported
 * but never block.
 */
export const REGRESSION_MIN_MEDIAN_MS = 50;

const RESULTS_DIR = fileURLToPath(new URL('../results/model', import.meta.url));

export interface CaseHistory {
  /** The last ten measured medians, oldest first. */
  readonly mediansMs: readonly number[];
  readonly updatedAt: string;
}

export interface RegressionCheck {
  readonly case: string;
  readonly currentMs: number;
  readonly rollingMedianMs: number | null;
  readonly ratio: number | null;
  readonly blocked: boolean;
  readonly reason: string;
}

interface HistoryFile {
  readonly machine: string;
  readonly cases: Record<string, CaseHistory>;
}

export function machineKey(fingerprint: MachineFingerprint): string {
  const raw = [
    fingerprint.platform,
    fingerprint.osName,
    fingerprint.osVersion,
    fingerprint.cpuModel,
    fingerprint.physicalCores ?? 'unknown',
    fingerprint.logicalCpus,
    fingerprint.totalMemoryBytes,
  ].join('|');
  let hash = 5381;
  for (let index = 0; index < raw.length; index += 1) {
    hash = ((hash * 33) ^ raw.charCodeAt(index)) >>> 0;
  }
  return hash.toString(16);
}

export function historyPath(fingerprint: MachineFingerprint): string {
  return join(RESULTS_DIR, `history-${machineKey(fingerprint)}.json`);
}

export function readHistory(fingerprint: MachineFingerprint): Record<string, CaseHistory> {
  const path = historyPath(fingerprint);
  if (!existsSync(path)) return {};
  try {
    const parsed = JSON.parse(readFileSync(path, 'utf8')) as HistoryFile;
    return parsed.cases;
  } catch {
    return {};
  }
}

export function recordMedian(
  fingerprint: MachineFingerprint,
  caseName: string,
  medianMs: number,
  options: { readonly contended?: boolean; readonly now?: Date } = {},
): RegressionCheck {
  const cases = { ...readHistory(fingerprint) };
  const previous = cases[caseName];
  const rolling =
    previous === undefined || previous.mediansMs.length === 0 ? null : medianOf(previous.mediansMs);
  const ratio = rolling === null || rolling === 0 ? null : medianMs / rolling;
  const comparable = rolling !== null && rolling >= REGRESSION_MIN_MEDIAN_MS;
  const blocked = comparable && ratio !== null && ratio > REGRESSION_BLOCK_RATIO;

  if (options.contended !== true) {
    const mediansMs = [...(previous?.mediansMs ?? []), medianMs].slice(-HISTORY_LENGTH);
    cases[caseName] = {
      mediansMs,
      updatedAt: (options.now ?? new Date()).toISOString(),
    };
    writeHistory(fingerprint, cases);
  }

  const reason = (() => {
    if (rolling === null) return 'no previous local median yet';
    if (options.contended === true) {
      return `machine was contended, so the result is reported but not compared (rolling median ${rolling.toFixed(1)} ms)`;
    }
    if (!comparable) {
      return `below the ${REGRESSION_MIN_MEDIAN_MS} ms regression floor, so the result is reported but not compared`;
    }
    if (blocked) {
      return `regression: ${medianMs.toFixed(1)} ms is ${(ratio! * 100).toFixed(0)}% of the rolling local median ${rolling.toFixed(1)} ms (block above ${(REGRESSION_BLOCK_RATIO * 100).toFixed(0)}%)`;
    }
    return `within ${(REGRESSION_BLOCK_RATIO * 100).toFixed(0)}% of the rolling local median ${rolling.toFixed(1)} ms`;
  })();

  return { case: caseName, currentMs: medianMs, rollingMedianMs: rolling, ratio, blocked, reason };
}

export function writeReport(name: string, payload: unknown): string {
  mkdirSync(RESULTS_DIR, { recursive: true });
  const path = join(RESULTS_DIR, name);
  mkdirSync(dirname(path), { recursive: true });
  writeFileSync(path, `${JSON.stringify(payload, null, 2)}\n`);
  return path;
}

function writeHistory(fingerprint: MachineFingerprint, cases: Record<string, CaseHistory>): void {
  const payload: HistoryFile = { machine: machineKey(fingerprint), cases };
  writeReport(`history-${machineKey(fingerprint)}.json`, payload);
}

function medianOf(values: readonly number[]): number {
  const sorted = [...values].sort((left, right) => left - right);
  return sorted[Math.floor(sorted.length / 2)] ?? 0;
}
