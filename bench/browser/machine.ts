import { execFileSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { cpus, release, totalmem } from 'node:os';

/**
 * The machine facts the performance contract needs (performance §1 and §6).
 *
 * The reference class is a mid-range 2023 laptop: a four-core CPU with 16 GB
 * of RAM, running Windows 11 or Ubuntu 22.04 or newer, with the latest two
 * stable Chromium versions, at 1440 x 900 CSS pixels and one workbook per tab.
 */
export interface MachineFacts {
  readonly platform: string;
  readonly osName: string;
  readonly osVersion: string;
  readonly cpuModel: string;
  /** Logical CPUs the operating system reports. */
  readonly logicalCpus: number;
  /** Physical CPU cores, or `null` when they could not be verified. */
  readonly physicalCores: number | null;
  readonly totalMemoryBytes: number;
  readonly browserVersion: string;
  readonly viewport: { readonly width: number; readonly height: number };
  readonly devicePixelRatio: number;
}

export interface ReferenceCheck {
  /** True when the machine satisfies the reference class of performance §1. */
  readonly matches: boolean;
  /** Every failed check, so a report can say what was missing. */
  readonly reasons: readonly string[];
}

const REFERENCE_VIEWPORT = { width: 1440, height: 900 };
/** 16 GB of RAM is about 15.7 GiB. */
const MINIMUM_RAM_BYTES = 15 * 1024 ** 3;
const MINIMUM_PHYSICAL_CORES = 4;

export function describeMachine(
  browserVersion: string,
  viewport: { width: number; height: number },
  devicePixelRatio: number,
): MachineFacts {
  const cpu = cpus()[0];
  return {
    platform: process.platform,
    osName: operatingSystemName(),
    osVersion: operatingSystemVersion(),
    cpuModel: cpu?.model ?? 'unknown',
    logicalCpus: cpus().length,
    physicalCores: physicalCoreCount(),
    totalMemoryBytes: totalmem(),
    browserVersion,
    viewport,
    devicePixelRatio,
  };
}

/**
 * A reference claim needs the machine to satisfy the contract and the
 * maintainer to declare the run as a reference run (performance §6). A run
 * that is not on the reference machine stays advisory.
 */
export function checkReferenceClass(
  facts: MachineFacts,
  declaredReferenceMachine: boolean,
): ReferenceCheck {
  const reasons: string[] = [];
  if (!declaredReferenceMachine) {
    reasons.push('the run was not declared as a reference-machine run');
  }
  if (facts.physicalCores === null) {
    reasons.push('the physical core count could not be verified');
  } else if (facts.physicalCores < MINIMUM_PHYSICAL_CORES) {
    reasons.push(`only ${facts.physicalCores} physical cores`);
  }
  if (facts.totalMemoryBytes < MINIMUM_RAM_BYTES) {
    reasons.push(`only ${Math.round(facts.totalMemoryBytes / 1024 ** 3)} GiB of RAM`);
  }
  if (!isWindows11OrNewer(facts) && !isUbuntu2204OrNewer(facts)) {
    reasons.push(`${facts.osName} ${facts.osVersion} is not Windows 11 or Ubuntu 22.04 or newer`);
  }
  if (
    facts.viewport.width !== REFERENCE_VIEWPORT.width ||
    facts.viewport.height !== REFERENCE_VIEWPORT.height
  ) {
    reasons.push(
      `the viewport is ${facts.viewport.width} x ${facts.viewport.height}, not ${REFERENCE_VIEWPORT.width} x ${REFERENCE_VIEWPORT.height}`,
    );
  }
  if (facts.devicePixelRatio !== 1) {
    reasons.push(`the device pixel ratio is ${facts.devicePixelRatio}, not 1`);
  }
  return { matches: reasons.length === 0, reasons };
}

function isWindows11OrNewer(facts: MachineFacts): boolean {
  if (facts.platform !== 'win32') return false;
  // Windows 11 reports kernel build 22000 or newer; Windows 10 stops at 19045.
  const build = Number(facts.osVersion.split('.')[2] ?? '0');
  return Number.isFinite(build) && build >= 22_000;
}

function isUbuntu2204OrNewer(facts: MachineFacts): boolean {
  if (facts.platform !== 'linux') return false;
  if (facts.osName !== 'Ubuntu') return false;
  const parts = facts.osVersion.split('.').map((part) => Number(part));
  const major = parts[0] ?? 0;
  const minor = parts[1] ?? 0;
  return major > 22 || (major === 22 && minor >= 4);
}

function operatingSystemName(): string {
  if (process.platform === 'win32') {
    const build = Number(release().split('.')[2] ?? '0');
    return build >= 22_000 ? 'Windows 11' : 'Windows 10';
  }
  if (process.platform === 'linux') {
    const release = readOsRelease();
    if (release !== null) {
      const name = release.get('NAME');
      if (name !== undefined) return name.replace(/"/g, '');
    }
    return 'Linux';
  }
  return process.platform;
}

function operatingSystemVersion(): string {
  if (process.platform === 'win32') return release();
  if (process.platform === 'linux') {
    const release = readOsRelease();
    const version = release?.get('VERSION_ID');
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

/**
 * The physical core count, not the logical CPU count: a four-core machine
 * with hyper-threading reports eight logical CPUs.
 */
export function physicalCoreCount(): number | null {
  if (process.platform === 'linux') return linuxPhysicalCores();
  if (process.platform === 'win32') return windowsPhysicalCores();
  return null;
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
