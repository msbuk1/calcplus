#!/usr/bin/env node
/**
 * Local repair-free verifier (file-compat §9 F22, testing-conformance §5).
 *
 *   node scripts/verify-repair-free.mjs [paths…] [--no-library-checks]
 *
 * With no path it verifies fixtures/compat/producers (the F22 matrix). Pass a
 * file or directory to check anything else. Every `.xlsx` and `.xlsm` is
 * checked structurally (relationships resolve, content types are complete,
 * every XML part is well formed, part names are safe). When the library checks
 * are on, each file is also loaded by openpyxl and converted by LibreOffice
 * headless, when those tools are installed. Missing tools are reported and
 * skipped, never treated as a failure.
 *
 * The deliberately hostile fixtures (corrupt, encrypted, zip-bomb) fail this
 * check by design and are not part of the default target; the open path's
 * Tier C tests own them.
 *
 * The authoritative "opens in Excel 365 without a repair prompt" run stays a
 * manual step on the maintainer's Windows machine (scripts/oracle/README.md).
 */
import { execFileSync, spawnSync } from 'node:child_process';
import { mkdtempSync, readFileSync, readdirSync, rmSync, statSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createServer } from 'vite';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const argumentsList = process.argv.slice(2);
const libraryChecks = !argumentsList.includes('--no-library-checks');
const requested = argumentsList.filter((argument) => !argument.startsWith('--'));
const targets =
  requested.length === 0 ? [join(root, 'fixtures', 'compat', 'producers')] : requested;

function collectPackages(path) {
  const absolute = resolve(path);
  let stats;
  try {
    stats = statSync(absolute);
  } catch (error) {
    console.error(`Cannot read ${path}: ${error.message}`);
    process.exitCode = 1;
    return [];
  }
  if (stats.isFile()) return /\.(xlsx|xlsm)$/i.test(absolute) ? [absolute] : [];
  const files = [];
  for (const entry of readdirSync(absolute).sort()) {
    files.push(...collectPackages(join(absolute, entry)));
  }
  return files;
}

async function loadVerifier() {
  const server = await createServer({
    root,
    configFile: join(root, 'vite.config.ts'),
    appType: 'custom',
    logLevel: 'silent',
    server: { middlewareMode: true },
  });
  try {
    const module = await server.ssrLoadModule('/src/ooxml/verify-package.ts');
    return module.verifyPackage;
  } finally {
    await server.close();
  }
}

function commandExists(command) {
  const probe = spawnSync(command, ['--version'], { stdio: 'ignore' });
  return probe.error === undefined;
}

function pythonHasOpenpyxl() {
  const probe = spawnSync('python3', ['-c', 'import openpyxl'], { stdio: 'ignore' });
  return probe.status === 0;
}

function checkOpenpyxl(file) {
  execFileSync(
    'python3',
    [
      '-c',
      'import sys, openpyxl; openpyxl.load_workbook(sys.argv[1], read_only=True).close()',
      file,
    ],
    { stdio: 'pipe' },
  );
}

function checkLibreOffice(file) {
  const profile = mkdtempSync(join(tmpdir(), 'calcplus-lo-'));
  const output = mkdtempSync(join(tmpdir(), 'calcplus-lo-out-'));
  try {
    const result = spawnSync(
      'soffice',
      [
        `-env:UserInstallation=${pathToFileURL(profile).href}`,
        '--headless',
        '--norestore',
        '--convert-to',
        'xlsx',
        '--outdir',
        output,
        file,
      ],
      { stdio: 'pipe', timeout: 120_000 },
    );
    if (result.status !== 0) {
      const stderr = result.stderr?.toString().trim() ?? '';
      throw new Error(
        `LibreOffice exited with ${result.status}${stderr === '' ? '' : `: ${stderr}`}`,
      );
    }
  } finally {
    rmSync(profile, { recursive: true, force: true });
    rmSync(output, { recursive: true, force: true });
  }
}

async function main() {
  const files = targets.flatMap(collectPackages);
  if (files.length === 0) {
    console.log('verify-repair-free: no .xlsx or .xlsm files to check.');
    return;
  }
  const verifyPackage = await loadVerifier();
  const openpyxl = libraryChecks && pythonHasOpenpyxl();
  const libreOffice = libraryChecks && commandExists('soffice');
  if (libraryChecks) {
    if (!openpyxl)
      console.log('verify-repair-free: openpyxl is not installed; its load check is skipped.');
    if (!libreOffice)
      console.log('verify-repair-free: soffice is not installed; its conversion check is skipped.');
  }

  let failures = 0;
  for (const file of files) {
    const label = file.startsWith(root) ? file.slice(root.length + 1) : file;
    const result = verifyPackage(new Uint8Array(readFileSync(file)), label);
    const notes = [];
    if (result.ok) {
      notes.push(`${result.parts.length} parts`);
      if (openpyxl) {
        try {
          checkOpenpyxl(file);
          notes.push('openpyxl ok');
        } catch (error) {
          result.problems.push({
            code: 'openpyxl-load',
            message: `openpyxl cannot load the file: ${firstLine(error)}`,
          });
        }
      }
      if (libreOffice) {
        try {
          checkLibreOffice(file);
          notes.push('LibreOffice ok');
        } catch (error) {
          result.problems.push({
            code: 'libreoffice-convert',
            message: `LibreOffice cannot open the file: ${firstLine(error)}`,
          });
        }
      }
    }
    if (result.problems.length === 0) {
      console.log(`OK   ${label} (${notes.join(', ')})`);
      continue;
    }
    failures += 1;
    console.log(`FAIL ${label}`);
    for (const problem of result.problems) {
      console.log(`     - [${problem.code}] ${problem.message}`);
    }
  }
  if (failures > 0) {
    console.error(`verify-repair-free: ${failures} of ${files.length} package(s) failed.`);
    process.exitCode = 1;
    return;
  }
  console.log(`verify-repair-free: ${files.length} package(s) are well formed.`);
}

function firstLine(error) {
  const text = error instanceof Error ? error.message : String(error);
  return text.split('\n')[0];
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : error);
  process.exit(1);
});
