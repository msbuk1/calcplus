#!/usr/bin/env node
/**
 * Build the calcplus launcher artifacts.
 *
 *   node scripts/build-launcher.mjs [--windows] [--linux] [--require-go]
 *
 * Linux x64 produces `release/calcplus-<version>-linux-x64.tar.gz`; Windows x64
 * produces `release/calcplus-<version>-windows-x64/calcplus.exe`. Both are
 * portable, unsigned and fully offline (distribution.md §2 and §8). The
 * embedded copy of `dist/` is staged under `launcher/webapp/dist` before the
 * Go build; the staged copy and `release/` are ignored by git.
 *
 * Go is an optional tool on the maintainer's machines, like `soffice` and
 * `openpyxl` in the fixture generators: when it is missing the script reports
 * a skip and exits successfully. Pass `--require-go` on a release run to make
 * a missing toolchain fail instead. Set `GO_BIN` to choose a specific Go
 * binary when PATH points at one that cannot build the artifacts (for example
 * gccgo, which the script refuses).
 */
import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import {
  cpSync,
  existsSync,
  mkdirSync,
  readFileSync,
  readdirSync,
  rmSync,
  statSync,
  writeFileSync,
} from 'node:fs';
import { dirname, join } from 'node:path';
import { homedir } from 'node:os';
import { fileURLToPath } from 'node:url';

const root = dirname(dirname(fileURLToPath(import.meta.url)));
const launcherDir = join(root, 'launcher');
const stageDir = join(launcherDir, 'webapp', 'dist');
const appDist = join(root, 'dist');
const releaseDir = join(root, 'release');
const notices = ['LICENSE', 'NOTICE', 'THIRD_PARTY_NOTICES.md'];

const args = process.argv.slice(2);
const knownFlags = new Set(['--windows', '--linux', '--require-go', '--no-build']);
for (const argument of args) {
  if (!knownFlags.has(argument)) {
    console.error(`build-launcher: unknown option ${argument}`);
    console.error(
      '  Usage: node scripts/build-launcher.mjs [--windows] [--linux] [--require-go] [--no-build]',
    );
    process.exit(1);
  }
}
const requireGo = args.includes('--require-go');
const requested = args.filter((argument) => argument === '--windows' || argument === '--linux');
const targets = requested.length === 0 ? ['windows', 'linux'] : requested;
const version = JSON.parse(readFileSync(join(root, 'package.json'), 'utf8')).version;

// GO_BIN picks a toolchain; otherwise the script looks for a standard Go in
// the usual places. The `go` on PATH can be gccgo (as on Pop!_OS), which is
// refused below because it cannot cross-compile the Windows .exe and does not
// link the static Linux binary the release promises.
const goCandidates =
  process.env.GO_BIN !== undefined
    ? [process.env.GO_BIN]
    : [
        'go',
        join('/usr/local/go/bin', 'go'),
        join('/usr/lib/go/bin', 'go'),
        join(homedir(), 'go', 'bin', 'go'),
      ];
let goCommand = null;
let goVersion = '';
let sawGccgo = false;
for (const candidate of goCandidates) {
  const probe = spawnSync(candidate, ['version'], { encoding: 'utf8' });
  if (probe.error !== undefined || probe.status !== 0) continue;
  if (probe.stdout.includes('gccgo')) {
    sawGccgo = true;
    continue;
  }
  goCommand = candidate;
  goVersion = probe.stdout.trim();
  break;
}
if (goCommand === null) {
  if (sawGccgo) {
    console.error('build-launcher: the only Go toolchains found are gccgo, which cannot build the');
    console.error(
      '  artifacts. gccgo does not cross-compile the Windows .exe and links a dynamic,',
    );
    console.error(
      '  not static, Linux binary. Install the standard Go toolchain from https://go.dev/dl/',
    );
    console.error('  (or set GO_BIN to it) and run this script again.');
    process.exit(1);
  }
  console.log('build-launcher: skipped — Go is not installed.');
  console.log('  Install Go 1.18 or newer from https://go.dev/dl/ and run this script again.');
  console.log('  The browser app still builds with `npm run build`; only the launcher needs Go.');
  process.exit(requireGo ? 1 : 0);
}
console.log(`Using ${goVersion} (${goCommand})`);

// The app is rebuilt first so the embedded copy is never stale: the release
// once shipped a bundle from before the ribbon icon work because dist/ was
// left over from an earlier build. --no-build keeps the old behaviour.
if (!args.includes('--no-build')) {
  console.log('Building the app (npm run build) so the embedded copy is current');
  const build = spawnSync('npm', ['run', 'build'], { cwd: root, stdio: 'inherit', shell: false });
  if (build.status !== 0) {
    console.error('build-launcher: npm run build failed; nothing was packaged.');
    process.exit(1);
  }
}
if (!existsSync(join(appDist, 'index.html'))) {
  console.error('build-launcher: dist/index.html is missing. Run `npm run build` first.');
  process.exit(1);
}

stageApp();
divider();
const artifacts = targets.map((target) => (target === 'windows' ? buildWindows() : buildLinux()));
writeReleaseChecksums(artifacts);
divider();
console.log(`Launcher artifacts are in ${releaseDir}`);

function stageApp() {
  console.log(`Staging dist/ for the embedded build (version ${version})`);
  rmSync(stageDir, { recursive: true, force: true });
  cpSync(appDist, stageDir, { recursive: true });
  for (const notice of notices) copyNotice(notice, stageDir);
}

function buildWindows() {
  const directory = join(releaseDir, `calcplus-${version}-windows-x64`);
  rmSync(directory, { recursive: true, force: true });
  mkdirSync(directory, { recursive: true });
  const binary = join(directory, 'calcplus.exe');
  goBuild('windows', 'amd64', binary);
  copyPackageFiles(directory);
  writeChecksums(directory);
  report(binary);
  return { path: binary, name: `calcplus-${version}-windows-x64/calcplus.exe` };
}

function buildLinux() {
  const name = `calcplus-${version}-linux-x64`;
  const directory = join(releaseDir, name);
  rmSync(directory, { recursive: true, force: true });
  mkdirSync(directory, { recursive: true });
  goBuild('linux', 'amd64', join(directory, 'calcplus'));
  // The tarball carries the app twice on purpose: the launcher embeds it, and
  // a plain dist/ copy lets a user read or replace the build without rebuilding
  // the launcher (licensing.md §3.4).
  cpSync(appDist, join(directory, 'dist'), { recursive: true });
  copyPackageFiles(directory);
  writeChecksums(directory);

  const tarball = join(releaseDir, `${name}.tar.gz`);
  rmSync(tarball, { force: true });
  const tar = spawnSync('tar', ['-czf', tarball, '-C', releaseDir, name], {
    stdio: 'inherit',
  });
  if (tar.error !== undefined || tar.status !== 0) {
    console.log('  tar is not available; the unarchived directory is ready to archive by hand.');
    return { path: directory, name };
  }
  report(tarball);
  return { path: tarball, name: `${name}.tar.gz` };
}

function goBuild(goos, goarch, output) {
  console.log(`Building ${goos}/${goarch} launcher`);
  // The build id is the embedded app's index.html hash: two builds of the
  // same version stay distinguishable, so a stale running instance is never
  // silently reused.
  const buildId = createHash('sha256')
    .update(readFileSync(join(stageDir, 'index.html')))
    .digest('hex')
    .slice(0, 12);
  const result = spawnSync(
    goCommand,
    [
      'build',
      '-trimpath',
      '-buildvcs=false',
      '-ldflags',
      `-s -w -X main.version=${version} -X main.buildID=${buildId}`,
      '-o',
      output,
      '.',
    ],
    {
      cwd: launcherDir,
      env: { ...process.env, GOOS: goos, GOARCH: goarch, CGO_ENABLED: '0' },
      stdio: 'inherit',
    },
  );
  if (result.error !== undefined || result.status !== 0) {
    console.error(`build-launcher: the ${goos}/${goarch} build failed.`);
    process.exit(1);
  }
}

/** Copy a root-level notice, or the launcher's own README, into a package. */
function copyNotice(name, directory) {
  const source = name === 'README.md' ? join(launcherDir, name) : join(root, name);
  cpSync(source, join(directory, name));
}

/** The licence and readme files every shipped package carries (licensing.md §6). */
function copyPackageFiles(directory) {
  for (const notice of notices) copyNotice(notice, directory);
  copyNotice('README.md', directory);
  copyGoLicence(directory);
}

/** The Go standard library's BSD-3-Clause notice travels with the binary. */
function copyGoLicence(directory) {
  const goroot = spawnSync(goCommand, ['env', 'GOROOT'], { encoding: 'utf8' });
  const gorootDir = goroot.status === 0 ? goroot.stdout.trim() : '';
  const licence = join(gorootDir, 'LICENSE');
  if (gorootDir === '' || !existsSync(licence)) {
    console.log('  note: the Go licence text was not found; add it to the package before release.');
    return;
  }
  cpSync(licence, join(directory, 'GO-LICENSE'));
}

/** SHA256SUMS over every file in a package directory except the sums file. */
function writeChecksums(directory) {
  const files = readdirSync(directory, { withFileTypes: true })
    .filter((entry) => entry.isFile() && entry.name !== 'SHA256SUMS')
    .map((entry) => entry.name)
    .sort();
  const lines = files.map(
    (name) =>
      `${createHash('sha256')
        .update(readFileSync(join(directory, name)))
        .digest('hex')}  ${name}`,
  );
  writeFileSync(join(directory, 'SHA256SUMS'), `${lines.join('\n')}\n`);
}

/** One SHA256SUMS beside the artifacts, so a downloader checks one file. */
function writeReleaseChecksums(artifacts) {
  const lines = artifacts
    .filter((artifact) => statSync(artifact.path).isFile())
    .map(
      (artifact) =>
        `${createHash('sha256').update(readFileSync(artifact.path)).digest('hex')}  ${artifact.name}`,
    );
  if (lines.length === 0) return;
  writeFileSync(join(releaseDir, 'SHA256SUMS'), `${lines.join('\n')}\n`);
  console.log(`Wrote ${join(releaseDir, 'SHA256SUMS')}`);
}

function report(path) {
  const size = readFileSync(path).byteLength;
  console.log(`Wrote ${path} (${formatSize(size)})`);
}

function formatSize(bytes) {
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(0)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function divider() {
  console.log('');
}
