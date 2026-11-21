import { spawn, spawnSync, type ChildProcess } from 'node:child_process';
import {
  chmodSync,
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  readdirSync,
  rmSync,
  writeFileSync,
} from 'node:fs';
import { createServer } from 'node:http';
import { connect, createServer as createTcpServer } from 'node:net';
import { networkInterfaces, tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';

/**
 * M7 lane 6 — the launcher and the published documentation skeleton.
 *
 * The contract is distribution.md §2 and §8: a Go static server bound to the
 * literal `127.0.0.1` on the stable port 47823, offline, serving the built
 * `dist/`, plus artifacts that build with or without Go on the machine.
 *
 * The runtime cases build the launcher with the local Go toolchain and speak
 * HTTP to it. They skip when Go is missing; the source and documentation cases
 * always run.
 */

const root = fileURLToPath(new URL('..', import.meta.url));
const launcherDir = join(root, 'launcher');
const goCommand = process.env.GO_BIN ?? 'go';
const go = spawnSync(goCommand, ['version'], { encoding: 'utf8' });
const goAvailable = go.error === undefined && go.status === 0;

const tempDirs: string[] = [];
const children: ChildProcess[] = [];
let binary = '';

function makeTemp(prefix: string): string {
  const directory = mkdtempSync(join(tmpdir(), prefix));
  tempDirs.push(directory);
  return directory;
}

/** A tiny stand-in for `dist/`: an index, a hashed asset, a nested index. */
function makeDist(): string {
  const directory = makeTemp('calcplus-dist-');
  writeFileSync(join(directory, 'index.html'), '<!doctype html><div id="root">calcplus</div>');
  mkdirSync(join(directory, 'assets'));
  writeFileSync(join(directory, 'assets', 'app.js'), 'console.log("calcplus");\n');
  mkdirSync(join(directory, 'nested'));
  writeFileSync(join(directory, 'nested', 'index.html'), '<p>nested</p>');
  return directory;
}

/** An unused loopback port. */
function freePort(): Promise<number> {
  return new Promise((resolvePort, reject) => {
    const server = createTcpServer();
    server.once('error', reject);
    server.listen(0, '127.0.0.1', () => {
      const address = server.address();
      const port = typeof address === 'object' && address !== null ? address.port : 0;
      server.close(() => resolvePort(port));
    });
  });
}

async function waitFor(check: () => boolean, timeoutMs = 10_000): Promise<void> {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    if (check()) return;
    await new Promise((done) => setTimeout(done, 50));
  }
  throw new Error('timed out waiting for the launcher');
}

function startLauncher(args: string[]): { child: ChildProcess; output: () => string } {
  const child = spawn(binary, args, { stdio: ['ignore', 'pipe', 'pipe'] });
  children.push(child);
  let output = '';
  const collect = (chunk: unknown) => {
    output += String(chunk);
  };
  child.stdout?.on('data', collect);
  child.stderr?.on('data', collect);
  return { child, output: () => output };
}

async function stop(child: ChildProcess): Promise<void> {
  if (child.exitCode !== null || child.signalCode !== null) return;
  const exited = new Promise<void>((done) => child.once('exit', () => done()));
  child.kill('SIGKILL');
  await exited;
}

/** True when nothing accepts a TCP connection at host:port. */
function refuses(port: number, host: string): Promise<boolean> {
  return new Promise((done) => {
    const socket = connect({ port, host, timeout: 750 });
    let settled = false;
    const finish = (refused: boolean) => {
      if (settled) return;
      settled = true;
      socket.destroy();
      done(refused);
    };
    socket.once('connect', () => finish(false));
    socket.once('error', () => finish(true));
    socket.once('timeout', () => finish(true));
  });
}

function nonLoopbackAddresses(): string[] {
  const addresses: string[] = [];
  for (const entries of Object.values(networkInterfaces())) {
    for (const entry of entries ?? []) {
      if (entry.family === 'IPv4' && !entry.internal) addresses.push(entry.address);
    }
  }
  return addresses;
}

/** An environment whose PATH holds only `pathValue`, with no GO_BIN override. */
function envWithoutGo(pathValue: string): NodeJS.ProcessEnv {
  const env: NodeJS.ProcessEnv = { ...process.env, PATH: pathValue, Path: pathValue };
  delete env.GO_BIN;
  return env;
}

describe('launcher source', () => {
  const source = readFileSync(join(launcherDir, 'main.go'), 'utf8');

  it('binds the literal 127.0.0.1 and no wildcard address', () => {
    expect(source).toContain('loopbackHost = "127.0.0.1"');
    expect(source).toContain('net.Listen("tcp4"');
    expect(source).not.toContain('0.0.0.0');
    expect(source).not.toContain(':47823'); // the address comes from one constant
  });

  it('keeps the stable default port from the contract', () => {
    expect(source).toContain('defaultPort = 47823');
  });

  it('stamps the launcher version from package.json', () => {
    const packageVersion = (
      JSON.parse(readFileSync(join(root, 'package.json'), 'utf8')) as { version: string }
    ).version;
    expect(source).toContain(`var version = "${packageVersion}"`);
  });

  it('opens an app window in a Chromium browser rather than a tab', () => {
    // ADR 0009: --app gives a window with no tabs, address bar or bookmarks.
    expect(source).toContain('"--app=" + target');
    expect(source).toContain('func appModeCommand(');
    // Chrome first, Edge as the fallback that ships with Windows.
    expect(source.indexOf('Google')).toBeLessThan(source.indexOf('Microsoft'));
  });

  it('keeps the default-browser path as the fallback', () => {
    // With no Chromium browser installed, or app mode off, the old behaviour
    // is unchanged and still reaches the user.
    expect(source).toContain('url.dll,FileProtocolHandler');
    expect(source).toContain('"xdg-open"');
  });

  it('has a --no-app-mode escape that opens an ordinary tab', () => {
    expect(source).toContain('"no-app-mode"');
    expect(source).toContain('openBrowser(appURL, !*noAppMode, out)');
  });

  it("uses the user's own browser profile, never a separate storage area", () => {
    // A dedicated --user-data-dir would hide recovery copies and recent files
    // from the browser the user opens by hand: storage that looks like data
    // loss, the trap distribution.md §2 warns about with --port N.
    expect(source).not.toContain('--user-data-dir');
  });

  it('says which browser opened the window', () => {
    expect(source).toContain('Opened in its own window with');
  });
});

describe('launcher build script', () => {
  it('skips with a clear message when Go is not installed', () => {
    const withoutGo = makeTemp('calcplus-no-go-');
    const result = spawnSync(process.execPath, [join(root, 'scripts', 'build-launcher.mjs')], {
      cwd: root,
      // GO_BIN is an exact choice: pointing it at nothing simulates a machine
      // without Go, wherever the script would otherwise look.
      env: { ...envWithoutGo(withoutGo), GO_BIN: join(withoutGo, 'go') },
      encoding: 'utf8',
    });
    expect(result.status).toBe(0);
    expect(result.stdout).toContain('build-launcher: skipped — Go is not installed.');
    expect(result.stdout).toContain('https://go.dev/dl/');
    expect(result.stdout).toContain('npm run build');
  });

  it('fails on a release run when --require-go is set and Go is missing', () => {
    const withoutGo = makeTemp('calcplus-no-go-required-');
    const result = spawnSync(
      process.execPath,
      [join(root, 'scripts', 'build-launcher.mjs'), '--require-go'],
      {
        cwd: root,
        env: { ...envWithoutGo(withoutGo), GO_BIN: join(withoutGo, 'go') },
        encoding: 'utf8',
      },
    );
    expect(result.status).toBe(1);
    expect(result.stdout).toContain('skipped — Go is not installed.');
  });

  it.skipIf(process.platform === 'win32')(
    'refuses a gccgo toolchain instead of mislabelling the artifacts',
    () => {
      const fake = makeTemp('calcplus-gccgo-');
      const goScript = join(fake, 'go');
      writeFileSync(goScript, '#!/bin/sh\necho "go version go1.18 gccgo (GCC) 14.2.0"\nexit 0\n');
      chmodSync(goScript, 0o755);
      const result = spawnSync(process.execPath, [join(root, 'scripts', 'build-launcher.mjs')], {
        cwd: root,
        env: { ...envWithoutGo(fake), GO_BIN: goScript },
        encoding: 'utf8',
      });
      expect(result.status).toBe(1);
      expect(result.stderr).toContain('gccgo');
      expect(result.stderr).toContain('GO_BIN');
    },
  );
});

describe('published documentation', () => {
  it('has an ADR index that lists every recorded ADR', () => {
    const index = readFileSync(join(root, 'docs', 'adr', 'README.md'), 'utf8');
    const decisions = readdirSync(join(root, 'docs', 'adr')).filter((name) =>
      /^\d{4}-.+\.md$/.test(name),
    );
    expect(decisions.length).toBeGreaterThan(0);
    for (const decision of decisions) {
      expect(index, `${decision} is missing from the ADR index`).toContain(decision);
    }
  });

  it('has a PRD, a technical design and a known-deviations page', () => {
    for (const page of ['README.md', 'prd.md', 'technical-design.md', 'known-deviations.md']) {
      expect(existsSync(join(root, 'docs', page)), `docs/${page} is missing`).toBe(true);
    }
  });

  it('documents the unsigned SmartScreen warning where the contract requires', () => {
    const launcherReadme = readFileSync(join(root, 'launcher', 'README.md'), 'utf8');
    expect(launcherReadme).toContain('Windows protected your PC');
    expect(launcherReadme).toContain('not code-signed');
    expect(launcherReadme).toContain('Run anyway');

    const deviations = readFileSync(join(root, 'docs', 'known-deviations.md'), 'utf8');
    expect(deviations).toContain('SmartScreen');
    expect(deviations).toContain('unsigned');
  });

  it('keeps every local link in the docs skeleton resolvable', () => {
    const pages = [
      'docs/README.md',
      'docs/prd.md',
      'docs/technical-design.md',
      'docs/known-deviations.md',
      'docs/adr/README.md',
    ];
    const failures: string[] = [];
    for (const page of pages) {
      const absolute = join(root, page);
      const text = readFileSync(absolute, 'utf8');
      for (const match of text.matchAll(/\[[^\]]*\]\(([^)\s]+)\)/g)) {
        const target = match[1];
        if (
          target === undefined ||
          target.startsWith('http://') ||
          target.startsWith('https://') ||
          target.startsWith('mailto:') ||
          target.startsWith('#')
        ) {
          continue;
        }
        const withoutAnchor = target.split('#')[0];
        if (withoutAnchor === undefined || withoutAnchor === '') continue;
        const resolved = resolve(dirname(absolute), withoutAnchor);
        if (!existsSync(resolved)) failures.push(`${page}: ${target}`);
      }
    }
    expect(failures).toEqual([]);
  });
});

describe.skipIf(!goAvailable)('launcher runtime', () => {
  beforeAll(() => {
    const output = makeTemp('calcplus-bin-');
    binary = join(output, process.platform === 'win32' ? 'calcplus.exe' : 'calcplus');
    const build = spawnSync(
      goCommand,
      ['build', '-trimpath', '-buildvcs=false', '-o', binary, '.'],
      {
        cwd: launcherDir,
        encoding: 'utf8',
        timeout: 120_000,
      },
    );
    if (build.status !== 0) {
      throw new Error(`go build failed: ${build.stderr}`);
    }
  }, 120_000);

  afterAll(async () => {
    for (const child of children) await stop(child);
    for (const directory of tempDirs) rmSync(directory, { recursive: true, force: true });
  });

  it('serves dist/ and keeps 404s for missing files and directories', async () => {
    const dist = makeDist();
    const port = await freePort();
    const launcher = startLauncher(['--dist', dist, '--port', String(port), '--no-open']);
    await waitFor(() => launcher.output().includes('Serving the local build'));

    try {
      const index = await fetch(`http://127.0.0.1:${port}/`);
      expect(index.status).toBe(200);
      expect(await index.text()).toContain('id="root"');
      expect(index.headers.get('x-content-type-options')).toBe('nosniff');
      expect(index.headers.get('cache-control')).toBe('no-cache');

      const asset = await fetch(`http://127.0.0.1:${port}/assets/app.js`);
      expect(asset.status).toBe(200);
      expect(asset.headers.get('content-type')).toMatch(/javascript/);
      expect(await asset.text()).toContain('calcplus');

      const nested = await fetch(`http://127.0.0.1:${port}/nested`);
      expect(nested.status).toBe(200);
      expect(await nested.text()).toContain('nested');

      expect((await fetch(`http://127.0.0.1:${port}/missing.txt`)).status).toBe(404);
      expect((await fetch(`http://127.0.0.1:${port}/assets/`)).status).toBe(404);
      expect((await fetch(`http://127.0.0.1:${port}/assets`)).status).toBe(404);

      const traversal = await fetch(`http://127.0.0.1:${port}/%2e%2e/%2e%2e/package.json`, {
        redirect: 'manual',
      });
      expect([301, 400, 404]).toContain(traversal.status);

      const health = await fetch(`http://127.0.0.1:${port}/__calcplus/health`);
      expect(health.status).toBe(200);
      const payload = (await health.json()) as { app: string; version: string };
      expect(payload.app).toBe('calcplus');
      expect(payload.version).toMatch(/^\d+\.\d+\.\d+$/);
    } finally {
      await stop(launcher.child);
    }
  }, 60_000);

  it('listens on loopback only: every non-loopback address refuses the port', async () => {
    const dist = makeDist();
    const port = await freePort();
    const launcher = startLauncher(['--dist', dist, '--port', String(port), '--no-open']);
    await waitFor(() => launcher.output().includes('Serving the local build'));

    try {
      const reachable = await refuses(port, '127.0.0.1');
      expect(reachable).toBe(false);

      for (const address of nonLoopbackAddresses()) {
        expect(
          await refuses(port, address),
          `${address}:${port} accepted a connection; the launcher must bind 127.0.0.1 only`,
        ).toBe(true);
      }
    } finally {
      await stop(launcher.child);
    }
  }, 60_000);

  it('reuses a running instance instead of starting a second server', async () => {
    const dist = makeDist();
    const port = await freePort();
    const first = startLauncher(['--dist', dist, '--port', String(port), '--no-open']);
    await waitFor(() => first.output().includes('Serving the local build'));

    try {
      const second = startLauncher(['--dist', dist, '--port', String(port), '--no-open']);
      await waitFor(() => second.child.exitCode !== null, 10_000);
      expect(second.child.exitCode).toBe(0);
      expect(second.output()).toContain(`calcplus is already running at http://127.0.0.1:${port}/`);
    } finally {
      await stop(first.child);
    }
  }, 60_000);

  it('fails with guidance when another program owns the port', async () => {
    const dist = makeDist();
    const other = createServer((_request, response) => response.end('not calcplus'));
    const port = await new Promise<number>((resolvePort, reject) => {
      other.once('error', reject);
      other.listen(0, '127.0.0.1', () => {
        const address = other.address();
        resolvePort(typeof address === 'object' && address !== null ? address.port : 0);
      });
    });

    try {
      const launcher = startLauncher(['--dist', dist, '--port', String(port), '--no-open']);
      await waitFor(() => launcher.child.exitCode !== null, 10_000);
      expect(launcher.child.exitCode).toBe(1);
      const output = launcher.output();
      expect(output).toContain('Another program may be using this port');
      expect(output).toContain('--port N');
    } finally {
      other.closeAllConnections();
      await new Promise<void>((done) => other.close(() => done()));
    }
  }, 60_000);

  it('rejects an unusable port before starting', async () => {
    const launcher = startLauncher(['--port', '0', '--no-open']);
    await waitFor(() => launcher.child.exitCode !== null, 10_000);
    expect(launcher.child.exitCode).toBe(1);
    expect(launcher.output()).toContain('--port 0 is not usable');
  }, 60_000);
});
