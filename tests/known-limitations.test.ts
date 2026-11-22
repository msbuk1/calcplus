import { readFileSync, readdirSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { EXTERNAL_EDIT_UNSUPPORTED_WARNING } from '@io/confirm';
import { saveWorkbookFile } from '@io/file-system';

/**
 * Pinning tests for docs/known-deviations.md L-1 to L-6. These are shipping
 * and environment limitations, so most pins are documentation or artifact
 * assertions: what the launcher, the build script and the docs promise is
 * what ships. Where a claim is environmental (SmartScreen itself, Firefox's
 * missing save-in-place API) the pin covers the code's response to it, not
 * the environment. If one of these fails, the limitation and what ships have
 * parted ways; fix one of them, do not weaken the pin silently.
 */

const root = fileURLToPath(new URL('..', import.meta.url));

function read(relative: string): string {
  return readFileSync(join(root, relative), 'utf8');
}

describe('known limitations (docs/known-deviations.md)', () => {
  it('pins known limitation L-1: the SmartScreen warning is documented in the launcher README and the About dialog', () => {
    // The warning itself is a Windows behaviour no test can trigger; the
    // honest pin is that the documented handling is published: what the
    // warning says, what to click, and how to verify the download without
    // trusting it. distribution.md §8 asks for this in both the README and
    // the About dialog, so both carry the pin.
    const readme = read('launcher/README.md');
    expect(readme).toContain('not code-signed');
    expect(readme).toContain('Windows protected your PC');
    expect(readme).toContain('More info');
    expect(readme).toContain('calcplus.exe');
    expect(readme).toContain('Run anyway');
    expect(readme).toContain('SHA256SUMS');
    const about = read('src/chrome/AboutDialog.tsx');
    expect(about).toContain('Windows protected your PC');
    expect(about).toContain('More info');
    expect(about).toContain('Run anyway');
    expect(about).toContain('SHA256SUMS');
  });

  it('pins known limitation L-2: the build produces no macOS artifact', () => {
    // The release script builds exactly the two v1 artifacts. Adding a macOS
    // target must update this pin and the doc together.
    const script = read('scripts/build-launcher.mjs');
    expect(script).toContain("['windows', 'linux']");
    expect(script).toContain('calcplus-${version}-windows-x64');
    expect(script).toContain('calcplus-${version}-linux-x64');
    expect(script).not.toMatch(/darwin|macos/i);
  });

  describe('L-3 Firefox cannot save in place', () => {
    interface Anchor {
      href: string;
      download: string;
      style: Record<string, string>;
      click?: () => void;
      remove?: () => void;
    }

    let anchor: Anchor;

    beforeEach(() => {
      vi.useFakeTimers();
      anchor = {
        href: '',
        download: '',
        style: {},
        click: () => undefined,
        remove: () => undefined,
      };
      (globalThis as unknown as { window: unknown }).window = {};
      (globalThis as unknown as { document: unknown }).document = {
        createElement: () => anchor,
        body: { appendChild: () => undefined },
      };
      URL.createObjectURL = (() => 'blob:download') as typeof URL.createObjectURL;
      URL.revokeObjectURL = (() => undefined) as typeof URL.revokeObjectURL;
    });

    afterEach(() => {
      vi.useRealTimers();
      delete (globalThis as unknown as { window?: unknown }).window;
      delete (globalThis as unknown as { document?: unknown }).document;
    });

    it('pins known limitation L-3: save downloads a new file and the app explains the download path', async () => {
      // A browser without the File System Access API (Firefox) saves by
      // download, not in place, and the app's message explains the download
      // path. (The once-per-document showing of that message is App state
      // with no unit seam; it is covered by the e2e persistence suite, not
      // pinned here.)
      const saved = await saveWorkbookFile(new Uint8Array([1, 2, 3]), 'Book.xlsx', null, {
        forceFallback: true,
        saveAs: false,
      });
      expect(saved).toEqual({ method: 'download', handle: null, fileName: 'Book.xlsx' });
      expect(anchor.download).toBe('Book.xlsx');
      expect(EXTERNAL_EDIT_UNSUPPORTED_WARNING).toContain('saves a downloaded copy');
    });
  });

  it('pins known limitation L-4: the grid exposes nothing to assistive technology in v1', () => {
    // v1 makes no assistive-technology claim: there is no grid mirror for a
    // screen reader and the PRD records the no-claim. Building the mirror
    // must update this pin and the doc together (the mirror contract is
    // accessibility.md, a v2 contract).
    for (const name of readdirSync(join(root, 'src', 'grid'))) {
      if (!name.endsWith('.tsx') && !name.endsWith('.ts')) continue;
      const source = read(join('src', 'grid', name));
      expect(source, `src/grid/${name} mirrors the grid to assistive technology`).not.toMatch(
        /role="grid"|role="gridcell"|aria-rowindex|aria-colindex|aria-rowcount|aria-colcount/,
      );
    }
    const prd = read('docs/prd.md');
    expect(prd).toContain('assistive-technology claim');
    expect(prd).toContain('makes no accessibility claim');
  });

  it('pins known limitation L-5: the launcher warns when it starts on a non-default port', () => {
    // `--port N` uses a separate browser origin, and the launcher says so at
    // startup. This pins the launcher source in the shape of the "launcher
    // source" tests in tests/m7-launcher.test.ts; the runtime message is
    // exercised there where Go is installed.
    const source = read('launcher/main.go');
    expect(source).toContain('*port != defaultPort');
    expect(source).toContain('has its own browser origin');
    expect(source).toContain('the recovery copy from port');
    expect(source).toContain('are not available here');
  });

  it('pins known limitation L-6: there is no updater and the app shows its version', () => {
    // Updating means replacing the executable; the pin covers the version the
    // app shows, the documented update path, and the launcher's dependency
    // surface (the Go standard library only: no updater rides along).
    const packageJson = JSON.parse(read('package.json')) as { version: string };
    expect(read('src/chrome/AboutDialog.tsx')).toContain(`Version ${packageJson.version}`);
    const readme = read('launcher/README.md');
    expect(readme).toContain('Replace the executable');
    expect(readme).toContain('no background updater');
    expect(read('launcher/go.mod')).not.toContain('require ');
  });
});
