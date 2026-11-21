import { defineConfig, devices } from '@playwright/test';

/**
 * The suite runs against a production build on its own port, not the dev
 * server.
 *
 * Running against `npm run dev` made the suite flaky: Vite transforms modules
 * on demand and accumulates HMR state, so several workers hitting one dev
 * server produce failures that vanish when the test runs alone. `reuseExistingServer`
 * made it worse by silently adopting whatever was already listening.
 *
 * `vite preview` serves the built `dist/`, which is what actually ships, and
 * its module graph is settled before the first test runs. The port is separate
 * from the dev server's 4173 so a test run never fights a dev server, and
 * `reuseExistingServer` is off so a stale process is a clear failure rather
 * than a quiet one.
 *
 * The build is part of the command, so `npm run test:e2e` stays one step. It
 * costs about 30 seconds and buys determinism.
 */
export default defineConfig({
  testDir: './tests/e2e',
  fullyParallel: true,
  reporter: 'list',
  use: {
    baseURL: 'http://127.0.0.1:4177',
    trace: 'retain-on-failure',
    viewport: { width: 1280, height: 800 },
    locale: 'en-US',
    timezoneId: 'UTC',
  },
  projects: [
    {
      name: 'chromium',
      use: { ...devices['Desktop Chrome'] },
    },
  ],
  webServer: {
    command: 'npm run build && npm run preview -- --host 127.0.0.1 --port 4177 --strictPort',
    url: 'http://127.0.0.1:4177',
    reuseExistingServer: false,
    timeout: 180_000,
  },
});
