import { readFileSync, writeFileSync } from 'node:fs';
import { cpus, totalmem } from 'node:os';
import { fileURLToPath } from 'node:url';
import { expect, test } from '@playwright/test';

const fixturePath = fileURLToPath(
  new URL('../../fixtures/compat/m1-chart-unknown.xlsx', import.meta.url),
);

/**
 * M1 browser memory harness (performance contract §6, strategy 2). It stands
 * up the CDP measurement plumbing and records a machine fingerprint; the full
 * corpus and absolute budgets are enforced by the nightly suite in later
 * milestones. The M1 smoke assertion only proves the harness can read heap
 * metrics and that the tiny fixture stays under the guaranteed envelope.
 */
test('M1 browser memory harness', async ({ page, browser }, testInfo) => {
  await page.goto('/?fsa=0');
  await page.getByTestId('file-input').setInputFiles(fixturePath);
  await expect(page.getByTestId('file-name')).toHaveText('m1-chart-unknown.xlsx');
  await expect(page.getByTestId('grid-canvas')).toBeVisible();

  const cdp = await page.context().newCDPSession(page);
  await cdp.send('Performance.enable');
  const performance = await cdp.send('Performance.getMetrics');
  const metrics = new Map(performance.metrics.map((metric) => [metric.name, metric.value]));
  const jsHeapUsed = metrics.get('JSHeapUsedSize') ?? 0;
  const jsHeapTotal = metrics.get('JSHeapTotalSize') ?? 0;
  const domNodes = metrics.get('Nodes') ?? 0;
  const canvas = await page.getByTestId('grid-canvas').boundingBox();

  const pageFingerprint = await page.evaluate(() => ({
    userAgent: navigator.userAgent,
    hardwareConcurrency: navigator.hardwareConcurrency,
    deviceMemory: (navigator as unknown as { deviceMemory?: number }).deviceMemory ?? null,
    viewport: { width: window.innerWidth, height: window.innerHeight },
    devicePixelRatio: window.devicePixelRatio,
  }));
  const cpu = cpus()[0];

  const result = {
    harness: 'bench/browser/memory.bench.ts',
    milestone: 'M1',
    generatedAt: new Date().toISOString(),
    fixture: 'fixtures/compat/m1-chart-unknown.xlsx',
    guaranteedEnvelopeBytes: 500 * 1024 * 1024,
    workingSet: {
      pageJsHeapUsedBytes: jsHeapUsed,
      pageJsHeapTotalBytes: jsHeapTotal,
      domNodes,
      engineWorker: null,
      canvasCssPixels: canvas === null ? null : Math.round(canvas.width * canvas.height),
    },
    machine: {
      platform: process.platform,
      arch: process.arch,
      cpuModel: cpu?.model ?? 'unknown',
      cpuCount: cpus().length,
      totalMemoryBytes: totalmem(),
      browserVersion: browser.version(),
      ...pageFingerprint,
    },
    notes:
      'M1 has no formula engine worker or 250k-cell corpus yet. This harness records the page heap and fingerprint; later milestones add the worker target and absolute envelope cases.',
  };

  expect(jsHeapUsed).toBeGreaterThan(0);
  expect(jsHeapUsed).toBeLessThan(result.guaranteedEnvelopeBytes);

  const output = process.env.CALCPLUS_BENCH_OUTPUT ?? testInfo.outputPath('browser-memory.json');
  writeFileSync(output, `${JSON.stringify(result, null, 2)}\n`);
  await testInfo.attach('browser-memory.json', {
    body: Buffer.from(JSON.stringify(result, null, 2)),
    contentType: 'application/json',
  });
  await testInfo.attach('fixture-bytes', {
    body: readFileSync(fixturePath),
    contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
});
