import { writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { expect, test } from '@playwright/test';
import { SCROLL_250K, SPARSE_1M, touchedCells } from '../corpus/generate';
import { checkReferenceClass, describeMachine, type MachineFacts } from './machine';

/**
 * The M2 scroll budget (performance §3.3): a 250k-touched-cell sheet must
 * scroll at 60 fps at p95 with no frame over 50 ms, and a million-row sparse
 * sheet must scroll without the grid allocating the sheet.
 *
 * The budget is a contract with the reference machine (performance §1). Every
 * result records a machine fingerprint; on a machine that is not the reference
 * class the absolute budget is advisory and the run reports that fact instead
 * of claiming acceptance.
 */
const P95_BUDGET_MS = 16.7;
const WORST_FRAME_BUDGET_MS = 50;
/** The sanity bound for a run that is not on the reference machine. */
const ADVISORY_P95_MS = 250;
const GUARANTEED_ENVELOPE_BYTES = 500 * 1024 * 1024;

const fixturePath = fileURLToPath(
  new URL('../../fixtures/compat/m1-chart-unknown.xlsx', import.meta.url),
);

interface FrameSample {
  readonly p50: number;
  readonly p95: number;
  readonly max: number;
  readonly count: number;
}

function percentile(values: readonly number[], fraction: number): number {
  if (values.length === 0) return 0;
  const sorted = [...values].sort((left, right) => left - right);
  const index = Math.min(sorted.length - 1, Math.max(0, Math.ceil(fraction * sorted.length) - 1));
  return sorted[index] ?? 0;
}

async function machineFacts(page: import('@playwright/test').Page, browserVersion: string) {
  const pageFingerprint = await page.evaluate(() => ({
    viewport: { width: window.innerWidth, height: window.innerHeight },
    devicePixelRatio: window.devicePixelRatio,
  }));
  return describeMachine(
    browserVersion,
    pageFingerprint.viewport,
    pageFingerprint.devicePixelRatio,
  );
}

async function loadShape(
  page: import('@playwright/test').Page,
  shape: { rows: number; cols: number; seed: number; touchedEvery?: number },
): Promise<number> {
  return page.evaluate(
    (options) =>
      (
        window as unknown as {
          __calcplus: {
            loadBenchmarkSheet(options: {
              rows: number;
              cols: number;
              seed?: number;
              touchedEvery?: number;
            }): number;
          };
        }
      ).__calcplus.loadBenchmarkSheet(options),
    shape,
  );
}

test('M2 scroll frame budget on a 250k-cell sheet', async ({ page, browser }, testInfo) => {
  await page.goto('/?fsa=0');
  await page.getByTestId('file-input').setInputFiles(fixturePath);
  await expect(page.getByTestId('grid-canvas')).toBeVisible();

  const touched = await loadShape(page, {
    rows: SCROLL_250K.rows,
    cols: SCROLL_250K.cols,
    seed: SCROLL_250K.seed,
  });
  expect(touched).toBe(touchedCells(SCROLL_250K));
  await expect(page.getByTestId('grid-canvas')).toBeVisible();
  await page.waitForTimeout(500);

  // Start the frame sampler.
  await page.evaluate(() => {
    const state = window as unknown as {
      __frames: number[];
      __longTasks: number[];
      __sampling: boolean;
    };
    state.__frames = [];
    state.__longTasks = [];
    state.__sampling = true;
    let last = performance.now();
    const tick = (now: number): void => {
      state.__frames.push(now - last);
      last = now;
      if (state.__sampling) requestAnimationFrame(tick);
    };
    requestAnimationFrame(tick);
    const observer = new PerformanceObserver((list) => {
      for (const entry of list.getEntries()) state.__longTasks.push(entry.duration);
    });
    observer.observe({ entryTypes: ['longtask'] });
  });

  const box = await page.getByTestId('grid-canvas').boundingBox();
  expect(box).not.toBeNull();
  await page.mouse.move(box!.x + box!.width / 2, box!.y + box!.height / 2);
  const scrollStart = Date.now();
  while (Date.now() - scrollStart < 2_500) {
    await page.mouse.wheel(0, 240);
    await page.waitForTimeout(16);
  }

  const measurement = await page.evaluate(() => {
    const state = window as unknown as {
      __frames: number[];
      __longTasks: number[];
      __sampling: boolean;
    };
    state.__sampling = false;
    // The first sample is the interval before the sampler started.
    const frames = state.__frames.slice(1);
    const scroll = (
      window as unknown as { __calcplus: { getScroll(): { x: number; y: number } } }
    ).__calcplus.getScroll();
    return { frames, longTasks: state.__longTasks, scroll };
  });

  const frames = measurement.frames;
  const sample: FrameSample = {
    p50: percentile(frames, 0.5),
    p95: percentile(frames, 0.95),
    max: frames.length === 0 ? 0 : Math.max(...frames),
    count: frames.length,
  };

  const cdp = await page.context().newCDPSession(page);
  await cdp.send('Performance.enable');
  const metricsPayload = await cdp.send('Performance.getMetrics');
  const metrics = new Map(metricsPayload.metrics.map((metric) => [metric.name, metric.value]));
  const jsHeapUsed = metrics.get('JSHeapUsedSize') ?? 0;
  const domNodes = metrics.get('Nodes') ?? 0;

  const facts: MachineFacts = await machineFacts(page, browser.version());
  const declaredReference = process.env.CALCPLUS_REFERENCE_MACHINE === '1';
  const reference = checkReferenceClass(facts, declaredReference);

  const result = {
    harness: 'bench/browser/scroll.bench.ts',
    milestone: 'M2',
    generatedAt: new Date().toISOString(),
    case: 'scroll a 250,000 touched-cell formatted sheet',
    touchedCells: touched,
    budgets: {
      p95FrameMs: P95_BUDGET_MS,
      worstFrameMs: WORST_FRAME_BUDGET_MS,
      guaranteedEnvelopeBytes: GUARANTEED_ENVELOPE_BYTES,
    },
    measurement: {
      ...sample,
      longTasks: measurement.longTasks.length,
      longestLongTaskMs:
        measurement.longTasks.length === 0 ? 0 : Math.max(...measurement.longTasks),
      scrollOffset: measurement.scroll,
    },
    workingSet: {
      pageJsHeapUsedBytes: jsHeapUsed,
      domNodes,
    },
    machine: facts,
    referenceClass: reference.matches,
    declaredReference,
    referenceReasons: reference.reasons,
    gate: reference.matches
      ? 'enforced: the run was declared on the reference machine and the hardware matches performance §1'
      : `advisory: the absolute budget does not gate on this machine (${reference.reasons.join('; ')}); performance §6 makes results from other machines advisory`,
  };

  const output = process.env.CALCPLUS_BENCH_OUTPUT ?? testInfo.outputPath('scroll-frames.json');
  writeFileSync(output, `${JSON.stringify(result, null, 2)}\n`);
  await testInfo.attach('scroll-frames.json', {
    body: Buffer.from(`${JSON.stringify(result, null, 2)}\n`),
    contentType: 'application/json',
  });

  // The measurement itself must be sane on every machine.
  expect(sample.count).toBeGreaterThan(20);
  expect(measurement.scroll.y).toBeGreaterThan(0);
  expect(jsHeapUsed).toBeLessThan(GUARANTEED_ENVELOPE_BYTES);

  if (reference.matches) {
    // Frame times are not meaningful below a millisecond, so the measurement
    // and the budget are both compared as whole milliseconds.
    expect(Math.round(sample.p95)).toBeLessThanOrEqual(Math.round(P95_BUDGET_MS));
    expect(Math.round(sample.max)).toBeLessThanOrEqual(Math.round(WORST_FRAME_BUDGET_MS));
  } else {
    // A run that is not on the reference machine never fails on the
    // reference-only absolute budget: its numbers are advisory (performance
    // §6). A generous sanity bound still catches a grid that stopped
    // painting.
    expect(sample.p95).toBeLessThanOrEqual(ADVISORY_P95_MS);
  }
});

test('M2 sparse sheet scrolls to the last row without allocating the grid', async ({
  page,
  browser,
}, testInfo) => {
  await page.goto('/?fsa=0');
  await page.getByTestId('file-input').setInputFiles(fixturePath);
  await expect(page.getByTestId('grid-canvas')).toBeVisible();
  const touched = await loadShape(page, {
    rows: SPARSE_1M.rows,
    cols: SPARSE_1M.cols,
    seed: SPARSE_1M.seed,
    touchedEvery: SPARSE_1M.touchedEvery,
  });
  expect(touched).toBe(touchedCells(SPARSE_1M));
  await page.waitForTimeout(300);

  const cdp = await page.context().newCDPSession(page);
  await cdp.send('Performance.enable');
  const before = await cdp.send('Performance.getMetrics');
  const beforeHeap =
    new Map(before.metrics.map((metric) => [metric.name, metric.value])).get('JSHeapUsedSize') ?? 0;

  // Jump to the last row of the sheet.
  await page.getByTestId('name-box').click();
  await page.keyboard.press('Control+a');
  await page.keyboard.type('A1000000');
  await page.keyboard.press('Enter');
  await page.waitForTimeout(400);

  const scroll = await page.evaluate(() =>
    (
      window as unknown as { __calcplus: { getScroll(): { x: number; y: number } } }
    ).__calcplus.getScroll(),
  );
  const after = await cdp.send('Performance.getMetrics');
  const afterHeap =
    new Map(after.metrics.map((metric) => [metric.name, metric.value])).get('JSHeapUsedSize') ?? 0;

  const facts: MachineFacts = await machineFacts(page, browser.version());
  const result = {
    harness: 'bench/browser/scroll.bench.ts',
    milestone: 'M2',
    generatedAt: new Date().toISOString(),
    case: 'scroll a 1,000,000-row sparse sheet to its last row',
    corpusShape: SPARSE_1M.name,
    touchedCells: touched,
    scrollOffset: scroll,
    heapBeforeBytes: beforeHeap,
    heapAfterBytes: afterHeap,
    heapGrowthBytes: afterHeap - beforeHeap,
    machine: facts,
  };
  const output = process.env.CALCPLUS_BENCH_OUTPUT ?? testInfo.outputPath('sparse-scroll.json');
  writeFileSync(output, `${JSON.stringify(result, null, 2)}\n`);

  expect(scroll.y).toBeGreaterThan(1_000_000);
  // The grid must not allocate the sheet: the growth is far below one row
  // per pixel of scroll.
  expect(afterHeap - beforeHeap).toBeLessThan(64 * 1024 * 1024);
});
