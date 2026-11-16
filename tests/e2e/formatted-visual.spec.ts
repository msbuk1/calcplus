import { Buffer } from 'node:buffer';
import { expect, test } from '@playwright/test';
import { formattedSheetXlsx } from '../helpers/formatted-sheet';

/**
 * The M4A visual baseline. It is pinned to a fixed viewport, scale factor and
 * locale, with a 0.1% pixel budget and a 2/255 per-channel threshold
 * (testing-conformance §6). The snapshot must be recorded on the pinned
 * Chromium runner of the maintainer's PC:
 *   npx playwright test tests/e2e/formatted-visual.spec.ts --update-snapshots
 * and then committed with the change. Rerun the comparison with:
 *   npx playwright test tests/e2e/formatted-visual.spec.ts
 */
test.describe('M4A formatted sheet', () => {
  test.use({
    viewport: { width: 1280, height: 800 },
    deviceScaleFactor: 1,
    colorScheme: 'light',
    reducedMotion: 'reduce',
  });

  test('paints every imported formatting feature', async ({ page }) => {
    await page.goto('/?fsa=0');
    await page.getByTestId('file-input').setInputFiles({
      name: 'm4a-formatted.xlsx',
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      buffer: Buffer.from(formattedSheetXlsx()),
    });
    await expect(page.getByTestId('file-name')).toHaveText('m4a-formatted.xlsx');
    await expect(page.getByTestId('grid-canvas')).toBeVisible();
    await page.waitForFunction(
      () => window.__calcplus.getCellDisplay(0, 0) === 'Quarterly report - formatted',
    );
    await page.evaluate(() => document.fonts.ready.then(() => true));
    // Wait for the scheduled canvas paint, not just the projection update.
    await page.evaluate(
      () =>
        new Promise<void>((resolve) => {
          requestAnimationFrame(() => requestAnimationFrame(() => resolve()));
        }),
    );

    // Geometry: the imported hidden row and column are collapsed, and the
    // imported sizes survive.
    const hiddenRow = await page.evaluate(() => window.__calcplus.getCellRect(5, 0));
    const hiddenColumn = await page.evaluate(() => window.__calcplus.getCellRect(0, 5));
    expect(hiddenRow?.height).toBe(0);
    expect(hiddenColumn?.width).toBe(0);
    const visibleRow = await page.evaluate(() => window.__calcplus.getCellRect(3, 0));
    expect(visibleRow?.height).toBeGreaterThan(20);

    // Projection: the renderer receives ready-to-paint formatting.
    const shrink = await page.evaluate(() => window.__calcplus.getCellStyle(4, 0));
    expect(shrink?.alignment.shrinkToFit).toBe(true);
    const rotation = await page.evaluate(() => window.__calcplus.getCellStyle(4, 1));
    expect(rotation?.alignment.textRotation).toBe(45);
    const underline = await page.evaluate(() => window.__calcplus.getCellStyle(4, 2));
    expect(underline?.font.underlineStyle).toBe('single');
    const doubleUnderline = await page.evaluate(() => window.__calcplus.getCellStyle(4, 3));
    expect(doubleUnderline?.font.underlineStyle).toBe('double');
    expect(doubleUnderline?.font.strike).toBe(true);
    const pattern = await page.evaluate(() => window.__calcplus.getCellStyle(6, 0));
    expect(pattern?.fill.pattern).toBe('lightGrid');
    expect(pattern?.fill.backgroundColor?.toLowerCase()).toBe('#ffffff');
    const merged = await page.evaluate(() => window.__calcplus.getCellStyle(0, 1));
    expect(merged?.merge).toEqual({ anchorRow: 0, anchorCol: 0, rowSpan: 1, colSpan: 4 });

    await expect(page.getByTestId('grid-canvas')).toHaveScreenshot('m4a-formatted-sheet.png', {
      threshold: 2 / 255,
      maxDiffPixelRatio: 0.001,
      animations: 'disabled',
    });
  });
});
