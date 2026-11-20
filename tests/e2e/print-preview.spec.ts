import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';
import { openFixture } from './grid-helpers';

/**
 * M7 lane 2 — the print preview window and `window.print()`.
 *
 * `print-areas.xlsx` is A4 landscape with two print areas, so the page count,
 * navigation and the A4 page-box trap are all exercised. The print output is
 * asserted structurally: the stacked page elements, and the real page count of
 * a PDF printed from the same document. No snapshot images are created.
 */

declare global {
  interface Window {
    __printCalls: number;
  }
}

async function openPreview(page: Page) {
  await page.getByTestId('file-menu-button').click();
  await page.getByRole('menuitem', { name: /^Print/ }).click();
  const dialog = page.getByRole('dialog', { name: 'Print Preview' });
  await expect(dialog).toBeVisible();
  return dialog;
}

test.beforeEach(async ({ page }) => {
  // Never open the real dialog: record that `window.print()` was reached.
  await page.addInitScript(() => {
    window.__printCalls = 0;
    window.print = () => {
      window.__printCalls += 1;
    };
  });
});

test('File → Print opens a two-page A4 preview that navigates by keyboard', async ({ page }) => {
  await openFixture(page, 'print-areas.xlsx');
  const dialog = await openPreview(page);

  // The map-first entry is on the File menu and the summary names the plan.
  await expect(dialog.getByTestId('print-preview-summary')).toContainText('A4 landscape');
  await expect(dialog.getByTestId('print-preview-summary')).toContainText('2 pages');

  // The SVG is the page itself, and it is labelled for a screen reader.
  const svg = dialog.locator('[data-testid="print-preview-page"] svg');
  await expect(svg).toHaveCount(1);
  await expect(svg).toHaveAttribute('aria-label', /^Page 1 of 2, /);
  await expect(svg).toHaveAttribute('role', 'img');

  // The live region announces the page and the buttons move between pages.
  const count = dialog.getByTestId('print-preview-count');
  await expect(count).toHaveText('Page 1 of 2');
  await expect(count).toHaveAttribute('aria-live', 'polite');
  await expect(dialog.getByRole('button', { name: 'Previous page' })).toBeDisabled();

  await dialog.getByRole('button', { name: 'Next page' }).click();
  await expect(count).toHaveText('Page 2 of 2');
  await expect(svg).toHaveAttribute('aria-label', /^Page 2 of 2, /);
  await expect(dialog.getByRole('button', { name: 'Next page' })).toBeDisabled();

  // Arrow keys, Page Up/Down and Home/End work from anywhere in the modal.
  await page.keyboard.press('ArrowLeft');
  await expect(count).toHaveText('Page 1 of 2');
  await page.keyboard.press('End');
  await expect(count).toHaveText('Page 2 of 2');
  await page.keyboard.press('PageUp');
  await expect(count).toHaveText('Page 1 of 2');
  await page.keyboard.press('PageDown');
  await expect(count).toHaveText('Page 2 of 2');
  await page.keyboard.press('Home');
  await expect(count).toHaveText('Page 1 of 2');

  // Escape closes and returns the keyboard to the grid.
  await page.keyboard.press('Escape');
  await expect(dialog).toBeHidden();
  await expect(page.getByTestId('grid-shell')).toBeFocused();
});

test('Print reaches window.print() with one page element per plan page', async ({ page }) => {
  await openFixture(page, 'print-areas.xlsx');
  const dialog = await openPreview(page);

  // Under print media the shell hides the chrome and the dialog, and leaves
  // only the plan's pages: one element per page, each with its own SVG.
  await page.emulateMedia({ media: 'print' });
  await expect(page.getByTestId('grid-shell')).toBeHidden();
  await expect(dialog).toBeHidden();
  const pages = page.getByTestId('print-page');
  await expect(pages).toHaveCount(2);
  for (let index = 0; index < 2; index += 1) {
    const block = pages.nth(index);
    await expect(block.locator('svg')).toHaveCount(1);
    await expect(block.locator('svg')).toHaveAttribute('data-print-page', String(index + 1));
    const box = await block.boundingBox();
    expect(box).not.toBeNull();
    // 841.92 × 594.96pt is Chromium's real A4 landscape page box; the page
    // element stays inside it, so the nominal 841.89pt width never overflows.
    expect((box?.width ?? 0) * 0.75).toBeLessThan(841.92);
    expect((box?.height ?? 0) * 0.75).toBeLessThan(594.96);
    expect((box?.width ?? 0) * 0.75).toBeGreaterThan(800);
  }
  await page.emulateMedia({ media: 'screen' });

  await dialog.getByRole('button', { name: 'Print', exact: true }).click();
  await expect.poll(() => page.evaluate(() => window.__printCalls)).toBe(1);
});

test('a print of the preview is exactly the plan’s page count, with no empty page', async ({
  page,
}) => {
  await openFixture(page, 'print-areas.xlsx');
  await openPreview(page);
  await expect(page.getByTestId('print-page')).toHaveCount(2);

  // The same print path `window.print()` drives; `printBackground: false` is
  // the default "Background graphics off" toggle, which must not change the
  // page count (all paint is SVG, never a CSS background).
  const pdf = await page.pdf({ printBackground: false, preferCSSPageSize: true });
  const text = pdf.toString('latin1');
  const counts = [...text.matchAll(/\/Type\s*\/Pages[\s\S]{0,200}?\/Count\s+(\d+)/g)].map((match) =>
    Number(match[1]),
  );
  expect(counts.length).toBeGreaterThan(0);
  expect(Math.max(...counts)).toBe(2);
});
