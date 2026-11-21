import { expect, test } from '@playwright/test';
import { clickCell, openFixture } from './grid-helpers';

/**
 * An invalid formula is refused at entry, as Excel refuses it: the cell keeps
 * its previous content and the editor stays open with the text so it can be
 * fixed. Closing the editor and discarding the text made the cell look cleared
 * and left the user with nothing to correct.
 */

const fixture = 'm1-chart-unknown.xlsx';

test('a refused formula keeps the editor open with the typed text', async ({ page }) => {
  await openFixture(page, fixture);

  // A cell that already holds a value keeps it, and the text stays editable.
  await clickCell(page, 1, 1);
  await page.keyboard.type('=SUM()');
  await page.keyboard.press('Enter');

  const editor = page.getByTestId('cell-editor');
  await expect(editor).toBeVisible();
  await expect(editor).toHaveValue('=SUM()');
  expect(await page.evaluate(() => window.__calcplus.getCellDisplay(1, 1))).toBe('42.5');
  expect(await page.evaluate(() => window.__calcplus.getState().status)).toContain(
    'The cell was not changed.',
  );

  // Correcting the formula commits it normally.
  await editor.fill('=SUM(1,2)');
  await page.keyboard.press('Enter');
  await expect(editor).toBeHidden();
  await expect.poll(() => page.evaluate(() => window.__calcplus.getCellDisplay(1, 1))).toBe('3');
});

test('an empty cell stays empty when the formula is refused, and Esc abandons it', async ({
  page,
}) => {
  await openFixture(page, fixture);
  await clickCell(page, 5, 5);
  await page.keyboard.type('=SUM()');
  await page.keyboard.press('Enter');

  const editor = page.getByTestId('cell-editor');
  await expect(editor).toBeVisible();
  await expect(editor).toHaveValue('=SUM()');
  expect(await page.evaluate(() => window.__calcplus.getCellDisplay(5, 5))).toBeNull();

  await page.keyboard.press('Escape');
  await expect(editor).toBeHidden();
  expect(await page.evaluate(() => window.__calcplus.getCellDisplay(5, 5))).toBeNull();
});

test('the formula bar keeps a refused draft', async ({ page }) => {
  await openFixture(page, fixture);
  await clickCell(page, 7, 7);
  const bar = page.getByTestId('formula-bar');
  await bar.fill('=SUM()');
  await bar.press('Enter');

  await expect(bar).toHaveValue('=SUM()');
  expect(await page.evaluate(() => window.__calcplus.getCellDisplay(7, 7))).toBeNull();
  expect(await page.evaluate(() => window.__calcplus.getState().status)).toContain(
    'The cell was not changed.',
  );

  // Fixing it in the bar commits.
  await bar.fill('=SUM(2,3)');
  await bar.press('Enter');
  await expect.poll(() => page.evaluate(() => window.__calcplus.getCellDisplay(7, 7))).toBe('5');
});
