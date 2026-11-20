import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';
import { clickCell, dragSelect, openFixture } from './grid-helpers';

/**
 * The three confirmed v1 gaps wired by the feature lane: Merge on the ribbon,
 * Clear All, Close with its unsaved prompt, and function autocomplete in both
 * formula-editing surfaces.
 */

const fixture = 'm2-grid.xlsx';

async function cellDisplay(page: Page, row: number, col: number): Promise<string | null> {
  return page.evaluate(
    ([targetRow, targetCol]) =>
      window.__calcplus.getCellDisplay(targetRow as number, targetCol as number),
    [row, col] as const,
  );
}

test('Merge & Center sits in the Alignment group with its four menu items', async ({ page }) => {
  await openFixture(page, fixture);
  await dragSelect(page, 2, 1, 3, 2);
  const merge = page.locator('[data-command="merge.center"]').first();
  await expect(merge).toBeVisible();
  await expect(merge).toBeEnabled();
  await expect(merge).toHaveAttribute('title', 'Merge & Center');

  // The dropdown lists the merge family, all live.
  await page
    .locator('.ribbon-split')
    .filter({ has: merge })
    .locator('.ribbon-split-toggle')
    .click();
  const menu = page.getByRole('menu');
  await expect(menu.getByRole('menuitem')).toHaveText([
    'Merge & Center',
    'Merge Across',
    'Merge Cells',
    'Unmerge',
  ]);
});

test('Merge & Center asks before discarding data and merges one step', async ({ page }) => {
  await openFixture(page, fixture);
  await dragSelect(page, 2, 1, 3, 2);

  let prompt = '';
  page.once('dialog', async (dialog) => {
    prompt = dialog.message();
    await dialog.accept();
  });
  await page.locator('[data-command="merge.center"]').first().click();
  await expect.poll(() => prompt).toContain('data here');
  await expect
    .poll(() => page.evaluate(() => window.__calcplus.getCellStyle(2, 1)?.merge))
    .toEqual({ anchorRow: 2, anchorCol: 1, rowSpan: 2, colSpan: 2 });

  await page.getByTestId('undo-button').click();
  await expect
    .poll(() => page.evaluate(() => window.__calcplus.getCellStyle(2, 1)?.merge))
    .toBeNull();
});

test('Clear All clears values and formats as one undo step', async ({ page }) => {
  await openFixture(page, fixture);
  await dragSelect(page, 2, 1, 3, 2);
  const before = await cellDisplay(page, 2, 1);
  expect(before).not.toBeNull();

  const clearToggle = page
    .locator('.ribbon-split')
    .filter({ has: page.locator('[data-command="edit.clearContents"]') })
    .locator('.ribbon-split-toggle');
  await clearToggle.click();
  await page.getByRole('menuitem', { name: 'Clear All', exact: true }).click();

  await expect.poll(() => cellDisplay(page, 2, 1)).toBeNull();
  await expect.poll(() => cellDisplay(page, 3, 2)).toBeNull();
  await expect(page.getByTestId('undo-button')).toBeEnabled();

  await page.getByTestId('undo-button').click();
  await expect.poll(() => cellDisplay(page, 2, 1)).toBe(before);
});

test('Close discards a clean workbook and prompts before a dirty one', async ({ page }) => {
  await openFixture(page, fixture);

  // A clean workbook closes without a prompt.
  await page.getByTestId('file-menu-button').click();
  await page.getByRole('menuitem', { name: 'Close', exact: true }).click();
  await expect(page.getByTestId('empty-open-button')).toBeVisible();

  // A dirty workbook asks first; declining keeps it open.
  await openFixture(page, fixture);
  await clickCell(page, 1, 1);
  await page.keyboard.type('9');
  await page.keyboard.press('Enter');

  let prompt = '';
  page.once('dialog', async (dialog) => {
    prompt = dialog.message();
    await dialog.dismiss();
  });
  await page.getByTestId('file-menu-button').click();
  await page.getByRole('menuitem', { name: 'Close', exact: true }).click();
  await expect.poll(() => prompt).toContain('unsaved changes');
  await expect(page.getByTestId('grid-canvas')).toBeVisible();
  await expect(page.getByTestId('save-button')).toHaveAttribute('data-unsaved', 'true');

  // Accepting discards the session and returns to the empty state.
  page.once('dialog', async (dialog) => dialog.accept());
  await page.getByTestId('file-menu-button').click();
  await page.getByRole('menuitem', { name: 'Close', exact: true }).click();
  await expect(page.getByTestId('empty-open-button')).toBeVisible();
});

test('the formula bar suggests function names and accepts the highlighted one', async ({
  page,
}) => {
  await openFixture(page, fixture);
  await clickCell(page, 1, 1);
  const bar = page.getByTestId('formula-bar');
  await bar.click();
  await bar.press('Control+a');
  await page.keyboard.type('=SU');

  const suggestions = page.getByTestId('formula-suggestions');
  await expect(suggestions).toBeVisible();
  await expect(suggestions.getByRole('option')).toHaveCount(2);
  await expect(suggestions.getByTestId('formula-suggestion-SUM')).toBeVisible();
  await expect(suggestions.getByTestId('formula-suggestion-SUMIF')).toBeVisible();
  await expect(page.getByTestId('formula-suggestion-signature')).toHaveText(
    'SUM(number1, [number2], ...)',
  );

  await page.keyboard.press('ArrowDown');
  await expect(page.getByTestId('formula-suggestion-SUMIF')).toHaveAttribute(
    'aria-selected',
    'true',
  );
  await page.keyboard.press('ArrowUp');
  await page.keyboard.press('Enter');
  await expect(bar).toHaveValue('=SUM(');
  // The caret sits after the opening parenthesis, so the signature hint shows
  // the argument being edited.
  await expect(page.getByTestId('formula-signature')).toContainText('SUM(number1');
});

test('the in-cell editor suggests function names and inserts the skeleton', async ({ page }) => {
  await openFixture(page, fixture);
  await clickCell(page, 1, 1);
  await page.keyboard.type('=AV');

  const suggestions = page.getByTestId('formula-suggestions');
  await expect(suggestions).toBeVisible();
  await expect(suggestions.getByRole('option')).toHaveCount(1);
  await expect(suggestions.getByTestId('formula-suggestion-AVERAGE')).toBeVisible();
  await page.keyboard.press('Enter');
  await expect(page.getByTestId('cell-editor')).toHaveValue('=AVERAGE(');
  await page.keyboard.press('Escape');
  await expect(page.getByTestId('cell-editor')).toBeHidden();
});
