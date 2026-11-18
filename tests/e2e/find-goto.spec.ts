import { expect, test } from '@playwright/test';
import { openFixture } from './grid-helpers';

const fixture = 'm1-chart-unknown.xlsx';

test('Ctrl+F opens Find, rejects an empty query and Esc returns to the grid', async ({ page }) => {
  await openFixture(page, fixture);

  await page.keyboard.press('Control+f');
  const dialog = page.getByRole('dialog', { name: 'Find' });
  await expect(dialog).toBeVisible();
  await expect(dialog.getByLabel('Find', { exact: true })).toBeFocused();
  // The capture's Direction dropdown, defaulting to Down (ticket 28 D3).
  await expect(dialog.getByLabel('Direction')).toHaveValue('down');

  await dialog.getByRole('button', { name: 'Find next' }).click();
  await expect(dialog.getByTestId('find-error')).toHaveText('Enter a value or formula to find.');
  // A rejected query never dispatches, so the selection stays where it was.
  await expect(page.getByTestId('name-box')).toHaveValue('A1');

  await page.keyboard.press('Escape');
  await expect(dialog).toBeHidden();
  await expect(page.getByTestId('grid-shell')).toBeFocused();
});

test('Find next selects a match and Find all lists it', async ({ page }) => {
  await openFixture(page, fixture);

  await page.getByTestId('name-box').click();
  await page.keyboard.press('Control+a');
  await page.keyboard.type('A1');
  await page.keyboard.press('Enter');

  await page.keyboard.press('Control+f');
  const dialog = page.getByRole('dialog', { name: 'Find' });
  await dialog.getByLabel('Find', { exact: true }).fill('Alpha');
  await dialog.getByRole('button', { name: 'Find next' }).click();
  await expect(page.getByTestId('name-box')).toHaveValue('A2');
  await expect(dialog).toBeVisible();

  await dialog.getByRole('button', { name: 'Find all' }).click();
  const results = dialog.getByTestId('find-results');
  await expect(results).toBeVisible();
  await expect(results.getByTestId('find-result-0')).toContainText('A2');
  await expect(results.getByTestId('find-result-0')).toContainText('Alpha');

  // Choosing a row moves the selection without closing the dialog.
  await dialog.getByTestId('find-result-0').click();
  await expect(page.getByTestId('name-box')).toHaveValue('A2');
  await expect(dialog).toBeVisible();

  await page.keyboard.press('Escape');
  await expect(page.getByTestId('grid-shell')).toBeFocused();
});

test('Find all respects the match case and whole cell options', async ({ page }) => {
  await openFixture(page, fixture);
  await page.keyboard.press('Control+f');
  const dialog = page.getByRole('dialog', { name: 'Find' });

  await dialog.getByLabel('Find', { exact: true }).fill('alpha');
  await dialog.getByRole('button', { name: 'Find all' }).click();
  await expect(dialog.getByTestId('find-summary')).toHaveText('Found 1 match.');
  await expect(dialog.getByTestId('find-result-0')).toContainText('A2');

  await dialog.getByTestId('find-match-case').check();
  await dialog.getByRole('button', { name: 'Find all' }).click();
  await expect(dialog.getByTestId('find-summary')).toHaveText('No matches found.');

  await dialog.getByTestId('find-match-case').uncheck();
  await dialog.getByTestId('find-whole-cell').check();
  await dialog.getByLabel('Find', { exact: true }).fill('Alph');
  await dialog.getByRole('button', { name: 'Find all' }).click();
  await expect(dialog.getByTestId('find-summary')).toHaveText('No matches found.');

  await dialog.getByTestId('find-whole-cell').uncheck();
  await dialog.getByRole('button', { name: 'Find all' }).click();
  await expect(dialog.getByTestId('find-summary')).toHaveText('Found 1 match.');
  await expect(dialog.getByTestId('find-result-0')).toContainText('A2');
});

test('Replace all changes every match in one undo step', async ({ page }) => {
  await openFixture(page, fixture);

  await page.keyboard.press('Control+h');
  const dialog = page.getByRole('dialog', { name: 'Find and Replace' });
  await expect(dialog).toBeVisible();
  await dialog.getByLabel('Find', { exact: true }).fill('Alpha');
  await dialog.getByLabel('Replace with').fill('Gamma');
  await dialog.getByRole('button', { name: 'Replace all' }).click();

  await expect
    .poll(() => page.evaluate(() => window.__calcplus.getCellDisplay(1, 0)))
    .toBe('Gamma');
  await expect(dialog.getByTestId('find-summary')).toContainText('Replaced 1 cell');

  // The replace is one transaction; Esc returns to the grid and undo restores.
  await page.keyboard.press('Escape');
  await expect(page.getByTestId('grid-shell')).toBeFocused();
  await page.keyboard.press('Control+z');
  await expect
    .poll(() => page.evaluate(() => window.__calcplus.getCellDisplay(1, 0)))
    .toBe('Alpha');
});

test('Replace replaces one match through the replace field', async ({ page }) => {
  await openFixture(page, fixture);
  await page.keyboard.press('Control+h');
  const dialog = page.getByRole('dialog', { name: 'Find and Replace' });

  await dialog.getByLabel('Find', { exact: true }).fill('Alpha');
  await dialog.getByLabel('Replace with').fill('Gamma');
  // Enter in the replacement field runs Replace, not Find next.
  await page.keyboard.press('Enter');
  await expect
    .poll(() => page.evaluate(() => window.__calcplus.getCellDisplay(1, 0)))
    .toBe('Gamma');
  await expect(dialog.getByTestId('find-summary')).toContainText('Replaced 1 cell');
});

test('Go to accepts a range and rejects invalid input without moving the selection', async ({
  page,
}) => {
  await openFixture(page, fixture);

  await page.getByTestId('name-box').click();
  await page.keyboard.press('Control+a');
  await page.keyboard.type('B2');
  await page.keyboard.press('Enter');
  await expect(page.getByTestId('name-box')).toHaveValue('B2');

  await page.keyboard.press('Control+g');
  const dialog = page.getByRole('dialog', { name: 'Go to' });
  await expect(dialog).toBeVisible();
  const input = dialog.getByLabel('Reference');
  await expect(input).toBeFocused();
  await input.fill('A1:B2');
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect(dialog).toBeHidden();
  await expect(page.getByTestId('name-box')).toHaveValue('A1:B2');
  await expect(page.getByTestId('grid-shell')).toBeFocused();

  await page.keyboard.press('Control+g');
  await dialog.getByLabel('Reference').fill('not a reference');
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect(dialog.getByTestId('goto-error')).toBeVisible();
  await expect(dialog).toBeVisible();
  // A rejected target leaves the selection exactly where it was.
  await expect(page.getByTestId('name-box')).toHaveValue('A1:B2');

  await page.keyboard.press('Escape');
  await expect(dialog).toBeHidden();
  await expect(page.getByTestId('grid-shell')).toBeFocused();
});

test('Go to returns focus to the invoking ribbon menu button on dismissal', async ({ page }) => {
  await openFixture(page, fixture);
  const menuButton = page.getByRole('button', { name: 'Find & Select' });
  await menuButton.click();
  await page.getByRole('menuitem', { name: 'Go to' }).click();
  const dialog = page.getByRole('dialog', { name: 'Go to' });
  await expect(dialog).toBeVisible();
  await dialog.getByRole('button', { name: 'Cancel' }).click();
  await expect(dialog).toBeHidden();
  await expect(menuButton).toBeFocused();
});

test('Find is reachable from the ribbon, compact bar and context menu', async ({ page }) => {
  await openFixture(page, fixture);

  // Ribbon Editing group: the Find & Select menu carries the three commands.
  await page.getByRole('button', { name: 'Find & Select' }).click();
  await page.getByRole('menuitem', { name: 'Find', exact: true }).click();
  const findDialog = page.getByRole('dialog', { name: 'Find' });
  await expect(findDialog).toBeVisible();
  await page.keyboard.press('Escape');

  // Context menu.
  await page.getByTestId('grid-canvas').click({ button: 'right' });
  await page.getByRole('menuitem', { name: 'Go to' }).click();
  const gotoDialog = page.getByRole('dialog', { name: 'Go to' });
  await expect(gotoDialog).toBeVisible();
  await page.keyboard.press('Escape');

  // Compact bar below 900 px.
  await page.setViewportSize({ width: 800, height: 900 });
  await page.getByRole('button', { name: 'Find', exact: true }).click();
  await expect(findDialog).toBeVisible();
  await page.keyboard.press('Escape');
  await expect(page.getByTestId('grid-shell')).toBeFocused();
});
