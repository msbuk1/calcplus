import { expect, test } from '@playwright/test';
import { clickCell, openFixture } from './grid-helpers';

const fixture = 'm1-chart-unknown.xlsx';

test('fx opens the Insert Function dialog with search, category, list and v1 actions', async ({
  page,
}) => {
  await openFixture(page, fixture);
  await clickCell(page, 1, 1);
  await page.getByTestId('fx-button').click();

  const dialog = page.getByRole('dialog', { name: 'Insert function' });
  await expect(dialog).toBeVisible();
  await expect(dialog.getByTestId('insert-function-search')).toBeFocused();
  await expect(dialog.getByTestId('insert-function-search')).toHaveAttribute(
    'placeholder',
    'Search for a function',
  );
  await expect(dialog.getByTestId('insert-function-category')).toHaveValue('Recently used');
  await expect(dialog.getByTestId('insert-function-category').locator('option')).toHaveText([
    'Recently used',
    'All',
    'Financial',
    'Date & Time',
    'Math & Trig',
    'Statistical',
    'Lookup & Reference',
    'Text',
    'Logical',
    'Information',
  ]);

  // No function has been inserted in this session, so Recently used is empty
  // rather than pretending at a history.
  await expect(dialog.getByTestId('insert-function-empty')).toBeVisible();

  // The v1 dialog has Insert and Cancel only; the V2 argument wizard must not
  // ship its Next / Insert with preview controls.
  await expect(dialog.getByRole('button', { name: 'Insert', exact: true })).toBeVisible();
  await expect(dialog.getByRole('button', { name: 'Cancel' })).toBeVisible();
  await expect(dialog.getByRole('button', { name: 'Next' })).toHaveCount(0);
  await expect(dialog.getByRole('button', { name: 'Insert with preview' })).toHaveCount(0);
});

test('search narrows the function list', async ({ page }) => {
  await openFixture(page, fixture);
  await page.getByTestId('fx-button').click();
  const dialog = page.getByRole('dialog', { name: 'Insert function' });
  const list = dialog.getByTestId('insert-function-list');

  await dialog.getByTestId('insert-function-search').fill('average');
  await expect(list.getByRole('option')).toHaveText(['AVERAGE']);

  await dialog.getByTestId('insert-function-search').fill('lookup');
  await expect(list.getByRole('option')).toHaveText(['HLOOKUP', 'VLOOKUP']);

  await dialog.getByTestId('insert-function-search').fill('zzz');
  await expect(dialog.getByTestId('insert-function-empty')).toBeVisible();
  await expect(list).toHaveCount(0);

  await dialog.getByTestId('insert-function-search').fill('');
  await expect(dialog.getByTestId('insert-function-empty')).toBeVisible();
});

test('a category switch lists only that category and shows its details', async ({ page }) => {
  await openFixture(page, fixture);
  await page.getByTestId('fx-button').click();
  const dialog = page.getByRole('dialog', { name: 'Insert function' });
  const list = dialog.getByTestId('insert-function-list');

  await dialog.getByTestId('insert-function-category').selectOption('Math & Trig');
  await expect(list.getByRole('option')).toHaveCount(10);
  await expect(dialog.getByTestId('insert-function-option-SUM')).toBeVisible();
  await expect(dialog.getByTestId('insert-function-option-AVERAGE')).toHaveCount(0);

  await dialog.getByTestId('insert-function-category').selectOption('All');
  await expect(list.getByRole('option')).toHaveCount(52);
  // The first alphabetical row is highlighted before a choice is made.
  await expect(dialog.getByTestId('insert-function-signature')).toHaveText('ABS(number)');

  await dialog.getByTestId('insert-function-option-AVERAGE').click();
  await expect(dialog.getByTestId('insert-function-option-AVERAGE')).toHaveAttribute(
    'aria-selected',
    'true',
  );
  await expect(dialog.getByTestId('insert-function-signature')).toHaveText(
    'AVERAGE(number1, [number2], ...)',
  );
  await expect(dialog.getByTestId('insert-function-description')).toHaveText(
    'Returns the average (arithmetic mean) of its arguments, which can be numbers or names, arrays or references that contain numbers.',
  );

  // An empty Excel category still renders the capture's picker, with nothing
  // in the list because calcplus publishes no financial function.
  await dialog.getByTestId('insert-function-category').selectOption('Financial');
  await expect(dialog.getByTestId('insert-function-empty')).toBeVisible();
});

test('up and down keys move the selection through the list', async ({ page }) => {
  await openFixture(page, fixture);
  await page.getByTestId('fx-button').click();
  const dialog = page.getByRole('dialog', { name: 'Insert function' });
  await dialog.getByTestId('insert-function-category').selectOption('Math & Trig');

  const search = dialog.getByTestId('insert-function-search');
  await search.focus();
  await page.keyboard.press('ArrowDown');
  // The list is alphabetical: ABS, INT, MOD, ...
  await expect(dialog.getByTestId('insert-function-option-ABS')).toBeFocused();
  await page.keyboard.press('ArrowDown');
  await expect(dialog.getByTestId('insert-function-option-INT')).toBeFocused();
  await expect(dialog.getByTestId('insert-function-signature')).toHaveText('INT(number)');
  await page.keyboard.press('ArrowUp');
  await expect(dialog.getByTestId('insert-function-option-ABS')).toBeFocused();
});

test('Insert puts the function skeleton in the formula bar with the caret inside', async ({
  page,
}) => {
  await openFixture(page, fixture);
  await clickCell(page, 1, 1);
  await page.getByTestId('fx-button').click();
  const dialog = page.getByRole('dialog', { name: 'Insert function' });
  await dialog.getByTestId('insert-function-category').selectOption('All');
  await dialog.getByTestId('insert-function-option-AVERAGE').click();
  await dialog.getByRole('button', { name: 'Insert', exact: true }).click();

  await expect(dialog).toBeHidden();
  const bar = page.getByTestId('formula-bar');
  await expect(bar).toHaveValue('=AVERAGE(');
  await expect(bar).toBeFocused();
  expect(await bar.evaluate((element) => (element as HTMLInputElement).selectionStart)).toBe(9);
  expect(await bar.evaluate((element) => (element as HTMLInputElement).selectionEnd)).toBe(9);
  await expect
    .poll(() => page.evaluate(() => window.__calcplus.getState().focusZone))
    .toBe('formula-bar');

  // Inserting is an editing action, not a cell change: nothing is in the undo
  // history until the user commits the formula.
  expect(await page.evaluate(() => window.__calcplus.getState().canUndo)).toBe(false);
  // The inserted function is now the first Recently used entry.
  await page.getByTestId('fx-button').click();
  await expect(
    page
      .getByRole('dialog', { name: 'Insert function' })
      .getByTestId('insert-function-option-AVERAGE'),
  ).toBeVisible();
});

test('Esc closes the dialog and returns focus to the grid', async ({ page }) => {
  await openFixture(page, fixture);
  await page.getByTestId('fx-button').click();
  const dialog = page.getByRole('dialog', { name: 'Insert function' });
  await expect(dialog).toBeVisible();

  await page.keyboard.press('Escape');
  await expect(dialog).toBeHidden();
  await expect(page.getByTestId('grid-shell')).toBeFocused();
  await expect.poll(() => page.evaluate(() => window.__calcplus.getState().focusZone)).toBe('grid');
});
