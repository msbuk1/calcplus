import { Buffer } from 'node:buffer';
import { expect, test } from '@playwright/test';
import type { Locator, Page } from '@playwright/test';
import { threeKeyTableXlsx } from './data-tools-fixture';

/**
 * The Sort dialog (sort-autofilter §2; design-refs m5-menus-dialogs §1): the
 * level toolbar, the "My data has headers" override, the key table and the
 * read-only target range. A three-key sort round-trips through `data.sort` and
 * stays one undo step.
 */

async function openSortTable(page: Page): Promise<void> {
  await page.goto('/?fsa=0');
  await page.getByTestId('file-input').setInputFiles({
    name: 'sort-table.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: Buffer.from(threeKeyTableXlsx()),
  });
  await expect(page.getByTestId('file-name')).toHaveText('sort-table.xlsx');
  await expect(page.getByTestId('grid-canvas')).toBeVisible();
}

async function selectRange(page: Page, address: string): Promise<void> {
  const nameBox = page.getByTestId('name-box');
  await nameBox.click();
  await page.keyboard.press('Control+a');
  await page.keyboard.type(address);
  await page.keyboard.press('Enter');
  await expect(nameBox).toHaveValue(address);
}

async function openSortDialog(page: Page): Promise<Locator> {
  await page.getByRole('tab', { name: 'Data' }).click();
  await page.getByRole('button', { name: 'Sort', exact: true }).click();
  const dialog = page.getByRole('dialog', { name: 'Customised Sort' });
  await expect(dialog).toBeVisible();
  return dialog;
}

function rows(page: Page): Promise<(string | null)[][]> {
  return page.evaluate(() =>
    Array.from({ length: 6 }, (_, row) =>
      Array.from({ length: 3 }, (_, col) => window.__calcplus.getCellDisplay(row, col)),
    ),
  );
}

const ORIGINAL = [
  ['Dept', 'City', 'Sales'],
  ['Sales', 'Bath', '100'],
  ['Ops', 'Bath', '120'],
  ['Ops', 'Bath', '90'],
  ['Sales', 'Bristol', '100'],
  ['Ops', 'Bristol', '50'],
];

test('the Sort dialog round-trips a three-key sort as one undo step', async ({ page }) => {
  await openSortTable(page);
  await selectRange(page, 'A1:C6');
  const dialog = await openSortDialog(page);

  // The capture's chrome: toolbar, header override, table headings and range.
  await expect(dialog.getByRole('button', { name: 'Add' })).toBeVisible();
  await expect(dialog.getByRole('button', { name: 'Delete' })).toBeVisible();
  await expect(dialog.getByRole('button', { name: 'Copy' })).toBeVisible();
  await expect(dialog.getByRole('button', { name: 'Move level up' })).toBeVisible();
  await expect(dialog.getByRole('button', { name: 'Move level down' })).toBeVisible();
  await expect(dialog.getByText('My data has headers')).toBeVisible();
  await expect(dialog.getByRole('checkbox', { name: 'My data has headers' })).toBeChecked();
  await expect(dialog.getByRole('columnheader', { name: 'Column' })).toBeVisible();
  await expect(dialog.getByRole('columnheader', { name: 'Sort On' })).toBeVisible();
  await expect(dialog.getByRole('columnheader', { name: 'Order' })).toBeVisible();
  await expect(dialog.getByTestId('sort-target')).toContainText('A1:C6');
  // Options holds v2 items only, so it is omitted rather than rendered empty.
  await expect(dialog.getByRole('button', { name: 'Options' })).toHaveCount(0);

  // One key row to start with, labelled "Sort by"; the header text names it.
  await expect(dialog.getByRole('rowheader', { name: 'Sort by' })).toBeVisible();
  await expect(dialog.getByTestId('sort-key-0-column')).toHaveValue('0');
  await expect(dialog.getByTestId('sort-key-0-on')).toHaveValue('cell-values');
  await expect(dialog.getByTestId('sort-key-0-order')).toHaveValue('ascending');

  // Add the second and third levels.
  await dialog.getByRole('button', { name: 'Add' }).click();
  await dialog.getByRole('button', { name: 'Add' }).click();
  await expect(dialog.getByRole('rowheader', { name: 'Then by' })).toHaveCount(2);

  await dialog.getByTestId('sort-key-0-column').selectOption({ label: 'City' });
  await dialog.getByTestId('sort-key-1-column').selectOption({ label: 'Dept' });
  await dialog.getByTestId('sort-key-1-order').selectOption('descending');
  await dialog.getByTestId('sort-key-2-column').selectOption({ label: 'Sales' });

  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect(dialog).toBeHidden();

  // City ascending, Dept descending, then Sales ascending breaks the Bath tie.
  await expect
    .poll(() => rows(page))
    .toEqual([
      ['Dept', 'City', 'Sales'],
      ['Sales', 'Bath', '100'],
      ['Ops', 'Bath', '90'],
      ['Ops', 'Bath', '120'],
      ['Sales', 'Bristol', '100'],
      ['Ops', 'Bristol', '50'],
    ]);
  await expect(page.getByTestId('status-message')).toContainText('Sorted A1:C6');

  // One sort is one transaction; undo restores the exact prior rows (§2.7).
  await page.keyboard.press('Control+z');
  await expect.poll(() => rows(page)).toEqual(ORIGINAL);
});

test('the header override relabels the columns and sorts the first row as data', async ({
  page,
}) => {
  await openSortTable(page);
  await selectRange(page, 'A1:C6');
  const dialog = await openSortDialog(page);

  await expect(dialog.getByTestId('sort-key-0-column').locator('option')).toHaveText([
    'Dept',
    'City',
    'Sales',
  ]);
  await dialog.getByRole('checkbox', { name: 'My data has headers' }).uncheck();
  await expect(dialog.getByTestId('sort-key-0-column').locator('option')).toHaveText([
    'A',
    'B',
    'C',
  ]);

  // The header row now sorts as data: City ascending puts "Bath" first.
  await dialog.getByTestId('sort-key-0-column').selectOption({ label: 'B' });
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect(dialog).toBeHidden();
  await expect.poll(() => page.evaluate(() => window.__calcplus.getCellDisplay(0, 1))).toBe('Bath');
});

test('the Home Sort & Filter menu entry opens the same dialog', async ({ page }) => {
  await openSortTable(page);
  await selectRange(page, 'A1:C6');

  // The menu entry is "Custom Sort" while the dialog title is "Customised
  // Sort", exactly as the oracle capture shows (ticket 28 D4).
  await page.getByRole('button', { name: 'Sort & Filter' }).click();
  await page.getByRole('menuitem', { name: 'Custom Sort' }).click();
  const dialog = page.getByRole('dialog', { name: 'Customised Sort' });
  await expect(dialog).toBeVisible();
  await expect(dialog.getByTestId('sort-target')).toContainText('A1:C6');
  await page.keyboard.press('Escape');
  await expect(dialog).toBeHidden();
});

test('Esc closes the Sort dialog without a transaction', async ({ page }) => {
  await openSortTable(page);
  await selectRange(page, 'A1:C6');
  const versionBefore = await page.evaluate(() => window.__calcplus.getState().version);

  const dialog = await openSortDialog(page);
  await dialog.getByRole('button', { name: 'Add' }).click();
  await page.keyboard.press('Escape');
  await expect(dialog).toBeHidden();
  await expect(page.getByTestId('grid-shell')).toBeFocused();
  expect(await page.evaluate(() => window.__calcplus.getState().version)).toBe(versionBefore);
  expect(await page.evaluate(() => window.__calcplus.getState().canUndo)).toBe(false);
});
