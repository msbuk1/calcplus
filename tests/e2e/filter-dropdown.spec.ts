import { Buffer } from 'node:buffer';
import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';
import { cityTableXlsx } from './data-tools-fixture';

/**
 * The AutoFilter dropdown (sort-autofilter §3.8; design-refs
 * m5-menus-dialogs §2). The capture's ruled deltas from ticket 28 are asserted
 * here too: the Apply button, the "Customised Sort" entry, the
 * "Clear Filter from '<column>'" wording and the absence of the greyed
 * out-of-scope entries.
 */

async function openCityTable(page: Page): Promise<void> {
  await page.goto('/?fsa=0');
  await page.getByTestId('file-input').setInputFiles({
    name: 'filter-table.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: Buffer.from(cityTableXlsx()),
  });
  await expect(page.getByTestId('file-name')).toHaveText('filter-table.xlsx');
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

/** Turn AutoFilter on from the Data tab, the path the ribbon offers. */
async function enableFilter(page: Page): Promise<void> {
  await page.getByRole('tab', { name: 'Data' }).click();
  const filter = page.getByRole('button', { name: 'Filter', exact: true });
  await filter.click();
  await expect(filter).toHaveAttribute('aria-pressed', 'true');
}

/** Alt+Down opens the dropdown for the active cell's column (§3.8). */
async function openDropdown(page: Page) {
  await page.keyboard.press('Alt+ArrowDown');
  const dropdown = page.getByRole('dialog', { name: 'Filter City' });
  await expect(dropdown).toBeVisible();
  return dropdown;
}

async function columnValues(page: Page, column: number, rows: number): Promise<(string | null)[]> {
  return page.evaluate(
    ([targetColumn, count]) =>
      Array.from({ length: count as number }, (_, row) =>
        window.__calcplus.getCellDisplay(row, targetColumn as number),
      ),
    [column, rows] as const,
  );
}

async function hiddenRows(page: Page): Promise<number[]> {
  return page.evaluate(() => window.__calcplus.getHiddenRows());
}

async function setupFilteredCity(page: Page): Promise<void> {
  await openCityTable(page);
  await selectRange(page, 'A1:B6');
  await enableFilter(page);
}

test('the dropdown opens on Alt+Down and applies a value checklist as one undo step', async ({
  page,
}) => {
  await setupFilteredCity(page);
  const versionFilterOn = await page.evaluate(() => window.__calcplus.getState().version);

  const dropdown = await openDropdown(page);
  // The dropdown owns the keyboard as soon as it opens, even though the ribbon
  // button that toggled the filter had focus.
  await expect(dropdown.getByRole('menuitem', { name: 'Sort A to Z' })).toBeFocused();

  // The capture's entries, in order, with the ruled deltas applied.
  await expect(dropdown.getByRole('menuitem', { name: 'Sort A to Z' })).toBeVisible();
  await expect(dropdown.getByRole('menuitem', { name: 'Sort Z to A' })).toBeVisible();
  await expect(dropdown.getByRole('menuitem', { name: 'Customised Sort' })).toBeVisible();
  await expect(dropdown.getByRole('menuitem', { name: "Clear Filter from 'City'" })).toBeDisabled();
  await expect(dropdown.getByRole('button', { name: 'Text Filters' })).toBeVisible();
  await expect(dropdown.getByLabel('Search values')).toBeVisible();
  await expect(dropdown.getByRole('checkbox', { name: 'Select All' })).toBeChecked();
  for (const city of ['Bath', 'Bristol', 'Cork', 'Leeds', 'York']) {
    await expect(dropdown.getByRole('checkbox', { name: city, exact: true })).toBeChecked();
  }
  await expect(dropdown.getByRole('button', { name: 'Apply' })).toBeVisible();
  // Out-of-scope entries are hidden, not greyed (ticket 28 D2).
  await expect(dropdown.getByText('Sort By Colour')).toHaveCount(0);
  await expect(dropdown.getByText('Filter By Colour')).toHaveCount(0);
  await expect(dropdown.getByText('Sheet View')).toHaveCount(0);

  // Searching is view-only: the checklist narrows and no transaction is made.
  await dropdown.getByLabel('Search values').fill('Br');
  await expect(dropdown.getByRole('checkbox', { name: 'Bristol' })).toBeVisible();
  await expect(dropdown.getByRole('checkbox', { name: 'Bath' })).toHaveCount(0);
  expect(await page.evaluate(() => window.__calcplus.getState().version)).toBe(versionFilterOn);

  await dropdown.getByLabel('Search values').fill('');
  await dropdown.getByRole('checkbox', { name: 'Leeds' }).uncheck();
  await dropdown.getByRole('checkbox', { name: 'York' }).uncheck();
  await dropdown.getByRole('button', { name: 'Apply' }).click();
  await expect(dropdown).toBeHidden();

  await expect(page.getByTestId('filter-records')).toHaveText('3 of 5 records');
  expect(await hiddenRows(page)).toEqual([3, 4]);
  await expect(page.getByTestId('status-message')).toContainText('Filter applied');

  // One filter change is one undo step, and undo restores the rows.
  await page.keyboard.press('Control+z');
  await expect.poll(() => hiddenRows(page)).toEqual([]);
  await expect(page.getByTestId('filter-records')).toHaveText('5 of 5 records');
});

test('Esc cancels the dropdown without a transaction', async ({ page }) => {
  await setupFilteredCity(page);
  const versionBefore = await page.evaluate(() => window.__calcplus.getState().version);
  const dropdown = await openDropdown(page);

  await dropdown.getByRole('checkbox', { name: 'Bristol' }).uncheck();
  await expect(dropdown.getByRole('checkbox', { name: 'Select All' })).not.toBeChecked();

  await page.keyboard.press('Escape');
  await expect(dropdown).toBeHidden();
  await expect(page.getByTestId('grid-shell')).toBeFocused();
  // Nothing was dispatched: the version, the rows and the checklist are untouched.
  expect(await page.evaluate(() => window.__calcplus.getState().version)).toBe(versionBefore);
  expect(await hiddenRows(page)).toEqual([]);
  await expect(page.getByTestId('filter-records')).toHaveText('5 of 5 records');

  // The cancelled ticks do not survive a reopen.
  await openDropdown(page);
  await expect(page.getByRole('checkbox', { name: 'Bristol' })).toBeChecked();
});

test("Clear Filter from '<column>' clears only that column", async ({ page }) => {
  await setupFilteredCity(page);
  let dropdown = await openDropdown(page);
  await dropdown.getByRole('checkbox', { name: 'Cork' }).uncheck();
  await dropdown.getByRole('button', { name: 'Apply' }).click();
  await expect(page.getByTestId('filter-records')).toHaveText('4 of 5 records');

  dropdown = await openDropdown(page);
  const clear = dropdown.getByRole('menuitem', { name: "Clear Filter from 'City'" });
  await expect(clear).toBeEnabled();
  await clear.click();
  await expect(dropdown).toBeHidden();
  await expect(page.getByTestId('filter-records')).toHaveText('5 of 5 records');
  expect(await hiddenRows(page)).toEqual([]);

  // With no criterion the entry is greyed, as the capture shows.
  dropdown = await openDropdown(page);
  await expect(dropdown.getByRole('menuitem', { name: "Clear Filter from 'City'" })).toBeDisabled();
});

test('Sort A to Z in the dropdown sorts the whole filter range', async ({ page }) => {
  await setupFilteredCity(page);
  const dropdown = await openDropdown(page);
  await dropdown.getByRole('menuitem', { name: 'Sort A to Z' }).click();
  await expect(dropdown).toBeHidden();

  await expect
    .poll(() => columnValues(page, 0, 6))
    .toEqual(['City', 'Bath', 'Bristol', 'Cork', 'Leeds', 'York']);
  // The filter is unchanged: sorting reorders visible rows only (§3.4).
  await expect(page.getByTestId('filter-records')).toHaveText('5 of 5 records');

  await page.keyboard.press('Control+z');
  await expect
    .poll(() => columnValues(page, 0, 6))
    .toEqual(['City', 'Bristol', 'Bath', 'Leeds', 'York', 'Cork']);
});

test('Customised Sort in the dropdown opens the Sort dialog for the filter range', async ({
  page,
}) => {
  await setupFilteredCity(page);
  const dropdown = await openDropdown(page);
  await dropdown.getByRole('menuitem', { name: 'Customised Sort' }).click();
  await expect(dropdown).toBeHidden();

  const dialog = page.getByRole('dialog', { name: 'Customised Sort' });
  await expect(dialog).toBeVisible();
  await expect(dialog.getByTestId('sort-target')).toContainText('A1:B6');
  await page.keyboard.press('Escape');
  await expect(dialog).toBeHidden();
});

test('Ctrl+Shift+L toggles AutoFilter on the selection', async ({ page }) => {
  await openCityTable(page);
  await selectRange(page, 'A1:B6');
  await page.keyboard.press('Control+Shift+L');
  await expect(page.getByTestId('filter-records')).toHaveText('5 of 5 records');
  await page.keyboard.press('Control+Shift+L');
  await expect(page.getByTestId('filter-records')).toHaveCount(0);
});

test('Text Filters opens the Custom AutoFilter window and applies two fields', async ({ page }) => {
  await setupFilteredCity(page);
  const dropdown = await openDropdown(page);

  await dropdown.getByRole('button', { name: 'Text Filters' }).click();
  await expect(dropdown.getByTestId('filter-type-options')).toBeVisible();
  await dropdown.getByTestId('filter-type-contains').click();

  const conditions = page.getByRole('dialog', { name: 'Custom AutoFilter' });
  await expect(conditions).toBeVisible();
  await expect(conditions.getByLabel('Operator 1')).toHaveValue('contains');
  await conditions.getByLabel('Value 1').fill('is');
  await conditions.getByRole('button', { name: 'OK' }).click();

  await expect(conditions).toBeHidden();
  await expect(dropdown).toBeHidden();
  await expect(page.getByTestId('filter-records')).toHaveText('1 of 5 records');
  expect(await hiddenRows(page)).toEqual([2, 3, 4, 5]);

  // Reopening prefills the window from the applied criterion.
  const reopened = await openDropdown(page);
  await reopened.getByRole('button', { name: 'Text Filters' }).click();
  await reopened.getByTestId('filter-type-contains').click();
  const prefilled = page.getByRole('dialog', { name: 'Custom AutoFilter' });
  await expect(prefilled.getByLabel('Value 1')).toHaveValue('is');
});
