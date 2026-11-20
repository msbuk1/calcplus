import { Buffer } from 'node:buffer';
import { expect, test } from '@playwright/test';
import type { Locator, Page } from '@playwright/test';
import {
  cityTableXlsx,
  deepHeaderTableXlsx,
  longValueTableXlsx,
  manyValuesTableXlsx,
  wideTableXlsx,
} from './data-tools-fixture';
import { cellRect } from './grid-helpers';

/**
 * The AutoFilter dropdown (sort-autofilter §3.8; design-refs
 * m5-menus-dialogs §2). The capture's ruled deltas from ticket 28 are asserted
 * here too: the Apply button, the "Customised Sort" entry, the
 * "Clear Filter from '<column>'" wording and the absence of the greyed
 * out-of-scope entries.
 */

async function openTable(page: Page, name: string, bytes: Uint8Array): Promise<void> {
  await page.goto('/?fsa=0');
  await page.getByTestId('file-input').setInputFiles({
    name,
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: Buffer.from(bytes),
  });
  await expect(page.getByTestId('file-name')).toHaveText(name);
  await expect(page.getByTestId('grid-canvas')).toBeVisible();
}

async function openCityTable(page: Page): Promise<void> {
  await openTable(page, 'filter-table.xlsx', cityTableXlsx());
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
async function openDropdown(page: Page, columnLabel = 'City') {
  await page.keyboard.press('Alt+ArrowDown');
  const dropdown = page.getByRole('dialog', { name: `Filter ${columnLabel}` });
  await expect(dropdown).toBeVisible();
  return dropdown;
}

/** Every edge of the panel, including its footer, is inside the window. */
async function expectInsideViewport(page: Page, panel: Locator): Promise<void> {
  const box = await panel.boundingBox();
  const viewport = page.viewportSize();
  expect(box).not.toBeNull();
  expect(viewport).not.toBeNull();
  expect(box!.x).toBeGreaterThanOrEqual(0);
  expect(box!.y).toBeGreaterThanOrEqual(0);
  expect(box!.x + box!.width).toBeLessThanOrEqual(viewport!.width);
  expect(box!.y + box!.height).toBeLessThanOrEqual(viewport!.height);
}

/** The filter header's top edge in document coordinates. */
async function headerTop(page: Page, headerRow: number): Promise<number> {
  const canvas = page.getByTestId('grid-canvas');
  const canvasBox = await canvas.boundingBox();
  expect(canvasBox).not.toBeNull();
  const rect = await cellRect(page, headerRow, 0);
  return canvasBox!.y + rect.y;
}

/**
 * Put a deep filter header in the lower part of the window. Selecting the
 * range already scrolls near it; a small wheel adjusts the rest.
 */
async function moveHeaderNearBottom(page: Page, headerRow: number): Promise<void> {
  const canvas = page.getByTestId('grid-canvas');
  const canvasBox = await canvas.boundingBox();
  expect(canvasBox).not.toBeNull();
  const current = await headerTop(page, headerRow);
  const target = 560;
  if (current < target - 8) {
    await page.mouse.move(
      canvasBox!.x + canvasBox!.width / 2,
      canvasBox!.y + canvasBox!.height / 2,
    );
    await page.mouse.wheel(0, current - target);
  }
  await expect.poll(() => headerTop(page, headerRow)).toBeGreaterThan(400);
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

test('filter-hidden rows collapse out of the geometry (sort-autofilter §3.3)', async ({ page }) => {
  await openCityTable(page);
  // Row geometry with every row visible: header, five city rows.
  const before = [];
  for (let row = 0; row < 6; row += 1) before.push(await cellRect(page, row, 0));

  await selectRange(page, 'A1:B6');
  await enableFilter(page);
  const dropdown = await openDropdown(page);
  await dropdown.getByRole('checkbox', { name: 'Leeds' }).uncheck();
  await dropdown.getByRole('checkbox', { name: 'York' }).uncheck();
  await dropdown.getByRole('button', { name: 'Apply' }).click();
  await expect(dropdown).toBeHidden();
  await expect(page.getByTestId('filter-records')).toHaveText('3 of 5 records');

  // Leeds and York collapse to zero height, so Bath and Cork sit together.
  const bath = await cellRect(page, 2, 0);
  const cork = await cellRect(page, 5, 0);
  expect(bath.height).toBe(before[2].height);
  expect(cork.height).toBe(before[5].height);
  expect(cork.y).toBe(bath.y + bath.height);
  for (const row of [3, 4]) {
    const hidden = await cellRect(page, row, 0);
    expect(hidden.height).toBe(0);
    expect(hidden.y).toBe(cork.y);
  }

  // A press where a filtered row used to paint hits the visible row that moved
  // there, never the hidden row itself.
  const collapsed = await cellRect(page, 3, 0);
  const canvas = await page.getByTestId('grid-canvas').boundingBox();
  expect(canvas).not.toBeNull();
  await page.mouse.click(
    canvas!.x + collapsed.x + collapsed.width / 2,
    canvas!.y + collapsed.y + 4,
  );
  await expect(page.getByTestId('active-cell')).toHaveText('A6');

  // Clearing the filter restores every row height exactly; the manual layer is
  // untouched and the visible rows space out again.
  await openDropdown(page);
  await page.getByRole('menuitem', { name: "Clear Filter from 'City'" }).click();
  await expect(page.getByTestId('filter-records')).toHaveText('5 of 5 records');
  for (let row = 0; row < 6; row += 1) {
    expect(await cellRect(page, row, 0)).toEqual(before[row]);
  }
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

test('a long value stays on one row and truncates with an ellipsis', async ({ page }) => {
  const longValue =
    'This deliberately long filter value must truncate at the dropdown edge rather than wrap character by character';
  await openTable(page, 'long-filter.xlsx', longValueTableXlsx(longValue));
  await selectRange(page, 'A1:A3');
  await enableFilter(page);
  await selectRange(page, 'A2');

  const dropdown = await openDropdown(page, 'Value');
  const row = dropdown.locator('.filter-check').filter({ hasText: 'deliberately long' });
  await expect(row).toBeVisible();

  const layout = await row.evaluate((label) => {
    const text = label.querySelector<HTMLElement>('.filter-check-text');
    if (text === null) throw new Error('The value text is not wrapped in its own element.');
    const style = getComputedStyle(text);
    return {
      labelHeight: label.getBoundingClientRect().height,
      lineRects: text.getClientRects().length,
      textOverflow: style.textOverflow,
      whiteSpace: style.whiteSpace,
      minWidth: style.minWidth,
      scrollWidth: text.scrollWidth,
      clientWidth: text.clientWidth,
    };
  });

  // A squashed flex item used to collapse this row to a few pixels and draw
  // the text over its neighbours. One line, clipped, is the Excel shape.
  expect(layout.labelHeight).toBeGreaterThanOrEqual(20);
  expect(layout.lineRects).toBe(1);
  expect(layout.textOverflow).toBe('ellipsis');
  expect(layout.whiteSpace).toBe('nowrap');
  expect(layout.minWidth).toBe('0px');
  expect(layout.scrollWidth).toBeGreaterThan(layout.clientWidth);
});

test('a column with many distinct values keeps the dropdown usable and Select All covers them all', async ({
  page,
}) => {
  const valueCount = 600;
  await openTable(page, 'many-filter.xlsx', manyValuesTableXlsx(valueCount));
  await selectRange(page, `A1:A${valueCount + 1}`);
  await enableFilter(page);
  await selectRange(page, 'A2');

  const dropdown = await openDropdown(page, 'Value');
  await expect(dropdown.getByRole('checkbox', { name: 'Value 0001' })).toBeVisible();
  // The model offers every value, but the list renders only its viewport
  // window, not one DOM row per value.
  expect(await dropdown.locator('.filter-checklist .filter-check').count()).toBeLessThan(60);

  // The search box narrows the list, and clearing it restores the full list at
  // the top.
  await dropdown.getByLabel('Search values').fill('Value 0600');
  await expect(dropdown.getByRole('checkbox', { name: 'Value 0600' })).toBeVisible();
  await expect(dropdown.getByRole('checkbox', { name: 'Value 0001' })).toHaveCount(0);
  await dropdown.getByLabel('Search values').fill('');
  await expect(dropdown.getByRole('checkbox', { name: 'Value 0001' })).toBeVisible();

  // The last value is reachable, so no value was silently capped away.
  await dropdown.getByTestId('filter-checklist').evaluate((list) => {
    list.scrollTop = list.scrollHeight;
  });
  await expect(dropdown.getByRole('checkbox', { name: 'Value 0600' })).toBeVisible();

  // Select All ticks the full model list while most values are outside the
  // DOM. Unchecking one row then leaves every other row visible.
  await dropdown.getByTestId('filter-select-all').uncheck();
  await dropdown.getByTestId('filter-select-all').check();
  await dropdown.getByRole('checkbox', { name: 'Value 0600' }).uncheck();
  await dropdown.getByRole('button', { name: 'Apply' }).click();
  await expect(dropdown).toBeHidden();
  await expect(page.getByTestId('filter-records')).toHaveText('599 of 600 records');
});

test('(Blanks) still appears when the column has blank rows', async ({ page }) => {
  await openCityTable(page);
  // A1:B7 reaches one row past the city table, so the Sales column has a blank.
  await selectRange(page, 'A1:B7');
  await enableFilter(page);
  await selectRange(page, 'B2');

  const dropdown = await openDropdown(page, 'Sales');
  const blanks = dropdown.getByRole('checkbox', { name: '(Blanks)' });
  await expect(blanks).toBeVisible();
  await expect(blanks).toBeChecked();
  // It is the last checklist row, after the values.
  await expect(dropdown.locator('.filter-check').last()).toContainText('(Blanks)');
});

test('the dropdown stays on screen when it opens near the top', async ({ page }) => {
  await openCityTable(page);
  await selectRange(page, 'A1:B6');
  await enableFilter(page);
  await selectRange(page, 'A2');

  const dropdown = await openDropdown(page);
  await expectInsideViewport(page, dropdown);
  await expect(dropdown).toHaveAttribute('data-flipped', 'false');
  await expect(dropdown.getByRole('button', { name: 'Apply' })).toBeVisible();
});

test('the dropdown flips above a header near the bottom and Apply still works', async ({
  page,
}) => {
  await openTable(page, 'deep-filter.xlsx', deepHeaderTableXlsx(100));
  await selectRange(page, 'A100:A105');
  await enableFilter(page);
  await moveHeaderNearBottom(page, 99);
  const header = await headerTop(page, 99);

  const dropdown = await openDropdown(page, 'Value');
  await expectInsideViewport(page, dropdown);
  await expect(dropdown).toHaveAttribute('data-flipped', 'true');
  const box = await dropdown.boundingBox();
  expect(box).not.toBeNull();
  expect(box!.y + box!.height).toBeLessThanOrEqual(header + 1);

  // The footer is the part that must never leave the window.
  const apply = dropdown.getByRole('button', { name: 'Apply' });
  await expect(apply).toBeVisible();
  await dropdown.getByRole('checkbox', { name: 'Leeds' }).uncheck();
  await apply.click();
  await expect(dropdown).toBeHidden();
  await expect(page.getByTestId('filter-records')).toHaveText('4 of 5 records');
});

test('the dropdown stays on screen when it opens near the right edge', async ({ page }) => {
  await openTable(page, 'wide-filter.xlsx', wideTableXlsx());
  await selectRange(page, 'Z1:Z2');
  await enableFilter(page);

  const dropdown = await openDropdown(page, 'Column 26');
  await expectInsideViewport(page, dropdown);
  const box = await dropdown.boundingBox();
  const viewport = page.viewportSize();
  expect(box).not.toBeNull();
  expect(viewport).not.toBeNull();
  expect(box!.x + box!.width).toBeLessThanOrEqual(viewport!.width);
  await expect(dropdown.getByRole('button', { name: 'Apply' })).toBeVisible();
});

test('a window resize dismisses the dropdown rather than leaving it detached', async ({ page }) => {
  await openCityTable(page);
  await selectRange(page, 'A1:B6');
  await enableFilter(page);
  await selectRange(page, 'A2');

  const dropdown = await openDropdown(page);
  await page.setViewportSize({ width: 900, height: 400 });
  await expect(dropdown).toBeHidden();
  await expect(page.getByTestId('grid-shell')).toBeFocused();
});

test('a short window clamps the panel and the value list takes the scroll', async ({ page }) => {
  await page.setViewportSize({ width: 1280, height: 360 });
  await openTable(page, 'short-window.xlsx', manyValuesTableXlsx(100));
  await selectRange(page, 'A1:A101');
  await enableFilter(page);
  await selectRange(page, 'A2');

  const dropdown = await openDropdown(page, 'Value');
  await expectInsideViewport(page, dropdown);
  // It fits on neither side of the anchor, so it is clamped rather than flipped.
  await expect(dropdown).toHaveAttribute('data-flipped', 'false');
  const sizes = await dropdown.getByTestId('filter-checklist').evaluate((element) => ({
    clientHeight: element.clientHeight,
    scrollHeight: element.scrollHeight,
  }));
  expect(sizes.scrollHeight).toBeGreaterThan(sizes.clientHeight);
  await expect(dropdown.getByRole('button', { name: 'Apply' })).toBeVisible();
});

test('the value list scrolls while the footer stays put', async ({ page }) => {
  const valueCount = 200;
  await openTable(page, 'scroll-filter.xlsx', manyValuesTableXlsx(valueCount));
  await selectRange(page, `A1:A${valueCount + 1}`);
  await enableFilter(page);

  const dropdown = await openDropdown(page, 'Value');
  const list = dropdown.getByTestId('filter-checklist');
  const apply = dropdown.getByRole('button', { name: 'Apply' });
  const footerBefore = await apply.boundingBox();
  const scrolling = await list.evaluate((element) => ({
    clientHeight: element.clientHeight,
    scrollHeight: element.scrollHeight,
  }));
  expect(scrolling.scrollHeight).toBeGreaterThan(scrolling.clientHeight);

  await list.evaluate((element) => {
    element.scrollTop = element.scrollHeight;
  });
  await expect(dropdown.getByRole('checkbox', { name: 'Value 0200' })).toBeVisible();
  const footerAfter = await apply.boundingBox();
  expect(footerAfter).toEqual(footerBefore);
  await expectInsideViewport(page, dropdown);
});
