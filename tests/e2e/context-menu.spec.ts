import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';
import { cellRect, clickCell, openFixture } from './grid-helpers';

const fixture = 'm1-chart-unknown.xlsx';

/** Right-click the centre of a cell without moving the selection first. */
async function rightClickCell(page: Page, row: number, col: number): Promise<void> {
  const rect = await cellRect(page, row, col);
  await page.getByTestId('grid-canvas').click({
    button: 'right',
    position: { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 },
  });
  await expect(page.getByRole('menu', { name: 'Context menu' })).toBeVisible();
}

/**
 * Right-click the row-number strip beside a row. For column 0 a cell's x is the
 * header width (the body starts at the header edge), so half of it is the
 * strip's centre at any zoom.
 */
async function rightClickRowHeader(page: Page, row: number): Promise<void> {
  const rect = await cellRect(page, row, 0);
  await page.getByTestId('grid-canvas').click({
    button: 'right',
    position: { x: rect.x / 2, y: rect.y + rect.height / 2 },
  });
  await expect(page.getByRole('menu', { name: 'Context menu' })).toBeVisible();
}

/** Right-click the column-letter strip above a column. */
async function rightClickColumnHeader(page: Page, col: number): Promise<void> {
  const rect = await cellRect(page, 0, col);
  const firstRow = await cellRect(page, 0, 0);
  await page.getByTestId('grid-canvas').click({
    button: 'right',
    position: { x: rect.x + rect.width / 2, y: firstRow.y / 2 },
  });
  await expect(page.getByRole('menu', { name: 'Context menu' })).toBeVisible();
}

test('context menu stays inside the viewport near the bottom edge', async ({ page }) => {
  await openFixture(page, fixture);

  // Right-click low on the grid: without viewport clamping the menu's last
  // items render below the window and cannot be clicked (the regression that
  // made the context-menu "Go to" item unreachable on short windows).
  const canvas = page.getByTestId('grid-canvas');
  const canvasBox = await canvas.boundingBox();
  expect(canvasBox).not.toBeNull();
  await canvas.click({
    button: 'right',
    position: { x: canvasBox!.width / 2, y: canvasBox!.height - 24 },
  });

  const menu = page.getByRole('menu', { name: 'Context menu' });
  await expect(menu).toBeVisible();
  const menuBox = await menu.boundingBox();
  const viewport = page.viewportSize();
  expect(viewport).not.toBeNull();
  expect(menuBox).not.toBeNull();
  expect(menuBox!.y).toBeGreaterThanOrEqual(0);
  expect(menuBox!.y + menuBox!.height).toBeLessThanOrEqual(viewport!.height);
  expect(menuBox!.x).toBeGreaterThanOrEqual(0);
  expect(menuBox!.x + menuBox!.width).toBeLessThanOrEqual(viewport!.width);

  // The last item is genuinely clickable, not just on screen in theory.
  await page.getByRole('menuitem', { name: 'Go to' }).click();
  await expect(page.getByRole('dialog', { name: 'Go to' })).toBeVisible();
});

test('right click inside a selected range keeps it for the context menu', async ({ page }) => {
  await openFixture(page, fixture);

  // Select a range the ticket-29 way: through the name box.
  const nameBox = page.getByTestId('name-box');
  await nameBox.click();
  await page.keyboard.press('Control+a');
  await page.keyboard.type('A1:E10');
  await page.keyboard.press('Enter');
  await expect(nameBox).toHaveValue('A1:E10');

  // Right click well inside the range: the selection must survive.
  const canvas = page.getByTestId('grid-canvas');
  await canvas.click({ button: 'right', position: { x: 100, y: 60 } });
  await expect(nameBox).toHaveValue('A1:E10');

  // The menu command must see the whole range: the status message names it.
  await page.getByRole('menuitem', { name: 'Clear contents' }).click();
  await expect(page.getByRole('contentinfo').getByText(/A1:E10/)).toBeVisible();
  await expect(nameBox).toHaveValue('A1:E10');
});

test('right click outside the selection moves it', async ({ page }) => {
  await openFixture(page, fixture);

  const nameBox = page.getByTestId('name-box');
  await nameBox.click();
  await page.keyboard.press('Control+a');
  await page.keyboard.type('A1:E10');
  await page.keyboard.press('Enter');
  await expect(nameBox).toHaveValue('A1:E10');

  const canvas = page.getByTestId('grid-canvas');
  const box = await canvas.boundingBox();
  expect(box).not.toBeNull();
  // Far bottom-right of the canvas: outside A1:E10 whatever the exact geometry.
  await canvas.click({
    button: 'right',
    position: { x: box!.width * 0.75, y: box!.height * 0.7 },
  });
  // A single cell far from the old range replaces it.
  await expect(nameBox).toHaveValue(/^[A-Z]+[0-9]+$/);
  await expect(nameBox).not.toHaveValue('A1:E10');
});

test('Insert ▸ Insert Rows Above inserts a row as one undo step', async ({ page }) => {
  await openFixture(page, fixture);
  await clickCell(page, 1, 1); // B2 holds 42.5 in the fixture
  await rightClickCell(page, 1, 1);

  await page.getByRole('menuitem', { name: 'Insert', exact: true }).click();
  await expect(page.getByRole('menu', { name: 'Insert' })).toBeVisible();

  const versionBefore = await page.evaluate(() => window.__calcplus.getState().version);
  await page.getByRole('menuitem', { name: 'Insert Rows Above', exact: true }).click();
  await expect(page.getByRole('menu', { name: 'Context menu' })).toBeHidden();

  // The row moved down: B3 now carries the old B2 value.
  await expect.poll(() => page.evaluate(() => window.__calcplus.getCellDisplay(2, 1))).toBe('42.5');
  expect(await page.evaluate(() => window.__calcplus.getState().version)).toBeGreaterThan(
    versionBefore,
  );
  expect(await page.evaluate(() => window.__calcplus.getState().canUndo)).toBe(true);

  // One undo restores the whole structural edit, not just one part of it.
  await page.getByTestId('undo-button').click();
  await expect.poll(() => page.evaluate(() => window.__calcplus.getCellDisplay(1, 1))).toBe('42.5');
  await expect.poll(() => page.evaluate(() => window.__calcplus.getCellDisplay(2, 1))).toBe('85');
});

test('Delete ▸ Delete Columns deletes and undoes', async ({ page }) => {
  await openFixture(page, fixture);
  await clickCell(page, 1, 1);
  await rightClickCell(page, 1, 1);

  await page.getByRole('menuitem', { name: 'Delete', exact: true }).click();
  await expect(page.getByRole('menu', { name: 'Delete' })).toBeVisible();
  await page.getByRole('menuitem', { name: 'Delete Columns', exact: true }).click();
  await expect(page.getByRole('menu', { name: 'Context menu' })).toBeHidden();

  // The old C2 ("inline text") shifts left into B2.
  await expect
    .poll(() => page.evaluate(() => window.__calcplus.getCellDisplay(1, 1)))
    .toBe('inline text');

  await page.getByTestId('undo-button').click();
  await expect.poll(() => page.evaluate(() => window.__calcplus.getCellDisplay(1, 1))).toBe('42.5');
});

test('Esc closes the flyout first, then the menu', async ({ page }) => {
  await openFixture(page, fixture);
  await rightClickCell(page, 1, 1);

  await page.getByRole('menuitem', { name: 'Insert', exact: true }).click();
  await expect(page.getByRole('menu', { name: 'Insert' })).toBeVisible();

  await page.keyboard.press('Escape');
  await expect(page.getByRole('menu', { name: 'Insert' })).toBeHidden();
  await expect(page.getByRole('menu', { name: 'Context menu' })).toBeVisible();

  await page.keyboard.press('Escape');
  await expect(page.getByRole('menu', { name: 'Context menu' })).toBeHidden();
});

test('Right arrow opens the flyout and Enter applies its item', async ({ page }) => {
  await openFixture(page, fixture);
  await clickCell(page, 1, 1);
  await rightClickCell(page, 1, 1);

  // The menu focuses Cut; four steps down reaches the Insert flyout entry.
  const insertItem = page.getByRole('menuitem', { name: 'Insert', exact: true });
  for (let step = 0; step < 4; step += 1) await page.keyboard.press('ArrowDown');
  await expect(insertItem).toBeFocused();
  await expect(insertItem).toHaveAttribute('aria-expanded', 'false');

  await page.keyboard.press('ArrowRight');
  await expect(insertItem).toHaveAttribute('aria-expanded', 'true');
  const firstChild = page.getByRole('menuitem', { name: 'Insert Rows Above', exact: true });
  await expect(firstChild).toBeFocused();

  const versionBefore = await page.evaluate(() => window.__calcplus.getState().version);
  await page.keyboard.press('Enter');
  await expect.poll(() => page.evaluate(() => window.__calcplus.getCellDisplay(2, 1))).toBe('42.5');
  expect(await page.evaluate(() => window.__calcplus.getState().version)).toBeGreaterThan(
    versionBefore,
  );
});

test('row header offers flat row items and inserts one row in one undo step', async ({ page }) => {
  await openFixture(page, fixture);
  await clickCell(page, 1, 1); // B2 holds 42.5; the row header picks row 2
  await rightClickRowHeader(page, 1);

  // Excel's row-header menu: flat row actions, with no column or cell shifts.
  const insertRows = page.getByRole('menuitem', { name: 'Insert Rows Above', exact: true });
  const deleteRows = page.getByRole('menuitem', { name: 'Delete Rows', exact: true });
  await expect(insertRows).toBeVisible();
  await expect(deleteRows).toBeVisible();
  await expect(page.getByRole('menuitem', { name: 'Insert', exact: true })).toHaveCount(0);
  await expect(page.getByRole('menuitem', { name: 'Delete', exact: true })).toHaveCount(0);
  await expect(
    page.getByRole('menuitem', { name: 'Insert Columns to the Left', exact: true }),
  ).toHaveCount(0);
  await expect(page.getByRole('menuitem', { name: 'Delete Columns', exact: true })).toHaveCount(0);
  await expect(
    page.getByRole('menuitem', { name: 'Insert Cells & Shift Right', exact: true }),
  ).toHaveCount(0);

  const versionBefore = await page.evaluate(() => window.__calcplus.getState().version);
  await insertRows.click();
  await expect(page.getByRole('menu', { name: 'Context menu' })).toBeHidden();

  // The row moved down: B3 now carries the old B2 value.
  await expect.poll(() => page.evaluate(() => window.__calcplus.getCellDisplay(2, 1))).toBe('42.5');
  expect(await page.evaluate(() => window.__calcplus.getState().version)).toBeGreaterThan(
    versionBefore,
  );
  expect(await page.evaluate(() => window.__calcplus.getState().canUndo)).toBe(true);

  // One undo restores the whole structural edit, not just one part of it.
  await page.getByTestId('undo-button').click();
  await expect.poll(() => page.evaluate(() => window.__calcplus.getCellDisplay(1, 1))).toBe('42.5');
  await expect.poll(() => page.evaluate(() => window.__calcplus.getCellDisplay(2, 1))).toBe('85');
});

test('column header offers flat column items and deletes one column', async ({ page }) => {
  await openFixture(page, fixture);
  await clickCell(page, 1, 1); // B2 holds 42.5; the column header picks column B
  await rightClickColumnHeader(page, 1);

  const insertColumns = page.getByRole('menuitem', {
    name: 'Insert Columns to the Left',
    exact: true,
  });
  const deleteColumns = page.getByRole('menuitem', { name: 'Delete Columns', exact: true });
  await expect(insertColumns).toBeVisible();
  await expect(deleteColumns).toBeVisible();
  await expect(page.getByRole('menuitem', { name: 'Insert', exact: true })).toHaveCount(0);
  await expect(page.getByRole('menuitem', { name: 'Delete', exact: true })).toHaveCount(0);
  await expect(page.getByRole('menuitem', { name: 'Insert Rows Above', exact: true })).toHaveCount(
    0,
  );
  await expect(page.getByRole('menuitem', { name: 'Delete Rows', exact: true })).toHaveCount(0);

  await deleteColumns.click();
  await expect(page.getByRole('menu', { name: 'Context menu' })).toBeHidden();

  // The old C2 ("inline text") shifts left into B2.
  await expect
    .poll(() => page.evaluate(() => window.__calcplus.getCellDisplay(1, 1)))
    .toBe('inline text');

  await page.getByTestId('undo-button').click();
  await expect.poll(() => page.evaluate(() => window.__calcplus.getCellDisplay(1, 1))).toBe('42.5');
});
