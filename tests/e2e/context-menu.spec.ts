import { expect, test } from '@playwright/test';
import { openFixture } from './grid-helpers';

const fixture = 'm1-chart-unknown.xlsx';

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
