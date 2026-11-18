import { expect, test } from '@playwright/test';
import { openFixture } from './grid-helpers';

const fixture = 'm1-chart-unknown.xlsx';

test('context menu stays inside the viewport near the bottom edge', async ({ page }) => {
  await openFixture(page, fixture);

  // Right-click low on the grid: without viewport clamping the menu's last
  // items render below the window and cannot be clicked (the regression that
  // made the context-menu "Go To" item unreachable on short windows).
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
  await page.getByRole('menuitem', { name: 'Go To' }).click();
  await expect(page.getByRole('dialog', { name: 'Go To' })).toBeVisible();
});
