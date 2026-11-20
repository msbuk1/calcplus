import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';
import { clickCell, openFixture } from './grid-helpers';

/**
 * The Borders split button's Excel 365 submenu: fifteen items in oracle order,
 * each with original artwork, the single-edge and thick/double commands, and
 * More Borders opening Format Cells on its Border tab.
 */

const fixture = 'm2-grid.xlsx';

const MENU_LABELS = [
  'Bottom Border',
  'Top Border',
  'Left Border',
  'Right Border',
  'No Border',
  'All Borders',
  'Outside Borders',
  'Thick Outside Borders',
  'Bottom Double Border',
  'Thick Bottom Border',
  'Top and Bottom Border',
  'Top and Thick Bottom Border',
  'Top and Double Bottom Border',
  'Inside Borders',
  'More Borders...',
];

async function openBordersMenu(page: Page) {
  const split = page
    .locator('.ribbon-split')
    .filter({ has: page.locator('[data-command="format.border"]') });
  await split.locator('.ribbon-split-toggle').click();
  return page.getByRole('menu', { name: 'Borders options' });
}

test('lists all fifteen borders with icons and two dividers', async ({ page }) => {
  await openFixture(page, fixture);
  await clickCell(page, 1, 1);
  const menu = await openBordersMenu(page);

  await expect(menu.getByRole('menuitem')).toHaveText(MENU_LABELS);
  await expect(menu.locator('.menu-icon .ribbon-glyph')).toHaveCount(MENU_LABELS.length);
  await expect(menu.locator('.menu-separator')).toHaveCount(2);
});

test('a single-edge item applies only that edge', async ({ page }) => {
  await openFixture(page, fixture);
  await clickCell(page, 1, 1);
  const menu = await openBordersMenu(page);
  await menu.getByRole('menuitem', { name: 'Bottom Border', exact: true }).click();

  const border = await page.evaluate(
    () => window.__calcplus.getCellStyle(1, 1)?.border.bottom ?? null,
  );
  expect(border).toEqual({ style: 'thin', color: '#000000', weight: 'thin' });
});

test('Thick Outside Borders uses the thick style the command accepts', async ({ page }) => {
  await openFixture(page, fixture);
  await clickCell(page, 1, 1);
  const menu = await openBordersMenu(page);
  await menu.getByRole('menuitem', { name: 'Thick Outside Borders', exact: true }).click();

  const border = await page.evaluate(() => window.__calcplus.getCellStyle(1, 1)?.border ?? null);
  for (const edge of [border?.left, border?.right, border?.top, border?.bottom]) {
    expect(edge).toEqual({ style: 'thick', color: '#000000', weight: 'thick' });
  }
});

test('More Borders opens Format Cells on its Border tab', async ({ page }) => {
  await openFixture(page, fixture);
  await clickCell(page, 1, 1);
  const menu = await openBordersMenu(page);
  await menu.getByRole('menuitem', { name: 'More Borders...', exact: true }).click();

  const dialog = page.getByRole('dialog', { name: 'Format Cells' });
  await expect(dialog).toBeVisible();
  await expect(dialog.getByRole('tab', { name: 'Border' })).toHaveAttribute(
    'aria-selected',
    'true',
  );
  await expect(dialog.getByTestId('format-border-style')).toBeVisible();
});
