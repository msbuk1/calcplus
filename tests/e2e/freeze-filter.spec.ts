import { Buffer } from 'node:buffer';
import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';
import { cityTableXlsx } from './data-tools-fixture';
import { cellRect, clickCell, openFixture, pixelAt } from './grid-helpers';

/**
 * The two feature gaps the maintainer hit in real use: Freeze Panes on the
 * View tab and the AutoFilter affordance in the grid's header row. The filter
 * dropdown itself is covered by `filter-dropdown.spec.ts`; this file proves the
 * button exists, carries a name, opens the same dropdown and shows the funnel
 * state, and that both ribbon placements dispatch AutoFilter.
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

/**
 * Wait for the scheduled canvas paint, not just the DOM update: the header
 * glyph pixels below are read from the canvas, which repaints on the next
 * animation frame and can otherwise still show the previous glyph.
 */
async function settle(page: Page): Promise<void> {
  await page.evaluate(
    () =>
      new Promise<void>((resolve) => {
        requestAnimationFrame(() => requestAnimationFrame(() => resolve()));
      }),
  );
}

test('a fresh sheet offers Freeze Panes; the label follows the split', async ({ page }) => {
  await page.goto('/?fsa=0');
  await page.getByTestId('file-menu-button').click();
  await page.getByRole('menuitem', { name: /^New$/ }).click();
  await expect(page.getByTestId('file-name')).toHaveText('Untitled.xlsx');
  await page.getByRole('tab', { name: 'View', exact: true }).click();

  // No split: the entry is Freeze Panes (ticket 43).
  await expect(page.getByRole('button', { name: 'Freeze Panes', exact: true })).toBeEnabled();
  await expect(page.getByRole('button', { name: 'Unfreeze Panes', exact: true })).toHaveCount(0);

  // Freezing at B1 flips the same entry to Unfreeze Panes.
  await clickCell(page, 1, 0);
  await page.getByRole('button', { name: 'Freeze Panes', exact: true }).click();
  await expect
    .poll(() => page.evaluate(() => window.__calcplus.getFreeze()))
    .toEqual({ rows: 1, cols: 0 });
  const unfreeze = page.getByRole('button', { name: 'Unfreeze Panes', exact: true });
  await expect(unfreeze).toBeEnabled();

  // Unfreezing from the split point flips it back.
  await unfreeze.click();
  await expect
    .poll(() => page.evaluate(() => window.__calcplus.getFreeze()))
    .toEqual({ rows: 0, cols: 0 });
  await expect(page.getByRole('button', { name: 'Freeze Panes', exact: true })).toBeEnabled();
});

test('Freeze Panes splits at the active cell, moves, unfreezes and undoes', async ({ page }) => {
  await openFixture(page, 'm2-grid.xlsx');
  await clickCell(page, 2, 1);
  await page.getByRole('tab', { name: 'View', exact: true }).click();

  // m2-grid.xlsx opens with a 1 × 1 split, so the entry already reads
  // Unfreeze Panes (ticket 43) and pressing at a new cell moves the split.
  const unfreeze = page.getByRole('button', { name: 'Unfreeze Panes', exact: true });
  await expect(unfreeze).toBeEnabled();
  await expect(page.getByRole('button', { name: 'Freeze Panes', exact: true })).toHaveCount(0);
  await unfreeze.click();
  await expect
    .poll(() => page.evaluate(() => window.__calcplus.getFreeze()))
    .toEqual({
      rows: 2,
      cols: 1,
    });

  // The frozen strips stay painted while the body scrolls.
  const frozen = await page.evaluate(() => window.__calcplus.getCellRect(0, 1));
  expect(frozen).not.toBeNull();
  const frozenBefore = await pixelAt(page, frozen!.x + 8, frozen!.y + 8);
  const box = await page.getByTestId('grid-canvas').boundingBox();
  await page.mouse.move(box!.x + box!.width / 2, box!.y + box!.height / 2);
  await page.mouse.wheel(0, 500);
  await page.waitForTimeout(150);
  expect(await pixelAt(page, frozen!.x + 8, frozen!.y + 8)).toEqual(frozenBefore);

  // A new active cell moves the split; the same cell unfreezes it.
  await selectRange(page, 'C5');
  await unfreeze.click();
  await expect
    .poll(() => page.evaluate(() => window.__calcplus.getFreeze()))
    .toEqual({
      rows: 4,
      cols: 2,
    });
  await unfreeze.click();
  await expect
    .poll(() => page.evaluate(() => window.__calcplus.getFreeze()))
    .toEqual({
      rows: 0,
      cols: 0,
    });
  // With no split the entry reads Freeze Panes again.
  const freeze = page.getByRole('button', { name: 'Freeze Panes', exact: true });
  await expect(freeze).toBeEnabled();
  await expect(unfreeze).toHaveCount(0);

  // One transaction: undo restores the split that was in place before the press.
  await page.getByTestId('undo-button').click();
  await expect
    .poll(() => page.evaluate(() => window.__calcplus.getFreeze()))
    .toEqual({
      rows: 4,
      cols: 2,
    });
  await expect(unfreeze).toBeEnabled();
});

test('the AutoFilter button in a header cell opens the column dropdown', async ({ page }) => {
  await openCityTable(page);
  await selectRange(page, 'A1:B6');
  await page.getByRole('tab', { name: 'Data' }).click();
  await page.getByRole('button', { name: 'Filter', exact: true }).click();

  const button = page.getByRole('button', { name: 'Filter column A' });
  await expect(button).toBeVisible();
  await expect(button).toHaveAttribute('aria-haspopup', 'dialog');

  // The canvas paints a chevron on an unfiltered header. Its centre pixel is
  // the white button face; the funnel that replaces it is a filled glyph.
  const headerCell = await cellRect(page, 0, 0);
  const glyphX = headerCell.x + headerCell.width - 9;
  const glyphY = headerCell.y + headerCell.height / 2;
  await settle(page);
  const arrowPixel = await pixelAt(page, glyphX, glyphY);

  // Pointer: the DOM button opens the existing dropdown, not a new surface.
  await button.click();
  const dropdown = page.getByRole('dialog', { name: 'Filter City' });
  await expect(dropdown).toBeVisible();
  await page.keyboard.press('Escape');
  await expect(dropdown).toBeHidden();

  // Keyboard: the same button is focusable and Enter opens the dropdown.
  await button.focus();
  await page.keyboard.press('Enter');
  await expect(page.getByRole('dialog', { name: 'Filter City' })).toBeVisible();
  await page.keyboard.press('Escape');

  // Filtered columns expose the funnel state and its accessible name.
  await button.click();
  await page
    .getByRole('dialog', { name: 'Filter City' })
    .getByRole('checkbox', { name: 'Bristol' })
    .uncheck();
  await page
    .getByRole('dialog', { name: 'Filter City' })
    .getByRole('button', { name: 'Apply' })
    .click();
  const filtered = page.getByRole('button', { name: 'Filter column A, filtered' });
  await expect(filtered).toBeVisible();
  await expect(filtered).toHaveAttribute('data-filtered', 'true');
  await settle(page);
  const funnelPixel = await pixelAt(page, glyphX, glyphY);
  expect(funnelPixel).not.toEqual(arrowPixel);
  expect(funnelPixel.r).toBeLessThan(200);
});

test('the Home Sort & Filter menu and the Data tab both dispatch AutoFilter', async ({ page }) => {
  await openCityTable(page);
  await selectRange(page, 'A1:B6');

  // Home → Editing → Sort & Filter → Filter.
  await page.getByRole('tab', { name: 'Home' }).click();
  await page.getByRole('button', { name: 'Sort & Filter', exact: true }).click();
  await page.getByRole('menuitem', { name: 'Filter', exact: true }).click();
  await expect(page.getByTestId('filter-records')).toHaveText('5 of 5 records');

  // Clear All from the Data tab keeps the AutoFilter but restores every row.
  await page.getByRole('tab', { name: 'Data' }).click();
  await page.getByRole('button', { name: 'Filter column A' }).click();
  await page
    .getByRole('dialog', { name: 'Filter City' })
    .getByRole('checkbox', { name: 'Bristol' })
    .uncheck();
  await page
    .getByRole('dialog', { name: 'Filter City' })
    .getByRole('button', { name: 'Apply' })
    .click();
  await expect(page.getByTestId('filter-records')).toHaveText('4 of 5 records');

  const clear = page.getByRole('button', { name: 'Clear', exact: true });
  await expect(clear).toBeEnabled();
  await clear.click();
  await expect(page.getByTestId('filter-records')).toHaveText('5 of 5 records');

  // Reapply leaves an unchanged filter as the no-op transaction §3.9 names.
  const reapply = page.getByRole('button', { name: 'Reapply', exact: true });
  await expect(reapply).toBeEnabled();
  await reapply.click();
  await expect(page.getByTestId('filter-records')).toHaveText('5 of 5 records');

  // The Data tab toggle still turns the filter itself off.
  await page.getByRole('button', { name: 'Filter', exact: true }).click();
  await expect(page.getByTestId('filter-records')).toHaveCount(0);
});
