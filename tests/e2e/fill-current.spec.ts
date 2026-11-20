import { Buffer } from 'node:buffer';
import { expect, test } from '@playwright/test';
import type { Locator, Page } from '@playwright/test';
import { formattedSheetXlsx } from '../helpers/formatted-sheet';
import { clickCell } from './grid-helpers';

/**
 * The ribbon Fill drop-down's Current swatch previews the active cell's real
 * fill (ticket 40): the selection's anchor (ticket 31), or the merge anchor
 * when the active cell sits inside a merged range. New keeps showing the
 * colour about to be applied.
 *
 * The formatted fixture gives a solid navy fill on the merged A1:D1 title and
 * a blank C3; the rest of the flow applies fills through the picker itself.
 */
const FIXTURE_NAME = 'm4a-formatted.xlsx';
const NAVY = 'rgb(31, 78, 120)';
const AMBER = 'rgb(255, 192, 0)';

async function openFormattedSheet(page: Page): Promise<void> {
  await page.goto('/?fsa=0');
  await page.getByTestId('file-input').setInputFiles({
    name: FIXTURE_NAME,
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: Buffer.from(formattedSheetXlsx()),
  });
  await page.waitForFunction(
    () => window.__calcplus.getCellDisplay(0, 0) === 'Quarterly report - formatted',
  );
}

async function openFillPicker(page: Page): Promise<Locator> {
  await page.getByRole('button', { name: 'Fill Colour options' }).click();
  const picker = page.getByRole('dialog', { name: 'Fill Colour' });
  await expect(picker).toBeVisible();
  return picker;
}

test('the Fill drop-down previews the active cell fill in Current', async ({ page }) => {
  await openFormattedSheet(page);
  await clickCell(page, 2, 2);

  let picker = await openFillPicker(page);
  let current = picker.getByTestId('fill-color-current');
  // C3 has no fill: the slashed No Fill swatch, named for assistive tech.
  await expect(current).toHaveAttribute('data-empty', 'true');
  await expect(current).toHaveAttribute('aria-label', 'Current colour No Fill');
  await expect(picker.getByTestId('fill-color-new')).toHaveAttribute('data-empty', 'true');

  // Apply a fill; the drop-down closes, and the readout follows the model.
  await picker.getByRole('button', { name: 'Colour #FFC000' }).click();
  await expect(picker).toBeHidden();
  await expect
    .poll(async () => page.evaluate(() => window.__calcplus.getCellStyle(2, 2)?.fill.color ?? null))
    .toBe('#FFC000');

  picker = await openFillPicker(page);
  current = picker.getByTestId('fill-color-current');
  await expect(current).not.toHaveAttribute('data-empty', 'true');
  await expect(current).toHaveCSS('background-color', AMBER);
  await expect(current).toHaveAttribute('aria-label', 'Current colour #FFC000');
  // New keeps showing the colour about to be applied, now the current fill.
  await expect(picker.getByTestId('fill-color-new')).toHaveCSS('background-color', AMBER);
  await page.keyboard.press('Escape');

  // B1 sits inside the merged A1:D1 title. The merged cell is one cell, so the
  // active cell collapses to its anchor and Current previews the anchor fill.
  await clickCell(page, 0, 1);
  await expect
    .poll(async () => page.evaluate(() => window.__calcplus.getState().active))
    .toEqual({ row: 0, col: 0 });

  picker = await openFillPicker(page);
  current = picker.getByTestId('fill-color-current');
  await expect(current).not.toHaveAttribute('data-empty', 'true');
  await expect(current).toHaveCSS('background-color', NAVY);
  await expect(current).toHaveAttribute('aria-label', 'Current colour #1F4E78');
  // The other style readouts read the same merged cell's anchor.
  await expect(page.getByRole('combobox', { name: 'Font size' })).toHaveValue('12');
  await expect(page.getByRole('combobox', { name: 'Number Format' })).toHaveValue('General');

  // No Fill over the merge clears the preview, again without a reload.
  await picker.getByRole('button', { name: 'No Fill' }).click();
  await expect(picker).toBeHidden();
  picker = await openFillPicker(page);
  current = picker.getByTestId('fill-color-current');
  await expect(current).toHaveAttribute('data-empty', 'true');
  await expect(current).toHaveAttribute('aria-label', 'Current colour No Fill');
  await page.keyboard.press('Escape');
});
