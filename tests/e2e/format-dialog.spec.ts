import { expect, test } from '@playwright/test';
import type { Locator, Page } from '@playwright/test';
import { cellRect, clickCell, openFixture, regionSignature } from './grid-helpers';

const fixture = 'm1-chart-unknown.xlsx';

type ColourField = 'format-font-color' | 'format-fill-color' | 'format-border-color';

/** The colour control opens a separate Colours popup instead of taking text directly. */
async function openColours(dialog: Locator, field: ColourField): Promise<Locator> {
  const control = dialog.getByTestId(field);
  await control.click();
  await expect(control).toHaveAttribute('aria-expanded', 'true');
  const popup = dialog.getByRole('dialog', { name: 'Colours' });
  await expect(popup).toBeVisible();
  return popup;
}

/** Pick a custom hex value and commit it from the popup without applying the dialog. */
async function pickCustomColour(dialog: Locator, field: ColourField, value: string): Promise<void> {
  const popup = await openColours(dialog, field);
  await popup.getByRole('tab', { name: 'Custom' }).click();
  await popup.getByTestId(`${field}-custom`).fill(value);
  await popup.getByRole('button', { name: 'OK' }).click();
  await expect(popup).toBeHidden();
}

async function styleAt(page: Page, row: number, col: number) {
  return page.evaluate(
    ([targetRow, targetCol]) => window.__calcplus.getCellStyle(targetRow, targetCol),
    [row, col] as const,
  );
}

async function canvasPoint(page: Page, row: number, col: number) {
  const box = await page.getByTestId('grid-canvas').boundingBox();
  if (box === null) throw new Error('The grid canvas has no box');
  const rect = await cellRect(page, row, col);
  return {
    x: box.x + rect.x + rect.width / 2,
    y: box.y + rect.y + rect.height / 2,
  };
}

test('Format Cells exposes every v1 tab and applies font, fill, border, alignment and number formats', async ({
  page,
}) => {
  await openFixture(page, fixture);
  await clickCell(page, 1, 1);

  const trigger = page.getByRole('button', { name: 'Number Settings' });
  await trigger.click();
  const dialog = page.getByRole('dialog', { name: 'Format Cells' });
  await expect(dialog).toBeVisible();
  await expect(dialog.getByRole('tab', { name: 'Font' })).toBeVisible();
  await expect(dialog.getByRole('tab', { name: 'Fill' })).toBeVisible();
  await expect(dialog.getByRole('tab', { name: 'Border' })).toBeVisible();
  await expect(dialog.getByRole('tab', { name: 'Alignment' })).toBeVisible();
  await expect(dialog.getByRole('tab', { name: 'Number' })).toBeVisible();
  // The Protection tab is V2 and must not ship.
  await expect(dialog.getByRole('tab', { name: 'Protection' })).toHaveCount(0);

  await dialog.getByLabel('Font', { exact: true }).fill('Arial');
  await dialog.getByLabel('Size').fill('14');
  await dialog.getByLabel('Bold').check();
  await dialog.getByLabel('Italic').check();
  await dialog.getByLabel('Underline').selectOption('double');
  await dialog.getByLabel('Strikethrough').check();
  await pickCustomColour(dialog, 'format-font-color', '#C00000');

  await dialog.getByRole('tab', { name: 'Fill' }).click();
  await dialog.getByLabel('Pattern', { exact: true }).selectOption('solid');
  await pickCustomColour(dialog, 'format-fill-color', '#FFF2CC');

  await dialog.getByRole('tab', { name: 'Border' }).click();
  await dialog.getByLabel('Preset').selectOption('outline');
  await dialog.getByLabel('Line style').selectOption('medium');
  await pickCustomColour(dialog, 'format-border-color', '#FF0000');

  await dialog.getByRole('tab', { name: 'Alignment' }).click();
  await dialog.getByLabel('Horizontal').selectOption('center');
  await dialog.getByLabel('Vertical').selectOption('top');
  await dialog.getByLabel('Wrap text').check();
  await dialog.getByLabel('Indent').fill('2');

  await dialog.getByRole('tab', { name: 'Number' }).click();
  await dialog.getByLabel('Category').selectOption('currency');
  await expect(dialog.getByLabel('Format code')).toHaveValue('$#,##0.00_);($#,##0.00)');

  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect(dialog).toBeHidden();

  await expect
    .poll(() => styleAt(page, 1, 1))
    .toMatchObject({
      font: {
        name: 'Arial',
        size: 14,
        bold: true,
        italic: true,
        underlineStyle: 'double',
        strike: true,
        color: '#C00000',
      },
      fill: { pattern: 'solid', color: '#FFF2CC' },
      border: {
        top: { style: 'medium', color: '#FF0000' },
        left: { style: 'medium', color: '#FF0000' },
      },
      alignment: { horizontal: 'center', vertical: 'top', wrapText: true, indent: 2 },
      numberFormat: '$#,##0.00_);($#,##0.00)',
    });
});

test('Format Cells commits every changed tab as one undo step', async ({ page }) => {
  await openFixture(page, fixture);
  await clickCell(page, 1, 1);
  const before = await styleAt(page, 1, 1);

  await page.getByRole('button', { name: 'Number Settings' }).click();
  const dialog = page.getByRole('dialog', { name: 'Format Cells' });
  await expect(dialog).toBeVisible();

  await dialog.getByLabel('Font', { exact: true }).fill('Arial');
  await dialog.getByLabel('Bold').check();
  await dialog.getByRole('tab', { name: 'Fill' }).click();
  await dialog.getByLabel('Pattern', { exact: true }).selectOption('solid');
  await pickCustomColour(dialog, 'format-fill-color', '#FFF2CC');
  await dialog.getByRole('tab', { name: 'Number' }).click();
  await dialog.getByLabel('Category').selectOption('currency');
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect(dialog).toBeHidden();

  await expect.poll(async () => (await styleAt(page, 1, 1))?.font.name).toBe('Arial');
  await expect.poll(async () => (await styleAt(page, 1, 1))?.fill.pattern).toBe('solid');
  await expect
    .poll(async () => (await styleAt(page, 1, 1))?.numberFormat)
    .toBe('$#,##0.00_);($#,##0.00)');

  // Font, fill and number format all revert with one Ctrl+Z.
  await page.keyboard.press('Control+z');
  await expect.poll(() => styleAt(page, 1, 1)).toEqual(before);
  expect(await page.evaluate(() => window.__calcplus.getState().canUndo)).toBe(false);
});

test('Format Cells rejects an invalid colour and number code without dispatching', async ({
  page,
}) => {
  await openFixture(page, fixture);
  await clickCell(page, 0, 0);
  const before = await styleAt(page, 0, 0);

  await page.getByRole('button', { name: 'Number Settings' }).click();
  const dialog = page.getByRole('dialog', { name: 'Format Cells' });
  await expect(dialog).toBeVisible();

  const colours = await openColours(dialog, 'format-font-color');
  await colours.getByRole('tab', { name: 'Custom' }).click();
  await colours.getByTestId('format-font-color-custom').fill('red');
  await colours.getByRole('button', { name: 'OK' }).click();
  await expect(colours.getByTestId('format-error-font-color')).toBeVisible();
  await expect(colours.getByTestId('format-font-color-custom')).toHaveAttribute(
    'aria-invalid',
    'true',
  );
  // The popup keeps its invalid draft; the dialog value and undo stack stay untouched.
  await expect(colours).toBeVisible();
  await expect(page.getByTestId('status-message')).not.toContainText('Formatted');
  expect(await page.evaluate(() => window.__calcplus.getState().canUndo)).toBe(false);
  expect(await styleAt(page, 0, 0)).toEqual(before);

  await colours.getByTestId('format-font-color-custom').fill('#FF0000');
  await colours.getByRole('button', { name: 'OK' }).click();
  await expect(colours).toBeHidden();
  await dialog.getByTestId('format-apply').click();
  await expect(dialog.getByTestId('format-error-font-color')).toHaveCount(0);
  await expect(dialog.getByTestId('format-applied')).toHaveText('Formatting applied.');
  await expect.poll(async () => (await styleAt(page, 0, 0))?.font.color).toBe('#FF0000');

  await dialog.getByRole('tab', { name: 'Number' }).click();
  await dialog.getByLabel('Format code').fill('0.00;0.00;0.00;@;garbage');
  await dialog.getByTestId('format-apply').click();
  await expect(dialog.getByTestId('format-error-number-code')).toBeVisible();
  await expect(dialog.getByTestId('format-error-summary')).toBeVisible();
  await expect(dialog.getByTestId('format-applied')).toHaveCount(0);
  expect((await styleAt(page, 0, 0))?.numberFormat).toBe(before?.numberFormat);
  expect(await page.evaluate(() => window.__calcplus.getState().canUndo)).toBe(true);

  await dialog.getByLabel('Format code').fill('0.00');
  await dialog.getByTestId('format-apply').click();
  await expect.poll(async () => (await styleAt(page, 0, 0))?.numberFormat).toBe('0.00');
  await expect(dialog.getByTestId('format-error-number-code')).toHaveCount(0);

  await dialog.getByRole('button', { name: 'Cancel' }).click();
  await expect(dialog).toBeHidden();
  await expect(page.getByRole('button', { name: 'Number Settings' })).toBeFocused();
});

test('Format Cells traps focus, closes on Esc and returns focus to the grid', async ({ page }) => {
  await openFixture(page, fixture);
  await clickCell(page, 1, 1);
  await page.getByRole('button', { name: 'Number Settings' }).click();
  const dialog = page.getByRole('dialog', { name: 'Format Cells' });
  await expect(dialog).toBeVisible();

  // The dialog owns the focus trap and its controls stay inside it.
  await dialog.getByLabel('Font', { exact: true }).focus();
  await page.keyboard.press('Tab');
  await expect(dialog.getByLabel('Size')).toBeFocused();

  await page.keyboard.press('Escape');
  await expect(dialog).toBeHidden();
  await expect(page.getByTestId('grid-shell')).toBeFocused();
  await expect.poll(() => page.evaluate(() => window.__calcplus.getState().focusZone)).toBe('grid');
});

test('Format Painter picks up, highlights the live target, applies once and Esc cancels', async ({
  page,
}) => {
  await openFixture(page, fixture);
  await clickCell(page, 0, 0);
  await page.keyboard.press('Control+b');
  await expect.poll(async () => (await styleAt(page, 0, 0))?.font.bold).toBe(true);

  const painter = page.getByRole('button', { name: 'Format Painter' });
  await painter.click();
  await expect(page.getByTestId('status-message')).toContainText('Format Painter picked up');
  await expect(page.getByTestId('grid-shell')).toBeFocused();

  // The live target is visible before the click commits the apply.
  const targetRect = await cellRect(page, 1, 1);
  const before = await regionSignature(
    page,
    targetRect.x,
    targetRect.y,
    targetRect.width,
    targetRect.height,
  );
  const targetPoint = await canvasPoint(page, 1, 1);
  await page.mouse.move(targetPoint.x, targetPoint.y);
  await expect
    .poll(() =>
      regionSignature(page, targetRect.x, targetRect.y, targetRect.width, targetRect.height),
    )
    .not.toBe(before);

  await page.mouse.click(targetPoint.x, targetPoint.y);
  await expect.poll(async () => (await styleAt(page, 1, 1))?.font.bold).toBe(true);

  // The apply is one undo transaction and removes the painter target.
  await page.keyboard.press('Control+z');
  await expect.poll(async () => (await styleAt(page, 1, 1))?.font.bold).toBe(false);
  await expect.poll(async () => (await styleAt(page, 0, 0))?.font.bold).toBe(true);

  // Esc abandons an active pickup without applying to the next click.
  await clickCell(page, 0, 0);
  await painter.click();
  await expect(page.getByTestId('status-message')).toContainText('Format Painter picked up');
  await page.keyboard.press('Escape');
  await expect(page.getByTestId('status-message')).toContainText('Format Painter cancelled');
  await clickCell(page, 2, 2);
  await expect.poll(async () => (await styleAt(page, 2, 2))?.font.bold).toBe(false);
});

test('Format Painter applies a dragged target range and is reachable from the context menu', async ({
  page,
}) => {
  await openFixture(page, fixture);
  await clickCell(page, 0, 0);
  await page.keyboard.press('Control+b');

  await page.getByRole('button', { name: 'Format Painter' }).click();
  const start = await canvasPoint(page, 1, 1);
  const end = await canvasPoint(page, 2, 2);
  await page.mouse.move(start.x, start.y);
  await page.mouse.down();
  await page.mouse.move(end.x, end.y, { steps: 5 });
  await page.mouse.up();

  await expect.poll(async () => (await styleAt(page, 1, 1))?.font.bold).toBe(true);
  await expect.poll(async () => (await styleAt(page, 2, 2))?.font.bold).toBe(true);

  await page.getByTestId('grid-canvas').click({ button: 'right' });
  const context = page.getByRole('menu', { name: 'Context menu' });
  await expect(context.getByRole('menuitem', { name: 'Format Painter' })).toBeVisible();
  await expect(context.getByRole('menuitem', { name: 'Format Cells' })).toBeVisible();
});
