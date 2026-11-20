import { expect, test } from '@playwright/test';
import type { Locator, Page } from '@playwright/test';
import { cellRect, clickCell, openFixture, regionSignature } from './grid-helpers';

const fixture = 'm1-chart-unknown.xlsx';

type ColourField =
  'format-font-color' | 'format-fill-color' | 'format-fill-background' | 'format-border-color';

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

/**
 * The Pattern control is a custom listbox (a native select cannot host the
 * drawn previews), so a choice is open, then the option by name.
 */
async function choosePattern(dialog: Locator, name: string): Promise<void> {
  const control = dialog.getByTestId('format-fill-pattern');
  await control.click();
  await expect(control).toHaveAttribute('aria-expanded', 'true');
  await dialog.getByRole('option', { name, exact: true }).click();
  await expect(control).toHaveAttribute('aria-expanded', 'false');
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
  // The launcher opens on its own tab (ticket 28); this walk starts at Font.
  await dialog.getByRole('tab', { name: 'Font' }).click();
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
  await choosePattern(dialog, 'Solid');
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
  // The launcher opens on its own tab (ticket 28); this walk starts at Font.
  await dialog.getByRole('tab', { name: 'Font' }).click();

  await dialog.getByLabel('Font', { exact: true }).fill('Arial');
  await dialog.getByLabel('Bold').check();
  await dialog.getByRole('tab', { name: 'Fill' }).click();
  await choosePattern(dialog, 'Solid');
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

test('the Fill pattern control previews every pattern in the chosen colours', async ({ page }) => {
  await openFixture(page, fixture);
  await clickCell(page, 1, 1);

  await page.getByRole('button', { name: 'Number Settings' }).click();
  const dialog = page.getByRole('dialog', { name: 'Format Cells' });
  await expect(dialog).toBeVisible();
  await dialog.getByRole('tab', { name: 'Fill' }).click();

  await pickCustomColour(dialog, 'format-fill-color', '#C00000');
  await pickCustomColour(dialog, 'format-fill-background', '#EAF1F8');

  const control = dialog.getByTestId('format-fill-pattern');
  const closedSwatch = control.locator('.format-pattern-swatch');
  // None is the starting pattern, so the closed preview is a blank white swatch.
  await expect(control).toHaveAttribute('aria-expanded', 'false');
  await expect(closedSwatch).toHaveCSS('background-color', 'rgb(255, 255, 255)');

  await control.click();
  await expect(control).toHaveAttribute('aria-expanded', 'true');
  await expect(dialog.getByRole('option')).toHaveCount(19);

  const grid = dialog.getByRole('option', { name: 'Light grid', exact: true });
  const gridSwatch = grid.locator('.format-pattern-swatch');
  // The texture is the fill colour; the pattern background shows through it.
  await expect(gridSwatch).toHaveCSS('background-color', 'rgb(234, 241, 248)');
  await expect(gridSwatch).toHaveCSS('background-image', /data:image\/svg/);
  await expect(gridSwatch).toHaveCSS('background-image', /C00000/);
  // Solid previews the fill colour itself and None the blank cell.
  const solid = dialog.getByRole('option', { name: 'Solid', exact: true });
  await expect(solid.locator('.format-pattern-swatch')).toHaveCSS(
    'background-color',
    'rgb(192, 0, 0)',
  );
  await expect(solid.locator('.format-pattern-swatch')).toHaveCSS('background-image', 'none');
  const none = dialog.getByRole('option', { name: 'None', exact: true });
  await expect(none.locator('.format-pattern-swatch')).toHaveCSS(
    'background-color',
    'rgb(255, 255, 255)',
  );

  // Esc closes the list only: the dialog stays open and nothing was chosen.
  await page.keyboard.press('Escape');
  await expect(control).toHaveAttribute('aria-expanded', 'false');
  await expect(control).toBeFocused();
  await expect(dialog).toBeVisible();
  await expect(control).toContainText('None');

  // Arrows walk the list from the selection and Enter chooses.
  await page.keyboard.press('ArrowDown');
  await expect(control).toHaveAttribute('aria-expanded', 'true');
  await expect(dialog.getByRole('listbox', { name: 'Pattern' })).toBeFocused();
  for (let step = 0; step < 15; step += 1) await page.keyboard.press('ArrowDown');
  await page.keyboard.press('Enter');
  await expect(control).toHaveAttribute('aria-expanded', 'false');
  await expect(control).toBeFocused();
  await expect(control).toContainText('Light grid');

  // The closed control previews the chosen pattern in the current colours.
  await expect(closedSwatch).toHaveCSS('background-color', 'rgb(234, 241, 248)');
  await expect(closedSwatch).toHaveCSS('background-image', /C00000/);

  // The colour controls re-render the previews: fill colour, then background.
  await pickCustomColour(dialog, 'format-fill-color', '#1F6FEB');
  await expect(closedSwatch).toHaveCSS('background-image', /1F6FEB/);
  await pickCustomColour(dialog, 'format-fill-background', '#FFF2CC');
  await expect(closedSwatch).toHaveCSS('background-color', 'rgb(255, 242, 204)');
  await control.click();
  await expect(gridSwatch).toHaveCSS('background-color', 'rgb(255, 242, 204)');
  await expect(gridSwatch).toHaveCSS('background-image', /1F6FEB/);
  await page.keyboard.press('Escape');

  // The choice reaches format.fill with the same arguments as before.
  await dialog.getByTestId('format-apply').click();
  await expect(dialog.getByTestId('format-applied')).toHaveText('Formatted B2.');
  await expect
    .poll(() => styleAt(page, 1, 1))
    .toMatchObject({
      fill: { pattern: 'lightGrid', color: '#1F6FEB', backgroundColor: '#FFF2CC' },
    });
  await dialog.getByRole('button', { name: 'Cancel' }).click();
  await expect(dialog).toBeHidden();
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
  // The launcher opens on its own tab (ticket 28); this walk starts at Font.
  await dialog.getByRole('tab', { name: 'Font' }).click();

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
  await expect(dialog.getByTestId('format-applied')).toHaveText('Formatted A1:B1.');
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

  // Point 5: applying the same value again writes nothing, and the dialog says
  // so instead of claiming it formatted the cell (issue 35).
  await dialog.getByTestId('format-apply').click();
  await expect(dialog.getByTestId('format-applied')).toHaveText('No formatting change for A1:B1.');

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
  // The launcher opens on its own tab (ticket 28); this walk starts at Font.
  await dialog.getByRole('tab', { name: 'Font' }).click();

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

test('group launchers open the dialog on their own tab', async ({ page }) => {
  await openFixture(page, fixture);

  await page.getByRole('button', { name: 'Alignment Settings' }).click();
  const dialog = page.getByRole('dialog', { name: 'Format Cells' });
  await expect(dialog).toBeVisible();
  await expect(dialog.getByRole('tab', { name: 'Alignment' })).toHaveAttribute(
    'aria-selected',
    'true',
  );
  await page.keyboard.press('Escape');
  await expect(dialog).toBeHidden();

  await page.getByRole('button', { name: 'Number Settings' }).click();
  await expect(dialog.getByRole('tab', { name: 'Number' })).toHaveAttribute(
    'aria-selected',
    'true',
  );
});
