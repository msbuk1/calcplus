import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';
import { clickCell, openFixture } from './grid-helpers';

const fixture = 'm1-chart-unknown.xlsx';

async function styleAt(page: Page, row: number, col: number) {
  return page.evaluate(
    ([targetRow, targetCol]) => window.__calcplus.getCellStyle(targetRow, targetCol),
    [row, col] as const,
  );
}

test('quick access sits above the ribbon at the top left with icons and tooltips', async ({
  page,
}) => {
  await openFixture(page, fixture);

  const toolbar = page.getByRole('toolbar', { name: 'Quick access' });
  await expect(toolbar).toBeVisible();
  const save = toolbar.getByTestId('save-button');
  const undo = toolbar.getByTestId('undo-button');
  const redo = toolbar.getByTestId('redo-button');

  await expect(save).toHaveAttribute('title', 'Save (Ctrl+S)');
  await expect(save).toHaveAttribute('aria-label', 'Save (Ctrl+S)');
  await expect(undo).toHaveAttribute('title', 'Undo (Ctrl+Z)');
  await expect(redo).toHaveAttribute('title', 'Redo (Ctrl+Y)');
  for (const button of [save, undo, redo]) {
    await expect(button.locator('svg.ribbon-glyph')).toHaveCount(1);
    await expect(button).toHaveText('');
  }

  // The toolbar counts as the ribbon focus zone; focusing it must not bounce
  // the keyboard to the ribbon body.
  await save.focus();
  await expect(save).toBeFocused();
  expect(await page.evaluate(() => window.__calcplus.getState().focusZone)).toBe('ribbon');

  // Save As is not in the toolbar; it stays in the File menu.
  await expect(toolbar.getByRole('button', { name: /Save As/ })).toHaveCount(0);
  await page.getByTestId('file-menu-button').click();
  await expect(page.getByTestId('save-as-button')).toBeVisible();

  // Above the ribbon tabs, and on the left half of the window.
  const toolbarBox = await toolbar.boundingBox();
  const tabsBox = await page.locator('.ribbon-tabs').boundingBox();
  expect(toolbarBox).not.toBeNull();
  expect(tabsBox).not.toBeNull();
  expect(toolbarBox!.y).toBeLessThan(tabsBox!.y);
  const viewport = page.viewportSize()!;
  expect(toolbarBox!.x).toBeLessThan(viewport.width / 2);
});

test('unsaved changes show on the Save button and no separate dot remains', async ({ page }) => {
  await openFixture(page, fixture);
  const save = page.getByTestId('save-button');
  await expect(save).not.toHaveAttribute('data-unsaved', 'true');

  await clickCell(page, 1, 1);
  await page.keyboard.type('43');
  await page.keyboard.press('Enter');

  await expect(save).toHaveAttribute('data-unsaved', 'true');
  await expect(page.getByTestId('save-unsaved-dot')).toBeVisible();
  await expect(page.getByTestId('dirty-indicator')).toHaveCount(0);
});

test('below 1100 px the quick access keeps its icons while group labels hide', async ({ page }) => {
  await page.setViewportSize({ width: 1000, height: 800 });
  await openFixture(page, fixture);

  await expect(page.getByTestId('save-button')).toBeVisible();
  await expect(page.getByTestId('undo-button')).toBeVisible();
  await expect(page.getByTestId('redo-button')).toBeVisible();
  await expect(page.getByTestId('save-button')).toHaveAttribute('title', 'Save (Ctrl+S)');
  await expect(page.locator('.ribbon-group-label').first()).toBeHidden();
});

test('text colour opens the shared picker, applies a swatch once and undoes once', async ({
  page,
}) => {
  await openFixture(page, fixture);
  await clickCell(page, 1, 1);

  await page.getByRole('button', { name: 'Font Colour options' }).click();
  const picker = page.getByRole('dialog', { name: 'Font Colour' });
  await expect(picker).toBeVisible();
  await expect(picker.getByRole('button', { name: 'Colour #FF0000' })).toBeVisible();
  await expect(picker.getByRole('button', { name: 'Automatic' })).toBeVisible();

  await picker.getByRole('button', { name: 'Colour #FF0000' }).click();
  await expect(picker).toBeHidden();
  await expect.poll(async () => (await styleAt(page, 1, 1))?.font.color).toBe('#FF0000');
  expect(await page.evaluate(() => window.__calcplus.getState().canUndo)).toBe(true);

  await page.keyboard.press('Control+z');
  await expect.poll(async () => (await styleAt(page, 1, 1))?.font.color).not.toBe('#FF0000');
  expect(await page.evaluate(() => window.__calcplus.getState().canUndo)).toBe(false);
});

test('More Colours opens the hex picker and applies one undoable change', async ({ page }) => {
  await openFixture(page, fixture);
  await clickCell(page, 1, 1);

  await page.getByRole('button', { name: 'Font Colour options' }).click();
  const picker = page.getByRole('dialog', { name: 'Font Colour' });
  await picker.getByRole('button', { name: 'More Colours' }).click();
  await expect(picker.getByRole('tab', { name: 'Custom' })).toHaveAttribute(
    'aria-selected',
    'true',
  );

  const hex = picker.getByTestId('font-color-custom');
  await expect(hex).toBeVisible();
  await hex.fill('#1a73e8');
  await picker.getByRole('button', { name: 'Apply' }).click();
  await expect(picker).toBeHidden();
  await expect.poll(async () => (await styleAt(page, 1, 1))?.font.color).toBe('#1A73E8');

  await page.keyboard.press('Control+z');
  await expect.poll(async () => (await styleAt(page, 1, 1))?.font.color).not.toBe('#1A73E8');
});

test('fill colour offers the swatch grid, No Fill and the shared hex picker', async ({ page }) => {
  await openFixture(page, fixture);
  await clickCell(page, 1, 1);

  await page.getByRole('button', { name: 'Fill Colour options' }).click();
  const picker = page.getByRole('dialog', { name: 'Fill Colour' });
  await expect(picker.getByRole('button', { name: 'Colour #FFC000' })).toBeVisible();
  await expect(picker.getByRole('button', { name: 'More Colours' })).toBeVisible();

  await picker.getByRole('button', { name: 'Colour #FFC000' }).click();
  await expect(picker).toBeHidden();
  await expect
    .poll(async () => (await styleAt(page, 1, 1))?.fill)
    .toMatchObject({
      pattern: 'solid',
      color: '#FFC000',
    });

  await page.getByRole('button', { name: 'Fill Colour options' }).click();
  await picker.getByRole('button', { name: 'No Fill' }).click();
  await expect(picker).toBeHidden();
  await expect
    .poll(async () => (await styleAt(page, 1, 1))?.fill)
    .toMatchObject({
      pattern: 'none',
    });
});

test('the colour picker is keyboard reachable and Esc returns focus to its trigger', async ({
  page,
}) => {
  await openFixture(page, fixture);
  await clickCell(page, 1, 1);

  const toggle = page.getByRole('button', { name: 'Font Colour options' });
  await toggle.focus();
  await page.keyboard.press('Enter');
  const picker = page.getByRole('dialog', { name: 'Font Colour' });
  await expect(picker).toBeVisible();
  await expect(picker.getByRole('tab', { name: 'Standard' })).toBeFocused();

  // Arrow keys walk the roving swatch grid; every swatch has an accessible name.
  const themeSwatches = picker.getByRole('group', { name: 'Theme colours' }).getByRole('button');
  await expect(themeSwatches.first()).toHaveAttribute('aria-label', /^Colour #[0-9A-F]{6}$/);
  await themeSwatches.first().focus();
  await page.keyboard.press('ArrowRight');
  await expect(themeSwatches.nth(1)).toBeFocused();

  // Tab closes the popup and returns to the trigger, as the ribbon menus do.
  await page.keyboard.press('Tab');
  await expect(picker).toBeHidden();
  await expect(toggle).toBeFocused();

  // Escape closes from any control and restores the trigger.
  await page.keyboard.press('Enter');
  await expect(picker).toBeVisible();
  await page.keyboard.press('Escape');
  await expect(picker).toBeHidden();
  await expect(toggle).toBeFocused();
});
