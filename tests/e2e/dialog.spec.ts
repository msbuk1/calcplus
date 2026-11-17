import { expect, test } from '@playwright/test';
import { openFixture } from './grid-helpers';

test('About traps focus, keeps backdrop clicks inside and closes on Esc', async ({ page }) => {
  await openFixture(page, 'm1-chart-unknown.xlsx');
  await page.getByTestId('file-menu-button').click();
  await page.getByRole('menuitem', { name: 'About calcplus' }).click();

  const dialog = page.getByRole('dialog', { name: 'About calcplus' });
  await expect(dialog).toBeVisible();
  const close = dialog.getByRole('button', { name: 'Close', exact: true });
  await expect(close).toBeFocused();

  // A backdrop press must not dismiss the dialog or let focus reach the chrome.
  await page.locator('.dialog-backdrop').click({ position: { x: 4, y: 4 } });
  await expect(dialog).toBeVisible();
  await expect(close).toBeFocused();

  // The dialog root is focusable; Shift+Tab from it stays inside and wraps to
  // the last control instead of reaching the chrome behind the modal.
  await dialog.focus();
  await expect(dialog).toBeFocused();
  await page.keyboard.press('Shift+Tab');
  await expect(close).toBeFocused();

  // Tab wraps from the last control to the first.
  await page.keyboard.press('Tab');
  await expect(dialog.getByRole('button', { name: 'Close dialog' })).toBeFocused();

  // Shift+Tab wraps from the first control to the last.
  await page.keyboard.press('Shift+Tab');
  await expect(close).toBeFocused();

  await page.keyboard.press('Escape');
  await expect(dialog).toBeHidden();
  await expect(page.getByTestId('grid-shell')).toBeFocused();
  await expect.poll(() => page.evaluate(() => window.__calcplus.getState().focusZone)).toBe('grid');
});

test('About returns focus to the invoking control when dismissed without Esc', async ({ page }) => {
  await openFixture(page, 'm1-chart-unknown.xlsx');
  const fileButton = page.getByTestId('file-menu-button');
  const dialog = page.getByRole('dialog', { name: 'About calcplus' });

  // The footer action behaves like the header close button: both dismiss
  // without cancelling, so focus returns to the File control that invoked it.
  await fileButton.click();
  await page.getByRole('menuitem', { name: 'About calcplus' }).click();
  await expect(dialog).toBeVisible();
  await dialog.getByRole('button', { name: 'Close', exact: true }).click();
  await expect(dialog).toBeHidden();
  await expect(fileButton).toBeFocused();

  await fileButton.click();
  await page.getByRole('menuitem', { name: 'About calcplus' }).click();
  await expect(dialog).toBeVisible();
  await dialog.getByRole('button', { name: 'Close dialog' }).click();
  await expect(dialog).toBeHidden();
  await expect(fileButton).toBeFocused();
});
