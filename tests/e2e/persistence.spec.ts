import { expect, test } from '@playwright/test';
import { clickCell, openFixture } from './grid-helpers';

/**
 * M7 lane 5 — the persistence flows (persistence §2, §3).
 *
 * The recent list remembers a file by name on this origin. The recovery flow
 * edits a workbook, forces the debounced autosave, then reloads: the browser
 * first asks about the unsaved page (its own wording), then calcplus offers
 * the recovery copy. Accepting it restores the edit even though the file on
 * disk never changed. These specs assert structure and model state; they never
 * create a platform screenshot.
 */
test('lists a recently opened file on this origin and removes it', async ({ page }) => {
  await openFixture(page, 'm1-chart-unknown.xlsx');

  await page.getByTestId('file-menu-button').click();
  await page.locator('[data-command="file.recent"]').click();
  const dialog = page.getByRole('dialog', { name: 'Recent files' });
  await expect(dialog.getByTestId('recent-list')).toBeVisible();
  await expect(dialog.getByText('m1-chart-unknown.xlsx')).toBeVisible();
  await expect(dialog.getByText('Name only')).toBeVisible();

  await dialog.getByRole('button', { name: /Remove m1-chart-unknown.xlsx/ }).click();
  await expect(page.getByTestId('recent-empty')).toBeVisible();
});

test('discards a recovery copy when the prompt is declined', async ({ page }) => {
  await openFixture(page, 'm1-chart-unknown.xlsx');
  await clickCell(page, 1, 1);
  await page.keyboard.type('99');
  await page.keyboard.press('Enter');
  await page.evaluate(async () => {
    const debug = window.__calcplus as unknown as { flushRecovery(): Promise<void> };
    await debug.flushRecovery();
  });

  const messages: string[] = [];
  page.on('dialog', async (dialog) => {
    messages.push(dialog.message());
    if (/Restore/.test(dialog.message())) await dialog.dismiss();
    else await dialog.accept();
  });

  await page.reload();

  await expect
    .poll(() => messages.some((message) => /Restore/.test(message)), { timeout: 10_000 })
    .toBe(true);
  await expect(page.getByTestId('file-name')).toHaveText('No workbook open');
  await expect(page.getByTestId('status-message')).toContainText('Recovery copy discarded');
  await expect
    .poll(() =>
      page.evaluate(() => {
        const debug = window.__calcplus as unknown as {
          getRecoveryPending(): Promise<readonly unknown[]>;
        };
        return debug.getRecoveryPending();
      }),
    )
    .toEqual([]);
});

test('restores an unsaved recovery copy after a reload', async ({ page }) => {
  await openFixture(page, 'm1-chart-unknown.xlsx');

  await clickCell(page, 1, 1);
  await page.keyboard.type('43');
  await page.keyboard.press('Enter');
  await expect(page.getByTestId('save-button')).toHaveAttribute('data-unsaved', 'true');

  await page.evaluate(async () => {
    const debug = window.__calcplus as unknown as { flushRecovery(): Promise<void> };
    await debug.flushRecovery();
  });
  const pending = await page.evaluate(() => {
    const debug = window.__calcplus as unknown as {
      getRecoveryPending(): Promise<readonly { displayName: string }[]>;
    };
    return debug.getRecoveryPending();
  });
  expect(pending).toHaveLength(1);
  expect(pending[0]!.displayName).toBe('m1-chart-unknown.xlsx');

  const messages: string[] = [];
  page.on('dialog', async (dialog) => {
    messages.push(dialog.message());
    await dialog.accept();
  });

  await page.reload();

  await expect
    .poll(() => messages.some((message) => /Restore/.test(message)), { timeout: 10_000 })
    .toBe(true);
  await expect(page.getByTestId('file-name')).toHaveText('m1-chart-unknown.xlsx');
  await expect(page.getByTestId('status-message')).toContainText('Restored unsaved changes');
  await expect(page.getByTestId('save-button')).toHaveAttribute('data-unsaved', 'true');
  await expect
    .poll(() =>
      page.evaluate(() => {
        const debug = window.__calcplus as unknown as {
          getCellDisplay(row: number, col: number): string | null;
        };
        return debug.getCellDisplay(1, 1);
      }),
    )
    .toBe('43');

  // Saving the restored document clears the recovery copy.
  const download = page.waitForEvent('download');
  await page.getByTestId('save-button').click();
  await download;
  await expect
    .poll(() =>
      page.evaluate(() => {
        const debug = window.__calcplus as unknown as {
          getRecoveryPending(): Promise<readonly unknown[]>;
        };
        return debug.getRecoveryPending();
      }),
    )
    .toEqual([]);
});
