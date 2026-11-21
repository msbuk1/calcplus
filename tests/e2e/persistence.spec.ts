import { Buffer } from 'node:buffer';
import { expect, test, type Page } from '@playwright/test';
import { clickCell, fixtureBytes, openFixture } from './grid-helpers';
import { FAKE_FILE_NAME, installFakeFs } from './fake-fs';

/**
 * M7 lane 5 — the persistence flows (persistence §2, §3).
 *
 * The recent list remembers a file by name on this origin. The recovery flow
 * edits a workbook, forces the debounced autosave, then reloads: the browser
 * first asks about the unsaved page (its own wording), then calcplus offers
 * the recovery copy. Accepting it restores the edit even though the file on
 * disk never changed. These specs assert structure and model state; they never
 * create a platform screenshot.
 *
 * The last three specs open through the real File System Access path instead
 * of the `?fsa=0` fallback (see `fake-fs.ts`). They cover the bug where a
 * file handle revived from a recovery copy carries no live permission, so
 * every Save after Restore forced Save As. That is wrong per persistence §4:
 * save must re-ask for permission (§1/§3) and write in place.
 *
 * The conflict specs drive the external-change dialog (persistence §4): the
 * file change comes from `__fakeFs.externalEdit`, and the fake save picker
 * enforces Chrome's user-gesture rule, so a Save As that reaches the picker
 * without its own button click fails exactly like the browser fails.
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

/** Opens the fixture through the File menu, so `document.handle` is set. */
async function openThroughFileMenu(page: Page): Promise<void> {
  await page.getByTestId('file-menu-button').click();
  // Scoped to the open menu: the empty state has its own Open button.
  await page.getByRole('menu').locator('[data-command="file.open"]').click();
  await expect(page.getByTestId('file-name')).toHaveText(FAKE_FILE_NAME);
  await expect(page.getByTestId('grid-canvas')).toBeVisible();
}

/** Records every `window.confirm` message so wording can be asserted. */
function recordPrompts(target: Page, accept: boolean): string[] {
  const prompts: string[] = [];
  target.on('dialog', async (dialog) => {
    prompts.push(dialog.message());
    if (accept) await dialog.accept();
    else await dialog.dismiss();
  });
  return prompts;
}

test('restores after the tab dies, re-asks for permission and saves in place', async ({
  page,
  context,
}) => {
  await installFakeFs(context);
  await page.goto('/');

  await openThroughFileMenu(page);
  await clickCell(page, 1, 1);
  await page.keyboard.type('43');
  await page.keyboard.press('Enter');
  await page.evaluate(async () => {
    const debug = window.__calcplus as unknown as { flushRecovery(): Promise<void> };
    await debug.flushRecovery();
  });

  // The tab dies. Closing the page is the honest version: the recovery copy
  // in IndexedDB survives, and the handle it carries comes back with no live
  // permission. A reload keeps too much of the old document around.
  await page.close();
  const revivedPage = await context.newPage();
  const prompts = recordPrompts(revivedPage, true);
  await revivedPage.goto('/');

  await expect
    .poll(() => prompts.some((message) => /Restore them\?/.test(message)), { timeout: 10_000 })
    .toBe(true);
  await expect(revivedPage.getByTestId('file-name')).toHaveText(FAKE_FILE_NAME);
  await expect(revivedPage.getByTestId('status-message')).toContainText('Restored unsaved changes');
  await expect
    .poll(() =>
      revivedPage.evaluate(() => {
        const debug = window.__calcplus as unknown as {
          getCellDisplay(row: number, col: number): string | null;
        };
        return debug.getCellDisplay(1, 1);
      }),
    )
    .toBe('43');

  const downloads: string[] = [];
  revivedPage.on('download', (download) => {
    downloads.push(download.suggestedFilename());
  });
  await revivedPage.getByTestId('save-button').click();

  await expect
    .poll(() => revivedPage.evaluate(() => window.__fakeFs.calls.close))
    .toBeGreaterThan(0);
  const calls = await revivedPage.evaluate(() => window.__fakeFs.calls);
  const state = await revivedPage.evaluate(() => window.__calcplus.getState());

  // Nothing in this page picked a file: the only handle is the revived one.
  expect(calls.openPicker).toBe(0);
  // The revived handle had no live permission, so Save re-asked for it.
  expect(calls.queryPermission.some((entry) => entry.mode === 'readwrite')).toBe(true);
  expect(calls.requestPermission.some((entry) => entry.mode === 'readwrite')).toBe(true);
  // Save wrote in place through createWritable, not a download or a Save As.
  expect(calls.getFile).toBeGreaterThan(0);
  expect(calls.createWritable).toBeGreaterThan(0);
  expect(calls.write).toBeGreaterThan(0);
  expect(calls.savePicker).toBe(0);
  expect(downloads).toEqual([]);
  // No "Save as a new file instead?" prompt appeared; only the Restore one.
  expect(prompts.filter((message) => /Save as a new file instead\?/.test(message))).toEqual([]);
  expect(prompts.filter((message) => !/Restore them\?/.test(message))).toEqual([]);
  await expect(revivedPage.getByTestId('status-message')).toContainText('Saved');
  await expect(revivedPage.getByTestId('status-message')).not.toContainText('downloaded copy');
  // A clean document drops the unsaved marker entirely.
  await expect(revivedPage.getByTestId('save-button')).not.toHaveAttribute('data-unsaved', 'true');
  expect(state.fileName).toBe(FAKE_FILE_NAME);
  // Saving the restored document clears the recovery copy.
  await expect
    .poll(() =>
      revivedPage.evaluate(() => {
        const debug = window.__calcplus as unknown as {
          getRecoveryPending(): Promise<readonly unknown[]>;
        };
        return debug.getRecoveryPending();
      }),
    )
    .toEqual([]);
});

test('says it does not have permission to write when save access is denied', async ({
  page,
  context,
}) => {
  await installFakeFs(context);
  await page.goto('/');
  await openThroughFileMenu(page);

  // The handle loses access and the user denies the re-request.
  await page.evaluate(() => {
    window.__fakeFs.plan.query = 'prompt';
    window.__fakeFs.plan.request = 'denied';
  });
  const prompts = recordPrompts(page, true);
  await page.getByTestId('save-button').click();

  // The prompt is an in-page dialog now: its "Save as a new file" button must
  // open the picker from its own click handler, which a native confirm cannot
  // do. The wording names the file and the failure.
  const dialog = page.getByRole('dialog', { name: 'Save as a new file instead?' });
  await expect(dialog).toBeVisible();
  await expect(dialog.getByTestId('choice-message')).toHaveText(
    'calcplus does not have permission to write "m1-chart-unknown.xlsx".',
  );
  await dialog.getByRole('button', { name: 'Cancel' }).click();

  // The permission was asked for before the prompt; nothing was written and
  // no native confirm was used anywhere in the flow.
  const calls = await page.evaluate(() => window.__fakeFs.calls);
  expect(calls.requestPermission.some((entry) => entry.mode === 'readwrite')).toBe(true);
  expect(calls.createWritable).toBe(0);
  expect(calls.savePicker).toBe(0);
  expect(prompts).toEqual([]);
  await expect(page.getByTestId('status-message')).toHaveText(
    'Save cancelled. The file was not touched.',
  );
});

test('says it could not read the file when the stamp is unreadable despite permission', async ({
  page,
  context,
}) => {
  await installFakeFs(context);
  await page.goto('/');
  await openThroughFileMenu(page);

  // Permission is fine, but the file cannot be read at all.
  await page.evaluate(() => {
    window.__fakeFs.plan.getFileThrows = true;
  });
  const prompts = recordPrompts(page, true);
  await page.getByTestId('save-button').click();

  const dialog = page.getByRole('dialog', { name: 'Save as a new file instead?' });
  await expect(dialog).toBeVisible();
  await expect(dialog.getByTestId('choice-message')).toHaveText(
    'calcplus could not read "m1-chart-unknown.xlsx" to check for external changes.',
  );
  await dialog.getByRole('button', { name: 'Cancel' }).click();

  // The stamp was read (and failed) despite granted permission.
  const calls = await page.evaluate(() => window.__fakeFs.calls);
  expect(calls.getFile).toBeGreaterThan(0);
  expect(calls.requestPermission).toEqual([]);
  expect(calls.createWritable).toBe(0);
  expect(prompts).toEqual([]);
  await expect(page.getByTestId('status-message')).toHaveText(
    'Save cancelled. The file was not touched.',
  );
});

/** Opens the fixture through the File menu and leaves one unsaved edit. */
async function openWithUnsavedEdit(page: Page): Promise<void> {
  await openThroughFileMenu(page);
  await clickCell(page, 1, 1);
  await page.keyboard.type('43');
  await page.keyboard.press('Enter');
  await expect(page.getByTestId('save-button')).toHaveAttribute('data-unsaved', 'true');
}

test('offers Reload, Overwrite, Save As and Cancel in order when the file changed on disk', async ({
  page,
  context,
}) => {
  await installFakeFs(context);
  await page.goto('/');
  await openWithUnsavedEdit(page);
  await page.evaluate(() => {
    window.__fakeFs.externalEdit();
  });
  await page.getByTestId('save-button').click();

  const dialog = page.getByRole('dialog', { name: 'File changed on disk' });
  await expect(dialog).toBeVisible();
  await expect(dialog.getByTestId('choice-message')).toContainText('changed on disk');
  // The four buttons are in the contract's order (persistence §4).
  const labels = await dialog
    .locator('.choice-action button, .dialog-actions button')
    .allTextContents();
  expect(labels).toEqual(['Reload', 'Overwrite', 'Save As', 'Cancel']);
  // Both destructive outcomes are stated plainly: one of them loses work
  // either way.
  const statements = await dialog.locator('.choice-statement').allTextContents();
  expect(statements[0]).toContain('unsaved changes are lost');
  expect(statements[1]).toContain('external edit is lost');
  // Nothing has been written while the question is open.
  const calls = await page.evaluate(() => window.__fakeFs.calls);
  expect(calls.write).toBe(0);
});

test('choosing Save As opens the picker and writes to the new file', async ({ page, context }) => {
  await installFakeFs(context);
  await page.goto('/');
  await openWithUnsavedEdit(page);
  await page.evaluate(() => {
    window.__fakeFs.externalEdit();
  });
  await page.getByTestId('save-button').click();

  const dialog = page.getByRole('dialog', { name: 'File changed on disk' });
  await expect(dialog).toBeVisible();
  await dialog.getByRole('button', { name: 'Save As' }).click();

  await expect.poll(() => page.evaluate(() => window.__fakeFs.calls.close)).toBeGreaterThan(0);
  const calls = await page.evaluate(() => window.__fakeFs.calls);
  // The picker opened from the button's own click and the write went through
  // the handle the picker handed out; the original file was not touched.
  expect(calls.savePicker).toBe(1);
  expect(calls.createWritableHandles).toEqual(calls.savePickerHandles);
  expect(calls.write).toBeGreaterThan(0);
  await expect(page.getByTestId('status-message')).toContainText('Saved');
  await expect(page.getByTestId('save-button')).not.toHaveAttribute('data-unsaved', 'true');
});

test('choosing Overwrite writes to the original file after an external edit', async ({
  page,
  context,
}) => {
  await installFakeFs(context);
  await page.goto('/');
  await openWithUnsavedEdit(page);
  await page.evaluate(() => {
    window.__fakeFs.externalEdit();
  });
  await page.getByTestId('save-button').click();

  const dialog = page.getByRole('dialog', { name: 'File changed on disk' });
  await expect(dialog).toBeVisible();
  await dialog.getByRole('button', { name: 'Overwrite' }).click();

  await expect.poll(() => page.evaluate(() => window.__fakeFs.calls.close)).toBeGreaterThan(0);
  const calls = await page.evaluate(() => window.__fakeFs.calls);
  // Overwrite keeps the local model and writes it over the original file: no
  // picker, and the write went through the handle the workbook already had.
  expect(calls.savePicker).toBe(0);
  expect(calls.savePickerHandles).toEqual([]);
  expect(calls.createWritableHandles).not.toEqual([]);
  expect(calls.write).toBeGreaterThan(0);
  await expect(page.getByTestId('status-message')).toContainText('Saved');
  await expect(page.getByTestId('save-button')).not.toHaveAttribute('data-unsaved', 'true');
});

test('choosing Reload discards local changes and takes the file on disk', async ({
  page,
  context,
}) => {
  // What the external writer put on disk, read from the real replacement file.
  const scratch = await context.newPage();
  await openFixture(scratch, 'm2-grid.xlsx');
  const diskValue = await scratch.evaluate(() => window.__calcplus.getCellDisplay(1, 1));
  await scratch.close();
  expect(diskValue).not.toBe('43');

  await installFakeFs(context);
  await page.goto('/');
  await openWithUnsavedEdit(page);
  await page.evaluate(
    (base64) => window.__fakeFs.externalEdit(base64),
    Buffer.from(fixtureBytes('m2-grid.xlsx')).toString('base64'),
  );
  await page.getByTestId('save-button').click();

  const dialog = page.getByRole('dialog', { name: 'File changed on disk' });
  await expect(dialog).toBeVisible();
  await dialog.getByRole('button', { name: 'Reload' }).click();

  await expect(page.getByTestId('status-message')).toContainText(
    'Reloaded the file from disk. Local changes were discarded.',
  );
  // The model now shows the disk version, not the local edit.
  await expect
    .poll(() => page.evaluate(() => window.__calcplus.getCellDisplay(1, 1)))
    .toBe(diskValue);
  await expect(page.getByTestId('save-button')).not.toHaveAttribute('data-unsaved', 'true');
  // Reload writes nothing.
  const calls = await page.evaluate(() => window.__fakeFs.calls);
  expect(calls.createWritable).toBe(0);
  expect(calls.savePicker).toBe(0);
});

test('declining with Cancel writes nothing at all', async ({ page, context }) => {
  await installFakeFs(context);
  await page.goto('/');
  await openWithUnsavedEdit(page);
  await page.evaluate(() => {
    window.__fakeFs.externalEdit();
  });
  await page.getByTestId('save-button').click();

  const dialog = page.getByRole('dialog', { name: 'File changed on disk' });
  await expect(dialog).toBeVisible();
  await dialog.getByRole('button', { name: 'Cancel' }).click();

  await expect(page.getByTestId('status-message')).toHaveText(
    'Save cancelled. The newer file on disk was not touched.',
  );
  const calls = await page.evaluate(() => window.__fakeFs.calls);
  expect(calls.createWritable).toBe(0);
  expect(calls.write).toBe(0);
  expect(calls.savePicker).toBe(0);
  await expect(page.getByTestId('save-button')).toHaveAttribute('data-unsaved', 'true');
});

test('Esc on the conflict dialog writes nothing', async ({ page, context }) => {
  await installFakeFs(context);
  await page.goto('/');
  await openWithUnsavedEdit(page);
  await page.evaluate(() => {
    window.__fakeFs.externalEdit();
  });
  await page.getByTestId('save-button').click();

  const dialog = page.getByRole('dialog', { name: 'File changed on disk' });
  await expect(dialog).toBeVisible();
  await page.keyboard.press('Escape');

  await expect(dialog).toBeHidden();
  await expect(page.getByTestId('status-message')).toHaveText(
    'Save cancelled. The newer file on disk was not touched.',
  );
  const calls = await page.evaluate(() => window.__fakeFs.calls);
  expect(calls.write).toBe(0);
  expect(calls.savePicker).toBe(0);
});

test('the picker opens after the permission prompt when Save As follows it', async ({
  page,
  context,
}) => {
  await installFakeFs(context);
  await page.goto('/');
  await openWithUnsavedEdit(page);
  await page.evaluate(() => {
    window.__fakeFs.externalEdit();
    // The handle lost access: the re-request needs a gesture and consumes it,
    // which is exactly why the picker failed after native confirms.
    window.__fakeFs.plan.query = 'prompt';
    window.__fakeFs.plan.request = 'granted';
  });
  await page.getByTestId('save-button').click();

  const dialog = page.getByRole('dialog', { name: 'File changed on disk' });
  await expect(dialog).toBeVisible();
  await dialog.getByRole('button', { name: 'Save As' }).click();

  await expect.poll(() => page.evaluate(() => window.__fakeFs.calls.close)).toBeGreaterThan(0);
  const calls = await page.evaluate(() => window.__fakeFs.calls);
  // The permission was re-asked first, and the picker still opened: from the
  // dialog button's own click, after the prompt had consumed the Save gesture.
  expect(calls.requestPermission.some((entry) => entry.mode === 'readwrite')).toBe(true);
  expect(calls.savePicker).toBe(1);
  expect(calls.createWritableHandles).toEqual(calls.savePickerHandles);
  await expect(page.getByTestId('status-message')).toContainText('Saved');
});

test('the save-as-instead dialog opens the picker from its own button', async ({
  page,
  context,
}) => {
  await installFakeFs(context);
  await page.goto('/');
  await openThroughFileMenu(page);
  await clickCell(page, 1, 1);
  await page.keyboard.type('43');
  await page.keyboard.press('Enter');

  await page.evaluate(() => {
    window.__fakeFs.plan.getFileThrows = true;
  });
  await page.getByTestId('save-button').click();

  const dialog = page.getByRole('dialog', { name: 'Save as a new file instead?' });
  await expect(dialog).toBeVisible();
  await dialog.getByRole('button', { name: 'Save as a new file' }).click();

  await expect.poll(() => page.evaluate(() => window.__fakeFs.calls.close)).toBeGreaterThan(0);
  const calls = await page.evaluate(() => window.__fakeFs.calls);
  expect(calls.savePicker).toBe(1);
  expect(calls.createWritableHandles).toEqual(calls.savePickerHandles);
  await expect(page.getByTestId('status-message')).toContainText('Saved');
});
