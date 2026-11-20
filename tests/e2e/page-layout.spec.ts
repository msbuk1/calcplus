import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';
import { dragSelect, openFixture } from './grid-helpers';

/**
 * Ticket 44 — the Page Layout tab's Page Setup group and the completed Sheet
 * Options group.
 *
 * The assertions are structural: the map's groups, the dialog's write model
 * (one transaction per Apply, undo restores it), the live plan line, and the
 * print area's effect on the preview. No snapshots.
 */

async function openPageLayout(page: Page): Promise<void> {
  await page.getByRole('tab', { name: 'Page Layout' }).click();
}

async function openPageSetup(page: Page): Promise<void> {
  await page.getByRole('button', { name: 'Page Setup', exact: true }).click();
  await expect(page.getByRole('dialog', { name: 'Page Setup' })).toBeVisible();
}

async function openPrintPreview(page: Page): Promise<void> {
  await page.getByTestId('file-menu-button').click();
  await page.getByRole('menuitem', { name: /^Print/ }).click();
  await expect(page.getByRole('dialog', { name: 'Print Preview' })).toBeVisible();
}

async function closePrintPreview(page: Page): Promise<void> {
  await page.keyboard.press('Escape');
  await expect(page.getByRole('dialog', { name: 'Print Preview' })).toBeHidden();
}

test('the Page Layout tab shows Page Setup and Sheet Options, and no v2 groups', async ({
  page,
}) => {
  await page.addInitScript(() => {
    window.print = () => undefined;
  });
  await openFixture(page, 'print-setup.xlsx');
  await openPageLayout(page);

  // The oracle's v1 groups, in order. Breaks, Themes, Background, Print
  // Titles and Scale to Fit are hidden: they have no placement.
  const groups = page.locator('.ribbon-group');
  await expect(groups).toHaveCount(2);
  await expect(page.getByRole('group', { name: 'Page Setup' })).toBeVisible();
  await expect(page.getByRole('group', { name: 'Sheet Options' })).toBeVisible();
  for (const hidden of ['Themes', 'Scale to Fit', 'Breaks']) {
    await expect(page.getByRole('group', { name: hidden })).toHaveCount(0);
  }

  // Page Setup: the dialog button and the Print Area drop-down.
  await expect(page.getByRole('button', { name: 'Page Setup', exact: true })).toBeVisible();
  await expect(page.getByRole('button', { name: 'Print Area', exact: true })).toBeVisible();

  // Sheet Options: two rows, each with View and Print.
  await expect(page.getByRole('button', { name: 'Gridlines: View' })).toBeVisible();
  await expect(page.getByRole('button', { name: 'Gridlines: Print' })).toBeVisible();
  await expect(page.getByRole('button', { name: 'Headings: View' })).toBeVisible();
  await expect(page.getByRole('button', { name: 'Headings: Print' })).toBeVisible();
  // print-setup.xlsx turns both printed flags off.
  await expect(page.getByRole('button', { name: 'Gridlines: Print' })).toHaveAttribute(
    'aria-pressed',
    'false',
  );
  await expect(page.getByRole('button', { name: 'Headings: Print' })).toHaveAttribute(
    'aria-pressed',
    'false',
  );
});

test('the Page Setup dialog writes one transaction and its plan line follows the draft', async ({
  page,
}) => {
  await page.addInitScript(() => {
    window.print = () => undefined;
  });
  await openFixture(page, 'print-setup.xlsx');
  await openPageLayout(page);
  const before = await page.evaluate(() => window.__calcplus.getState());
  expect(before.canUndo).toBe(false);

  await openPageSetup(page);

  // The dialog opens on the sheet's model values and shows the current plan.
  await expect(page.getByTestId('page-setup-landscape')).toBeChecked();
  await expect(page.getByTestId('page-setup-plan')).toContainText('Plan: 1 page on A4 landscape.');

  // Switching to portrait re-plans the draft live, before anything is written.
  await page.getByTestId('page-setup-portrait').check();
  await expect(page.getByTestId('page-setup-plan')).toContainText('A4 portrait');
  expect(await page.evaluate(() => window.__calcplus.getState().canUndo)).toBe(false);

  // The margins tab: a large top and bottom margin adds a page.
  await page.getByRole('tab', { name: 'Margins' }).click();
  await page.getByTestId('page-setup-margin-top').fill('4.5');
  await page.getByTestId('page-setup-margin-bottom').fill('4.5');
  await expect(page.getByTestId('page-setup-plan')).toContainText('Plan: 2 pages on A4 portrait.');

  // OK is one transaction, whatever the number of changed fields.
  await page.getByRole('button', { name: 'OK', exact: true }).click();
  await expect(page.getByRole('dialog', { name: 'Page Setup' })).toBeHidden();
  const after = await page.evaluate(() => window.__calcplus.getState());
  expect(after.canUndo).toBe(true);
  expect(after.version).toBe(before.version + 1);

  // Reopening shows the written values.
  await openPageSetup(page);
  await expect(page.getByTestId('page-setup-portrait')).toBeChecked();
  await page.getByRole('tab', { name: 'Margins' }).click();
  await expect(page.getByTestId('page-setup-margin-top')).toHaveValue('4.5');
  await page.getByRole('button', { name: 'Cancel', exact: true }).click();

  // Undo restores the exact prior page setup in one step.
  await page.keyboard.press('Control+z');
  await openPageSetup(page);
  await expect(page.getByTestId('page-setup-landscape')).toBeChecked();
  await page.getByRole('tab', { name: 'Margins' }).click();
  await expect(page.getByTestId('page-setup-margin-top')).toHaveValue('0.75');
  await page.getByRole('button', { name: 'Cancel', exact: true }).click();
});

test('Set and Clear Print Area change the preview immediately', async ({ page }) => {
  await page.addInitScript(() => {
    window.print = () => undefined;
  });
  await openFixture(page, 'print-areas.xlsx');
  await dragSelect(page, 0, 0, 9, 2);
  await openPageLayout(page);

  await page.getByRole('button', { name: 'Print Area', exact: true }).click();
  await page.getByRole('menuitem', { name: 'Set Print Area' }).click();
  await page.getByRole('button', { name: 'Print Area', exact: true }).click();

  await openPrintPreview(page);
  await expect(page.getByTestId('print-preview-summary')).toContainText('1 page');
  await expect(page.getByTestId('print-preview-summary')).not.toContainText('2 pages');
  await closePrintPreview(page);

  // Clearing removes the name, so the plan falls back to the used range
  // (still one page), and the command is one undoable transaction.
  await page.getByRole('button', { name: 'Print Area', exact: true }).click();
  await page.getByRole('menuitem', { name: 'Clear Print Area' }).click();
  expect(await page.evaluate(() => window.__calcplus.getState().canUndo)).toBe(true);
  await openPrintPreview(page);
  await expect(page.getByTestId('print-preview-summary')).toContainText('1 page');
  await closePrintPreview(page);

  // Undo the clear and then the set: the imported two-area plan returns, so
  // both writes restored the exact prior name.
  await page.keyboard.press('Control+z');
  await page.keyboard.press('Control+z');
  await openPrintPreview(page);
  await expect(page.getByTestId('print-preview-summary')).toContainText('2 pages');
  await closePrintPreview(page);
});

test('the Print check box writes printed state only, and undo restores it', async ({ page }) => {
  await page.addInitScript(() => {
    window.print = () => undefined;
  });
  await openFixture(page, 'print-setup.xlsx');
  await openPageLayout(page);

  const printGridlines = page.getByRole('button', { name: 'Gridlines: Print' });
  const viewGridlines = page.getByRole('button', { name: 'Gridlines: View' });
  const viewBefore = await viewGridlines.getAttribute('aria-pressed');

  await printGridlines.click();
  await expect(printGridlines).toHaveAttribute('aria-pressed', 'true');
  // The screen toggle is a different command; its state did not move.
  await expect(viewGridlines).toHaveAttribute('aria-pressed', viewBefore ?? 'false');
  expect(await page.evaluate(() => window.__calcplus.getState().canUndo)).toBe(true);

  // Undo is one step and returns the printed flag exactly.
  await page.keyboard.press('Control+z');
  await expect(printGridlines).toHaveAttribute('aria-pressed', 'false');
  expect(await page.evaluate(() => window.__calcplus.getState().canUndo)).toBe(false);
  await expect(viewGridlines).toHaveAttribute('aria-pressed', viewBefore ?? 'false');
});

test('the preview opens Page Setup over itself and the pages re-plan', async ({ page }) => {
  await page.addInitScript(() => {
    window.print = () => undefined;
  });
  await openFixture(page, 'print-setup.xlsx');

  await openPrintPreview(page);
  await expect(page.getByTestId('print-preview-summary')).toContainText('1 page');
  await expect(page.getByTestId('print-preview-summary')).toContainText('A4 landscape');

  // The preview offers the dialog and stays mounted underneath it.
  await page.getByTestId('print-preview-page-setup').click();
  const setup = page.getByRole('dialog', { name: 'Page Setup' });
  await expect(setup).toBeVisible();
  // The dialog is above the preview, not behind it (same backdrop stacking).
  const box = await setup.boundingBox();
  expect(box).not.toBeNull();
  const topmost = await page.evaluate(
    ([x, y]) =>
      document
        .elementFromPoint(x as number, y as number)
        ?.closest('[role="dialog"]')
        ?.querySelector('.dialog-title')?.textContent ?? null,
    [box!.x + box!.width / 2, box!.y + 20] as const,
  );
  expect(topmost).toBe('Page Setup');

  await page.getByTestId('page-setup-portrait').check();
  await page.getByRole('tab', { name: 'Margins' }).click();
  await page.getByTestId('page-setup-margin-top').fill('4.5');
  await page.getByTestId('page-setup-margin-bottom').fill('4.5');
  await expect(page.getByTestId('page-setup-plan')).toContainText('Plan: 2 pages on A4 portrait.');

  // Apply writes one transaction, and the preview behind re-plans while the
  // dialog is still open.
  await page.getByTestId('page-setup-apply').click();
  await expect(setup).toBeVisible();
  await expect(page.getByTestId('print-preview-summary')).toContainText('A4 portrait');
  await expect(page.getByTestId('print-preview-summary')).toContainText('2 pages');

  // Esc closes only the top dialog; the preview stays and the keyboard returns
  // to the button that opened the dialog.
  await page.keyboard.press('Escape');
  await expect(setup).toBeHidden();
  await expect(page.getByRole('dialog', { name: 'Print Preview' })).toBeVisible();
  await expect(page.getByTestId('print-preview-page-setup')).toBeFocused();

  await closePrintPreview(page);
});

test('Add to Print Area appends to the existing areas, and one undo restores them', async ({
  page,
}) => {
  await page.addInitScript(() => {
    window.print = () => undefined;
  });
  await openFixture(page, 'print-areas.xlsx');

  // The file has two areas; add a third from the grid selection.
  await dragSelect(page, 9, 4, 11, 5);
  await openPageLayout(page);
  await page.getByRole('button', { name: 'Print Area', exact: true }).click();
  await page.getByRole('menuitem', { name: 'Add to Print Area' }).click();

  await openPrintPreview(page);
  await expect(page.getByTestId('print-preview-summary')).toContainText('3 pages');
  await closePrintPreview(page);

  // Add is one name transaction; undo restores the imported two areas.
  expect(await page.evaluate(() => window.__calcplus.getState().canUndo)).toBe(true);
  await page.keyboard.press('Control+z');
  await openPrintPreview(page);
  await expect(page.getByTestId('print-preview-summary')).toContainText('2 pages');
  await closePrintPreview(page);
});

test('the Margins tab centre check boxes are one transaction and round-trip', async ({ page }) => {
  await page.addInitScript(() => {
    window.print = () => undefined;
  });
  await openFixture(page, 'print-areas.xlsx');
  await openPageLayout(page);
  const before = await page.evaluate(() => window.__calcplus.getState());

  await openPageSetup(page);
  await page.getByRole('tab', { name: 'Margins' }).click();
  await expect(page.getByTestId('page-setup-centre-horizontal')).not.toBeChecked();
  await page.getByTestId('page-setup-centre-horizontal').check();
  await page.getByTestId('page-setup-centre-vertical').check();
  await page.getByRole('button', { name: 'OK', exact: true }).click();

  const after = await page.evaluate(() => window.__calcplus.getState());
  expect(after.version).toBe(before.version + 1);
  expect(after.canUndo).toBe(true);

  // Reopening reads the written flags.
  await openPageSetup(page);
  await page.getByRole('tab', { name: 'Margins' }).click();
  await expect(page.getByTestId('page-setup-centre-horizontal')).toBeChecked();
  await expect(page.getByTestId('page-setup-centre-vertical')).toBeChecked();
  await page.getByRole('button', { name: 'Cancel', exact: true }).click();

  // One undo restores both flags.
  await page.keyboard.press('Control+z');
  await openPageSetup(page);
  await page.getByRole('tab', { name: 'Margins' }).click();
  await expect(page.getByTestId('page-setup-centre-horizontal')).not.toBeChecked();
  await expect(page.getByTestId('page-setup-centre-vertical')).not.toBeChecked();
  await page.getByRole('button', { name: 'Cancel', exact: true }).click();
});
