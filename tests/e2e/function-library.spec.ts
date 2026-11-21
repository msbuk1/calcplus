import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';
import { clickCell, openFixture } from './grid-helpers';

const fixture = 'm1-chart-unknown.xlsx';

/**
 * The Formulas tab's Function Library group follows the Excel 365 oracle:
 * a tall Insert Function, AutoSum, then the eight category menus. Choosing a
 * function from a category menu takes the Insert Function dialog's own insert
 * path, not a second one.
 */

async function openFormulasTab(page: Page): Promise<void> {
  await openFixture(page, fixture);
  await clickCell(page, 1, 1);
  await page.getByRole('tab', { name: 'Formulas' }).click();
}

test('the ten Function Library controls render in the oracle order', async ({ page }) => {
  await openFormulasTab(page);
  const group = page.getByRole('group', { name: 'Function Library' });
  // The DOM packs the tall control first, then row 1, then row 2.
  const names = await group
    .getByRole('button')
    .evaluateAll((buttons) => buttons.map((button) => button.getAttribute('aria-label')));
  expect(names).toEqual([
    'Insert Function',
    'AutoSum (Alt+=)',
    'Recently Used',
    'Financial',
    'Logical',
    'Text',
    'Date & Time',
    'Lookup & Reference',
    'Math & Trig',
    'More Functions',
  ]);

  // Financial stays visible but greyed: the catalogue has no financial
  // functions yet, and the oracle still shows all ten controls.
  await expect(group.getByRole('button', { name: 'Financial', exact: true })).toBeDisabled();

  // No V2 auditing, name or calculation chrome joined the group.
  for (const name of [
    'Trace Precedents',
    'Trace Dependents',
    'Show Formulas',
    'Error Checking',
    'Evaluate Formula',
    'Watch Window',
    'Calculation Options',
    'Calculate Now',
    'Calculate Sheet',
  ]) {
    await expect(page.getByRole('button', { name, exact: true })).toHaveCount(0);
  }
  await expect(page.getByRole('group', { name: 'Defined Names' })).toHaveCount(0);
  await expect(page.getByRole('group', { name: 'Formula Auditing' })).toHaveCount(0);
  await expect(page.getByRole('group', { name: 'Calculation' })).toHaveCount(0);
});

test('a category menu inserts its function skeleton with the caret inside', async ({ page }) => {
  await openFormulasTab(page);
  const group = page.getByRole('group', { name: 'Function Library' });

  await group.getByRole('button', { name: 'Math & Trig', exact: true }).click();
  const menu = page.getByRole('menu', { name: 'Math & Trig options' });
  await expect(menu).toBeVisible();
  // The same alphabetised rows the Insert Function dialog shows for this
  // category; every row dispatches the existing insert Function command.
  await expect(menu.getByRole('menuitem')).toHaveText([
    'ABS',
    'INT',
    'MOD',
    'POWER',
    'PRODUCT',
    'ROUND',
    'ROUNDDOWN',
    'ROUNDUP',
    'SQRT',
    'SUM',
  ]);
  await menu.getByRole('menuitem', { name: 'SUM', exact: true }).click();

  const bar = page.getByTestId('formula-bar');
  await expect(bar).toHaveValue('=SUM(');
  await expect(bar).toBeFocused();
  expect(await bar.evaluate((element) => (element as HTMLInputElement).selectionStart)).toBe(5);
  expect(await bar.evaluate((element) => (element as HTMLInputElement).selectionEnd)).toBe(5);
  await expect
    .poll(() => page.evaluate(() => window.__calcplus.getState().focusZone))
    .toBe('formula-bar');
});

test('Recently Used fills as functions are inserted', async ({ page }) => {
  await openFormulasTab(page);
  const group = page.getByRole('group', { name: 'Function Library' });

  await group.getByRole('button', { name: 'Math & Trig', exact: true }).click();
  await page
    .getByRole('menu', { name: 'Math & Trig options' })
    .getByRole('menuitem', { name: 'SUM', exact: true })
    .click();

  await group.getByRole('button', { name: 'Recently Used', exact: true }).click();
  const menu = page.getByRole('menu', { name: 'Recently Used options' });
  await expect(menu.getByRole('menuitem')).toHaveText(['SUM']);
  await menu.getByRole('menuitem', { name: 'SUM', exact: true }).click();
  await expect(page.getByTestId('formula-bar')).toHaveValue('=SUM(');
});

test('More Functions opens the Insert Function dialog', async ({ page }) => {
  await openFormulasTab(page);
  const group = page.getByRole('group', { name: 'Function Library' });

  await group.getByRole('button', { name: 'More Functions', exact: true }).click();
  const dialog = page.getByRole('dialog', { name: 'Insert function' });
  await expect(dialog).toBeVisible();
  await expect(dialog.getByTestId('insert-function-search')).toBeFocused();
});
