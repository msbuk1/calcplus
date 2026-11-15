import { Buffer } from 'node:buffer';
import { readFileSync } from 'node:fs';
import { expect, test } from '@playwright/test';
import { openWorkbook } from '../../src/ooxml/open';
import { clickCell, fixtureBytes, fixturePath, openFixture } from './grid-helpers';

const fixtureName = 'm1-chart-unknown.xlsx';

test('opens a workbook, renders values, edits, commits and undoes', async ({ page }) => {
  await openFixture(page, fixtureName);
  await expect(page.getByTestId('compat-banner')).toContainText('Charts are preserved');

  const before = await page.evaluate(() => window.__calcplus.getCellDisplay(1, 1));
  expect(before).toBe('42.5');

  await clickCell(page, 1, 1);
  await expect(page.getByTestId('active-cell')).toHaveText('B2');

  await page.keyboard.type('43');
  await expect(page.getByTestId('cell-editor')).toBeVisible();
  await page.keyboard.press('Enter');
  await expect(page.getByTestId('dirty-indicator')).toBeVisible();

  const afterEdit = await page.evaluate(() => window.__calcplus.getCellDisplay(1, 1));
  expect(afterEdit).toBe('43');

  await page.getByTestId('undo-button').click();
  const afterUndo = await page.evaluate(() => window.__calcplus.getCellDisplay(1, 1));
  expect(afterUndo).toBe('42.5');
});

test('no-op save downloads the original bytes exactly', async ({ page }, testInfo) => {
  await openFixture(page, fixtureName);
  const downloadPromise = page.waitForEvent('download');
  await page.getByTestId('save-button').click();
  const download = await downloadPromise;
  await download.saveAs(testInfo.outputPath('identity.xlsx'));
  const saved = new Uint8Array(readFileSync(testInfo.outputPath('identity.xlsx')));
  expect(Buffer.compare(Buffer.from(saved), Buffer.from(fixtureBytes(fixtureName)))).toBe(0);
});

test('edited save downloads a workbook that preserves Tier B parts', async ({ page }, testInfo) => {
  await openFixture(page, fixtureName);
  await clickCell(page, 1, 1);
  await page.keyboard.type('43');
  await page.keyboard.press('Enter');
  await expect(page.getByTestId('dirty-indicator')).toBeVisible();

  const downloadPromise = page.waitForEvent('download');
  await page.getByTestId('save-button').click();
  const download = await downloadPromise;
  await download.saveAs(testInfo.outputPath('edited.xlsx'));
  const saved = new Uint8Array(readFileSync(testInfo.outputPath('edited.xlsx')));
  const reopened = openWorkbook(saved);
  expect(reopened.workbook.sheets[0]!.cells.get('1:1')?.value).toEqual({
    type: 'number',
    value: 43,
  });
  expect(
    Buffer.compare(
      Buffer.from(reopened.workbook.opaqueParts.get('xl/charts/chart1.xml')!.bytes),
      Buffer.from(
        openWorkbook(fixtureBytes(fixtureName)).workbook.opaqueParts.get('xl/charts/chart1.xml')!
          .bytes,
      ),
    ),
  ).toBe(0);
});

test('creates a new blank workbook and saves it through Save As', async ({ page }, testInfo) => {
  await page.goto('/?fsa=0');
  await page.getByTestId('new-button').click();
  await expect(page.getByTestId('file-name')).toHaveText('Untitled.xlsx');
  await expect(page.getByTestId('grid-canvas')).toBeVisible();

  await clickCell(page, 0, 0);
  await page.keyboard.type('Hello');
  await page.keyboard.press('Enter');
  await expect(page.getByTestId('dirty-indicator')).toBeVisible();

  const downloadPromise = page.waitForEvent('download');
  await page.getByTestId('save-as-button').click();
  const download = await downloadPromise;
  await download.saveAs(testInfo.outputPath('untitled.xlsx'));
  const saved = new Uint8Array(readFileSync(testInfo.outputPath('untitled.xlsx')));
  const reopened = openWorkbook(saved);
  expect(reopened.workbook.sheets).toHaveLength(1);
  expect(reopened.workbook.sheets[0]!.name).toBe('Sheet1');
  expect(reopened.workbook.sheets[0]!.cells.get('0:0')?.value).toEqual({
    type: 'text',
    value: 'Hello',
  });
});

test('F2 and Enter without a change does not rewrite a formula cell', async ({ page }) => {
  await openFixture(page, fixtureName);
  // B3 is the formula cell B2*2 with cached value 85.
  await clickCell(page, 2, 1);
  await expect(page.getByTestId('active-cell')).toHaveText('B3');
  await page.keyboard.press('F2');
  await expect(page.getByTestId('cell-editor')).toHaveValue('B2*2');
  await page.keyboard.press('Enter');
  await expect(page.getByTestId('dirty-indicator')).toHaveCount(0);
  const formula = await page.evaluate(
    () =>
      (
        window as unknown as {
          __calcplus: { getCell(r: number, c: number): { formula: string | null } | null };
        }
      ).__calcplus.getCell(2, 1)?.formula,
  );
  expect(formula).toBe('B2*2');
});

test('prompts and keeps the current document when opening with unsaved edits', async ({ page }) => {
  await openFixture(page, fixtureName);
  await clickCell(page, 0, 0);
  await page.keyboard.type('changed');
  await page.keyboard.press('Enter');
  await expect(page.getByTestId('dirty-indicator')).toBeVisible();

  const dialogPromise = page.waitForEvent('dialog');
  await page.getByTestId('file-input').setInputFiles(fixturePath(fixtureName));
  const dialog = await dialogPromise;
  expect(dialog.message()).toMatch(/unsaved changes/i);
  await dialog.dismiss();

  await expect(page.getByTestId('status-message')).toContainText('Unsaved changes were kept');
  await expect(page.getByTestId('file-name')).toHaveText(fixtureName);
  await expect(page.getByTestId('dirty-indicator')).toBeVisible();
});
