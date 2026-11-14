import { readFileSync } from 'node:fs';
import { Buffer } from 'node:buffer';
import { fileURLToPath } from 'node:url';
import { expect, test } from '@playwright/test';
import { openWorkbook } from '../../src/ooxml/open';

const fixturePath = new URL('../../fixtures/compat/m1-chart-unknown.xlsx', import.meta.url);
const fixtureBytes = new Uint8Array(readFileSync(fixturePath));

async function openFixture(page: import('@playwright/test').Page): Promise<void> {
  await page.goto('/?fsa=0');
  await page.getByTestId('file-input').setInputFiles(fileURLToPath(fixturePath));
  await expect(page.getByTestId('file-name')).toHaveText('m1-chart-unknown.xlsx');
  await expect(page.getByTestId('grid-canvas')).toBeVisible();
}

async function readDownload(
  download: import('@playwright/test').Download,
  outputPath: string,
): Promise<Uint8Array> {
  await download.saveAs(outputPath);
  return new Uint8Array(readFileSync(outputPath));
}

test('opens a workbook, renders values, edits, commits and undoes', async ({ page }) => {
  await openFixture(page);
  await expect(page.getByTestId('compat-banner')).toContainText('Charts are preserved');

  const before = await page.evaluate(() =>
    (
      window as unknown as { __calcplus: { getCellDisplay(r: number, c: number): string } }
    ).__calcplus.getCellDisplay(1, 1),
  );
  expect(before).toBe('42.5');

  const canvas = page.getByTestId('grid-canvas');
  const box = await canvas.boundingBox();
  expect(box).not.toBeNull();
  await page.mouse.click(box!.x + 48 + 104 + 20, box!.y + 24 + 22 + 10);
  await expect(page.getByTestId('active-cell')).toHaveText('B2');

  await page.keyboard.type('43');
  await expect(page.getByTestId('cell-editor')).toBeVisible();
  await page.keyboard.press('Enter');
  await expect(page.getByTestId('dirty-indicator')).toBeVisible();

  const afterEdit = await page.evaluate(() =>
    (
      window as unknown as { __calcplus: { getCellDisplay(r: number, c: number): string } }
    ).__calcplus.getCellDisplay(1, 1),
  );
  expect(afterEdit).toBe('43');

  await page.getByTestId('undo-button').click();
  const afterUndo = await page.evaluate(() =>
    (
      window as unknown as { __calcplus: { getCellDisplay(r: number, c: number): string } }
    ).__calcplus.getCellDisplay(1, 1),
  );
  expect(afterUndo).toBe('42.5');
});

test('no-op save downloads the original bytes exactly', async ({ page }, testInfo) => {
  await openFixture(page);
  const downloadPromise = page.waitForEvent('download');
  await page.getByTestId('save-button').click();
  const saved = await readDownload(await downloadPromise, testInfo.outputPath('identity.xlsx'));
  expect(Buffer.compare(Buffer.from(saved), Buffer.from(fixtureBytes))).toBe(0);
});

test('edited save downloads a workbook that preserves Tier B parts', async ({ page }, testInfo) => {
  await openFixture(page);
  const canvas = page.getByTestId('grid-canvas');
  const box = await canvas.boundingBox();
  await page.mouse.click(box!.x + 48 + 104 + 20, box!.y + 24 + 22 + 10);
  await page.keyboard.type('43');
  await page.keyboard.press('Enter');
  await expect(page.getByTestId('dirty-indicator')).toBeVisible();

  const downloadPromise = page.waitForEvent('download');
  await page.getByTestId('save-button').click();
  const saved = await readDownload(await downloadPromise, testInfo.outputPath('edited.xlsx'));
  const reopened = openWorkbook(saved);
  expect(reopened.workbook.sheets[0]!.cells.get('1:1')?.value).toEqual({
    type: 'number',
    value: 43,
  });
  expect(
    Buffer.compare(
      Buffer.from(reopened.workbook.opaqueParts.get('xl/charts/chart1.xml')!.bytes),
      Buffer.from(fixtureBytes.length > 0 ? readChart(fixtureBytes) : new Uint8Array()),
    ),
  ).toBe(0);
});

function readChart(bytes: Uint8Array): Uint8Array {
  return openWorkbook(bytes).workbook.opaqueParts.get('xl/charts/chart1.xml')!.bytes;
}

test('creates a new blank workbook and saves it through Save As', async ({ page }, testInfo) => {
  await page.goto('/?fsa=0');
  await page.getByTestId('new-button').click();
  await expect(page.getByTestId('file-name')).toHaveText('Untitled.xlsx');
  await expect(page.getByTestId('grid-canvas')).toBeVisible();

  const canvas = page.getByTestId('grid-canvas');
  const box = await canvas.boundingBox();
  expect(box).not.toBeNull();
  await page.mouse.click(box!.x + 48 + 20, box!.y + 24 + 10);
  await page.keyboard.type('Hello');
  await page.keyboard.press('Enter');
  await expect(page.getByTestId('dirty-indicator')).toBeVisible();

  const downloadPromise = page.waitForEvent('download');
  await page.getByTestId('save-as-button').click();
  const saved = await readDownload(await downloadPromise, testInfo.outputPath('untitled.xlsx'));
  const reopened = openWorkbook(saved);
  expect(reopened.workbook.sheets).toHaveLength(1);
  expect(reopened.workbook.sheets[0]!.name).toBe('Sheet1');
  expect(reopened.workbook.sheets[0]!.cells.get('0:0')?.value).toEqual({
    type: 'text',
    value: 'Hello',
  });
});

test('F2 and Enter without a change does not rewrite a formula cell', async ({ page }) => {
  await openFixture(page);
  const canvas = page.getByTestId('grid-canvas');
  const box = await canvas.boundingBox();
  expect(box).not.toBeNull();
  // B3 is the formula cell B2*2 with cached value 85.
  await page.mouse.click(box!.x + 48 + 104 + 20, box!.y + 24 + 22 + 22 + 10);
  await expect(page.getByTestId('active-cell')).toHaveText('B3');
  await page.keyboard.press('F2');
  await expect(page.getByTestId('cell-editor')).toHaveValue('85');
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
  await openFixture(page);
  const canvas = page.getByTestId('grid-canvas');
  const box = await canvas.boundingBox();
  expect(box).not.toBeNull();
  await page.mouse.click(box!.x + 48 + 20, box!.y + 24 + 10);
  await page.keyboard.type('changed');
  await page.keyboard.press('Enter');
  await expect(page.getByTestId('dirty-indicator')).toBeVisible();

  const dialogPromise = page.waitForEvent('dialog');
  await page.getByTestId('file-input').setInputFiles(fileURLToPath(fixturePath));
  const dialog = await dialogPromise;
  expect(dialog.message()).toMatch(/unsaved changes/i);
  await dialog.dismiss();

  await expect(page.getByTestId('status-message')).toContainText('Unsaved changes were kept');
  await expect(page.getByTestId('file-name')).toHaveText('m1-chart-unknown.xlsx');
  await expect(page.getByTestId('dirty-indicator')).toBeVisible();
});
