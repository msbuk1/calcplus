import { readFile } from 'node:fs/promises';
import { expect, test } from '@playwright/test';
import { clickCell, openFixture } from './grid-helpers';

/**
 * M7 lane 3 — File → Export PDF end to end (print-pdf §2, §7).
 *
 * `print-areas.xlsx` is the A4 landscape fixture the preview spec uses, so the
 * PDF must have the same page count the preview reports. The download is
 * inspected from its bytes: Poppler is not assumed on every machine.
 */

test('File → Export PDF downloads a PDF with the preview’s page count', async ({ page }) => {
  await openFixture(page, 'print-areas.xlsx');

  // The preview states the plan the export must match.
  await page.getByTestId('file-menu-button').click();
  await page.getByRole('menuitem', { name: /^Print/ }).click();
  const dialog = page.getByRole('dialog', { name: 'Print Preview' });
  await expect(dialog).toBeVisible();
  await expect(dialog.getByTestId('print-preview-summary')).toContainText('A4 landscape');
  await expect(dialog.getByTestId('print-preview-summary')).toContainText('2 pages');
  await page.keyboard.press('Escape');
  await expect(dialog).toBeHidden();

  const [download] = await Promise.all([
    page.waitForEvent('download'),
    (async () => {
      await page.getByTestId('file-menu-button').click();
      await page.getByRole('menuitem', { name: /^Export PDF/ }).click();
    })(),
  ]);

  expect(download.suggestedFilename()).toBe('print-areas.pdf');
  const path = await download.path();
  expect(path).not.toBeNull();
  const bytes = await readFile(path!);
  const pdf = bytes.toString('latin1');

  expect(pdf.startsWith('%PDF-1.3')).toBe(true);
  expect(pdf.trimEnd().endsWith('%%EOF')).toBe(true);
  expect((pdf.match(/\/Type\s*\/Page(?![s])/g) ?? []).length).toBe(2);
  // Subsetted, renamed, embedded fonts: no Reserved Font Name ships.
  expect(pdf).toContain('/FontFile2');
  expect(pdf).toContain('/CALP+');
  expect(pdf).not.toMatch(/[A-Z]{6}\+(Selawik|Carlito|Liberation)/);

  await expect(page.getByTestId('status-message')).toContainText('Exported print-areas.pdf');
});

test('Export PDF follows the paper chosen in the preview', async ({ page }) => {
  await openFixture(page, 'print-areas.xlsx');

  // Choose Statement (612 × 396pt in landscape) in the preview, then close it.
  await page.getByTestId('file-menu-button').click();
  await page.getByRole('menuitem', { name: /^Print/ }).click();
  const dialog = page.getByRole('dialog', { name: 'Print Preview' });
  await expect(dialog).toBeVisible();
  await dialog.getByTestId('print-preview-paper').selectOption('6');
  await expect(dialog.getByTestId('print-preview-summary')).toContainText('Statement landscape');
  await page.keyboard.press('Escape');
  await expect(dialog).toBeHidden();

  const [download] = await Promise.all([
    page.waitForEvent('download'),
    (async () => {
      await page.getByTestId('file-menu-button').click();
      await page.getByRole('menuitem', { name: /^Export PDF/ }).click();
    })(),
  ]);

  const path = await download.path();
  expect(path).not.toBeNull();
  const pdf = (await readFile(path!)).toString('latin1');
  // The PDF's page box is the chosen paper, not the A4 the file declared.
  expect(pdf).toContain('/MediaBox [0 0 612 396]');
  expect(pdf.match(/\/Type\s*\/Page(?![s])/g) ?? []).toHaveLength(3);
});

test('a CJK page is rasterised at 300 DPI in the Chromium fallback path', async ({ page }) => {
  // A blank workbook whose only text cannot embed in any bundled face.
  await page.goto('/?fsa=0');
  await page.getByTestId('file-menu-button').click();
  await page.getByRole('menuitem', { name: /^New$/ }).click();
  await expect(page.getByTestId('file-name')).toHaveText('Untitled.xlsx');

  await clickCell(page, 0, 0);
  await page.keyboard.press('F2');
  await expect(page.getByTestId('cell-editor')).toBeVisible();
  await page.keyboard.insertText('汉字测试');
  await page.keyboard.press('Enter');

  const [download] = await Promise.all([
    page.waitForEvent('download'),
    (async () => {
      await page.getByTestId('file-menu-button').click();
      await page.getByRole('menuitem', { name: /^Export PDF/ }).click();
    })(),
  ]);
  expect(download.suggestedFilename()).toBe('Untitled.pdf');

  const path = await download.path();
  const pdf = (await readFile(path!)).toString('latin1');
  // A full-page 300 DPI PNG: one image XObject covering the Letter page.
  expect(pdf).toContain('/Subtype /Image');
  expect(pdf).toMatch(/\/Width 2550/);
  await expect(page.getByTestId('status-message')).toContainText('rasterised at 300 DPI');
});
