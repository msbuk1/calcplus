import { expect, test } from '@playwright/test';
import {
  cellRect,
  clickCell,
  focusedTestId,
  openFixture,
  pixelAt,
  regionSignature,
} from './grid-helpers';

const gridFixture = 'm2-grid.xlsx';

test('moves with Tab and Enter and keeps the focus in the grid', async ({ page }) => {
  await openFixture(page, gridFixture);
  await clickCell(page, 2, 1);
  await expect(page.getByTestId('active-cell')).toHaveText('B3');

  await page.keyboard.press('Tab');
  await expect(page.getByTestId('active-cell')).toHaveText('C3');
  expect(await focusedTestId(page)).toBe('grid-shell');

  await page.keyboard.press('Enter');
  await expect(page.getByTestId('active-cell')).toHaveText('C4');
  expect(await focusedTestId(page)).toBe('grid-shell');

  await page.keyboard.press('Shift+Tab');
  await expect(page.getByTestId('active-cell')).toHaveText('B4');
  expect(await focusedTestId(page)).toBe('grid-shell');
});

test('selects the whole column from the column header', async ({ page }) => {
  await openFixture(page, gridFixture);
  const box = await page.getByTestId('grid-canvas').boundingBox();
  expect(box).not.toBeNull();
  const target = await cellRect(page, 2, 1);
  // The column header strip sits above the cells.
  await page.mouse.click(box!.x + target.x + target.width / 2, box!.y + 12);

  // The visible cells of the column carry the selection tint.
  const tinted = await pixelAt(page, target.x + 10, target.y + 5);
  expect(tinted.r).toBeLessThan(255);
  expect(tinted.b).toBeGreaterThan(tinted.r);

  // The active cell stays near the top: the view did not scroll to the far
  // edge of the sheet.
  await expect(page.getByTestId('name-box')).toHaveValue('B:B');
  expect(await focusedTestId(page)).toBe('grid-shell');
});

test('selects the whole row from the row header', async ({ page }) => {
  await openFixture(page, gridFixture);
  const box = await page.getByTestId('grid-canvas').boundingBox();
  expect(box).not.toBeNull();
  const target = await cellRect(page, 2, 1);
  // The row header strip sits to the left of the cells.
  await page.mouse.click(box!.x + 20, box!.y + target.y + target.height / 2);

  // The visible cells of the row carry the selection tint. B3 is the active
  // cell of the row selection — it stays unfilled (ticket 31 A) — so the
  // sample sits in C3 and the active cell is asserted on its own.
  const tinted = await pixelAt(page, target.x + target.width + 10, target.y + 5);
  expect(tinted.r).toBeLessThan(255);
  expect(tinted.b).toBeGreaterThan(tinted.r);
  const activeCellPixel = await pixelAt(page, target.x + 10, target.y + 5);
  expect({ r: activeCellPixel.r, g: activeCellPixel.g, b: activeCellPixel.b }).toEqual({
    r: 255,
    g: 255,
    b: 255,
  });

  await expect(page.getByTestId('name-box')).toHaveValue('3:3');
  expect(await focusedTestId(page)).toBe('grid-shell');
});

test('hands the focus back to the grid after the name box', async ({ page }) => {
  await openFixture(page, gridFixture);
  await page.getByTestId('name-box').click();
  expect(await focusedTestId(page)).toBe('name-box');
  await page.keyboard.press('Control+a');
  await page.keyboard.type('B12');
  await page.keyboard.press('Enter');
  await expect(page.getByTestId('active-cell')).toHaveText('B12');
  expect(await focusedTestId(page)).toBe('grid-shell');
});

test('keeps a frozen row painted while the rows below scroll', async ({ page }) => {
  await openFixture(page, 'm2-freeze-row.xlsx');
  const frozenRow = await cellRect(page, 0, 1);
  const scrollingRow = await cellRect(page, 3, 1);
  const frozenBefore = await regionSignature(page, frozenRow.x, frozenRow.y, 200, 20);
  const scrollingBefore = await regionSignature(page, scrollingRow.x, scrollingRow.y, 200, 20);

  const box = await page.getByTestId('grid-canvas').boundingBox();
  expect(box).not.toBeNull();
  await page.mouse.move(box!.x + box!.width / 2, box!.y + box!.height / 2);
  await page.mouse.wheel(0, 400);
  await page.waitForTimeout(150);

  // The frozen row still shows the same pixels; the rows below moved.
  expect(await regionSignature(page, frozenRow.x, frozenRow.y, 200, 20)).toBe(frozenBefore);
  expect(await regionSignature(page, scrollingRow.x, scrollingRow.y, 200, 20)).not.toBe(
    scrollingBefore,
  );
});

test('keeps a frozen column painted while the columns beside it scroll', async ({ page }) => {
  await openFixture(page, 'm2-freeze-column.xlsx');
  const frozenColumn = await cellRect(page, 2, 0);
  const scrollingColumn = await cellRect(page, 2, 1);
  const frozenBefore = await regionSignature(page, frozenColumn.x, frozenColumn.y, 60, 20);
  // Sample the whole scrolling column, including its right-aligned value, so
  // a blank left edge cannot make the before and after regions look equal.
  const scrollingBefore = await regionSignature(
    page,
    scrollingColumn.x,
    scrollingColumn.y,
    scrollingColumn.width,
    20,
  );

  const box = await page.getByTestId('grid-canvas').boundingBox();
  expect(box).not.toBeNull();
  await page.mouse.move(box!.x + box!.width / 2, box!.y + box!.height / 2);
  await page.mouse.wheel(600, 0);
  await page.waitForTimeout(150);

  // The frozen column still shows the same pixels; the columns beside it moved.
  expect(await regionSignature(page, frozenColumn.x, frozenColumn.y, 60, 20)).toBe(frozenBefore);
  expect(
    await regionSignature(page, scrollingColumn.x, scrollingColumn.y, scrollingColumn.width, 20),
  ).not.toBe(scrollingBefore);
});

test('edits a merged range through its anchor', async ({ page }) => {
  await openFixture(page, gridFixture);
  // B1 is covered by the merged range A1:B1.
  await clickCell(page, 0, 1);
  await expect(page.getByTestId('active-cell')).toHaveText('A1');

  await page.keyboard.press('F2');
  await expect(page.getByTestId('cell-editor')).toBeVisible();
  const editor = await page.getByTestId('cell-editor').boundingBox();
  const anchor = await cellRect(page, 0, 0);
  expect(editor).not.toBeNull();
  // The editor covers the whole merged range, not one covered cell.
  expect(editor!.width).toBeGreaterThan(anchor.width);
  await page.keyboard.press('Escape');
});

test('opens the editor when an input method starts on the selected cell', async ({ page }) => {
  await openFixture(page, gridFixture);
  await clickCell(page, 2, 1);
  await expect(page.getByTestId('cell-editor')).toHaveCount(0);

  // Chromium dispatches compositionstart only to an editable focused host, so
  // the shell itself is the editable target that starts the composition.
  const client = await page.context().newCDPSession(page);
  await client.send('Input.imeSetComposition', {
    text: 'ニ',
    selectionStart: 1,
    selectionEnd: 1,
  });
  // The composition opens the editor on the selected cell.
  await expect(page.getByTestId('cell-editor')).toBeVisible();
  await client.send('Input.insertText', { text: 'ニ' });
  await expect(page.getByTestId('cell-editor')).toHaveValue('ニ');

  await page.keyboard.press('Enter');
  await expect(page.getByTestId('cell-editor')).toHaveCount(0);
  expect(await page.evaluate(() => window.__calcplus.getCellDisplay(2, 1))).toBe('ニ');
});

test('reaches the last row of the sheet and shows it', async ({ page }) => {
  await openFixture(page, gridFixture);
  await page.getByTestId('name-box').click();
  await page.keyboard.press('Control+a');
  await page.keyboard.type('A1048576');
  await page.keyboard.press('Enter');
  await expect(page.getByTestId('name-box')).toHaveValue('A1048576');
  await page.waitForTimeout(200);

  // The row header strip paints a label for the row that is on screen.
  const headerPixels = await regionSignature(page, 4, 26, 40, 16);
  expect(headerPixels).not.toBe('');
  const painted = await page.evaluate(() => {
    const canvas = document.querySelector<HTMLCanvasElement>('[data-testid="grid-canvas"]');
    if (canvas === null) return 0;
    const context = canvas.getContext('2d');
    if (context === null) return 0;
    const ratio = canvas.width / canvas.getBoundingClientRect().width;
    const data = context.getImageData(
      Math.round(4 * ratio),
      Math.round(26 * ratio),
      Math.round(40 * ratio),
      Math.round(16 * ratio),
    ).data;
    let dark = 0;
    for (let index = 0; index < data.length; index += 4) {
      if (data[index] < 200 && data[index + 1] < 200 && data[index + 2] < 200) dark += 1;
    }
    return dark;
  });
  expect(painted).toBeGreaterThan(0);
});
