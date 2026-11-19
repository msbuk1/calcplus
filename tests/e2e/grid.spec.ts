import { expect, test } from '@playwright/test';
import { cellRect, clickCell, dragSelect, openFixture, pixelAt } from './grid-helpers';

const gridFixture = 'm2-grid.xlsx';

test('selects a cell and mirrors the address in the name box', async ({ page }) => {
  await openFixture(page, gridFixture);
  await clickCell(page, 2, 1);
  await expect(page.getByTestId('active-cell')).toHaveText('B3');
  await expect(page.getByTestId('name-box')).toHaveValue('B3');
  await expect(page.getByTestId('formula-bar')).toHaveValue('0.25');
});

test('edits a cell in place, commits with Enter and cancels with Escape', async ({ page }) => {
  await openFixture(page, gridFixture);
  await clickCell(page, 2, 1);
  await page.keyboard.type('12.5');
  await expect(page.getByTestId('cell-editor')).toBeVisible();
  await expect(page.getByTestId('cell-editor')).toHaveValue('12.5');
  await page.keyboard.press('Enter');
  await expect(page.getByTestId('cell-editor')).toHaveCount(0);
  expect(await page.evaluate(() => window.__calcplus.getCellDisplay(2, 1))).toBe('12.50');
  await expect(page.getByTestId('active-cell')).toHaveText('B4');

  await clickCell(page, 2, 1);
  await page.keyboard.press('F2');
  await expect(page.getByTestId('cell-editor')).toHaveValue('12.5');
  await page.keyboard.press('Escape');
  await expect(page.getByTestId('cell-editor')).toHaveCount(0);
  expect(await page.evaluate(() => window.__calcplus.getCellDisplay(2, 1))).toBe('12.50');
});

test('commits an edit from the formula bar', async ({ page }) => {
  await openFixture(page, gridFixture);
  await clickCell(page, 2, 1);
  await page.getByTestId('formula-bar').click();
  await page.keyboard.press('Control+a');
  await page.keyboard.type('99');
  await page.keyboard.press('Enter');
  expect(await page.evaluate(() => window.__calcplus.getCellDisplay(2, 1))).toBe('99.00');
  await expect(page.getByTestId('dirty-indicator')).toBeVisible();
});

test('moves down after a formula-bar commit and stays on Ctrl+Enter', async ({ page }) => {
  await openFixture(page, gridFixture);
  await clickCell(page, 2, 1);
  await page.getByTestId('formula-bar').click();
  await page.keyboard.press('Control+a');
  await page.keyboard.type('5');
  await page.keyboard.press('Enter');
  await expect(page.getByTestId('active-cell')).toHaveText('B4');
  expect(await page.evaluate(() => window.__calcplus.getCellDisplay(2, 1))).toBe('5.00');

  await page.getByTestId('formula-bar').click();
  await page.keyboard.press('Control+a');
  await page.keyboard.type('6');
  await page.keyboard.press('Control+Enter');
  await expect(page.getByTestId('active-cell')).toHaveText('B4');
  expect(await page.evaluate(() => window.__calcplus.getCellDisplay(3, 1))).toBe('6.00');
});

test('navigates to an address from the name box', async ({ page }) => {
  await openFixture(page, gridFixture);
  await page.getByTestId('name-box').click();
  await page.keyboard.press('Control+a');
  await page.keyboard.type('B12');
  await page.keyboard.press('Enter');
  await expect(page.getByTestId('active-cell')).toHaveText('B12');
});

test('keeps the frozen row and column in place while scrolling', async ({ page }) => {
  await openFixture(page, gridFixture);
  expect(await page.evaluate(() => window.__calcplus.getFreeze())).toEqual({ rows: 1, cols: 1 });

  const frozenRowBefore = await cellRect(page, 0, 1);
  const frozenColumnBefore = await cellRect(page, 5, 0);
  const scrollingBefore = await cellRect(page, 5, 1);

  const box = await page.getByTestId('grid-canvas').boundingBox();
  expect(box).not.toBeNull();
  await page.mouse.move(box!.x + box!.width / 2, box!.y + box!.height / 2);
  await page.mouse.wheel(0, 600);
  await page.waitForTimeout(120);

  expect((await page.evaluate(() => window.__calcplus.getScroll())).y).toBeGreaterThan(0);
  // The frozen row keeps its place, the frozen column keeps its place across,
  // and the scrolling pane moved up.
  const frozenRowAfter = await cellRect(page, 0, 1);
  expect(frozenRowAfter.y).toBe(frozenRowBefore.y);
  const frozenColumnAfter = await cellRect(page, 5, 0);
  expect(frozenColumnAfter.x).toBe(frozenColumnBefore.x);
  const scrollingAfter = await cellRect(page, 5, 1);
  expect(scrollingAfter.y).toBeLessThan(scrollingBefore.y);
});

test('moves the active cell with the arrow keys and extends with Shift', async ({ page }) => {
  await openFixture(page, gridFixture);
  await clickCell(page, 2, 1);
  await page.keyboard.press('ArrowDown');
  await expect(page.getByTestId('active-cell')).toHaveText('B4');
  await page.keyboard.press('ArrowRight');
  await expect(page.getByTestId('active-cell')).toHaveText('C4');

  await page.keyboard.press('Shift+ArrowDown');
  const selection = await page.evaluate(() => window.__calcplus.getState().selection);
  expect(selection.anchor).toEqual({ row: 3, col: 2 });
  expect(selection.focus).toEqual({ row: 4, col: 2 });
  // The active cell stays where the selection began (ticket 31 A).
  await expect(page.getByTestId('active-cell')).toHaveText('C4');

  // A plain arrow collapses the range and moves from the active cell, not from
  // the end the Shift extension reached.
  await page.keyboard.press('ArrowRight');
  await expect(page.getByTestId('active-cell')).toHaveText('D4');
  await expect(page.getByTestId('name-box')).toHaveValue('D4');
});

test('a drag keeps its origin active and paints it unfilled inside the range', async ({ page }) => {
  await openFixture(page, gridFixture);
  // F10:G12 is empty in the fixture, so every sampled pixel is the paint alone.
  await dragSelect(page, 9, 5, 11, 6);
  await expect(page.getByTestId('name-box')).toHaveValue('F10:G12');

  const state = await page.evaluate(() => window.__calcplus.getState());
  expect(state.active).toEqual({ row: 9, col: 5 });
  await expect(page.getByTestId('active-cell')).toHaveText('F10');

  const origin = await cellRect(page, 9, 5);
  const other = await cellRect(page, 10, 6);
  // The active cell keeps the grid background; the rest of the range is tinted.
  const unfilled = await pixelAt(page, origin.x + origin.width / 2, origin.y + origin.height / 2);
  expect({ r: unfilled.r, g: unfilled.g, b: unfilled.b }).toEqual({ r: 255, g: 255, b: 255 });
  const tinted = await pixelAt(page, other.x + other.width / 2, other.y + other.height / 2);
  expect(tinted.r).toBeLessThan(255);
  expect(tinted.b).toBeGreaterThan(tinted.r);

  // No independent outline: the interior edge of the active cell is unpainted.
  const inside = await pixelAt(page, origin.x + origin.width - 2, origin.y + origin.height / 2);
  expect({ r: inside.r, g: inside.g, b: inside.b }).toEqual({ r: 255, g: 255, b: 255 });

  // The fill handle stays at the range corner, not at the active cell.
  const handle = await page.evaluate(() => window.__calcplus.getFillHandle());
  expect(handle).not.toBeNull();
  const corner = await cellRect(page, 11, 6);
  expect(handle!.x).toBeGreaterThan(corner.x + corner.width / 2);
  expect(handle!.y).toBeGreaterThan(corner.y + corner.height / 2);

  // Both headers of the selected columns shade for the full extent.
  for (const column of [origin, other]) {
    const headerPixel = await pixelAt(page, column.x + 20, 12);
    expect({ r: headerPixel.r, g: headerPixel.g, b: headerPixel.b }).toEqual({
      r: 26,
      g: 115,
      b: 232,
    });
  }
});

test('jumps to the edge of the data with Ctrl and the arrow keys', async ({ page }) => {
  await openFixture(page, gridFixture);
  await clickCell(page, 2, 1);
  await page.keyboard.press('Control+ArrowDown');
  await expect(page.getByTestId('active-cell')).toHaveText('B302');
  await page.keyboard.press('Control+Home');
  await expect(page.getByTestId('active-cell')).toHaveText('A1');
});

test('cycles the focus zones with F6 and returns to the grid with Escape', async ({ page }) => {
  await openFixture(page, gridFixture);
  await clickCell(page, 2, 1);
  expect((await page.evaluate(() => window.__calcplus.getState())).focusZone).toBe('grid');
  await expect(page.getByTestId('grid-shell')).toBeFocused();

  await page.keyboard.press('F6');
  expect((await page.evaluate(() => window.__calcplus.getState())).focusZone).toBe('ribbon');
  await expect(page.locator('.ribbon')).toBeFocused();
  await page.keyboard.press('F6');
  expect((await page.evaluate(() => window.__calcplus.getState())).focusZone).toBe('name-box');
  await expect(page.getByTestId('name-box')).toBeFocused();
  await page.keyboard.press('F6');
  expect((await page.evaluate(() => window.__calcplus.getState())).focusZone).toBe('formula-bar');
  await expect(page.getByTestId('formula-bar')).toBeFocused();
  await page.keyboard.press('F6');
  expect((await page.evaluate(() => window.__calcplus.getState())).focusZone).toBe('grid');
  await expect(page.getByTestId('grid-shell')).toBeFocused();

  await page.keyboard.press('F6');
  expect((await page.evaluate(() => window.__calcplus.getState())).focusZone).toBe('ribbon');
  await expect(page.locator('.ribbon')).toBeFocused();
  await page.keyboard.press('Escape');
  expect((await page.evaluate(() => window.__calcplus.getState())).focusZone).toBe('grid');
  await expect(page.getByTestId('grid-shell')).toBeFocused();
});

test('selects the anchor of a merged range', async ({ page }) => {
  await openFixture(page, gridFixture);
  await clickCell(page, 0, 1);
  await expect(page.getByTestId('active-cell')).toHaveText('A1');
  const style = await page.evaluate(() => window.__calcplus.getCellStyle(0, 1));
  expect(style?.merge).toEqual({ anchorRow: 0, anchorCol: 0, rowSpan: 1, colSpan: 2 });
});

test('renders the imported fill, alignment and number format', async ({ page }) => {
  await openFixture(page, gridFixture);
  const header = await page.evaluate(() => window.__calcplus.getCellStyle(1, 0));
  expect(header?.font.bold).toBe(true);
  expect(header?.fill.color).toBe('#FFF2CC');
  expect(header?.alignment.horizontal).toBe('center');

  const amount = await page.evaluate(() => window.__calcplus.getCellStyle(2, 1));
  expect(amount?.numberFormat).toBe('#,##0.00');
  expect(amount?.displayText).toBe('0.25');

  // The canvas paints the imported solid fill.
  const rect = await cellRect(page, 1, 0);
  const pixel = await pixelAt(page, rect.x + 4, rect.y + 4);
  expect(pixel.r).toBe(255);
  expect(pixel.g).toBe(242);
  expect(pixel.b).toBe(204);
});

test('scrolls a sparse sheet to the last row without allocating the grid', async ({ page }) => {
  await openFixture(page, gridFixture);
  const box = await page.getByTestId('grid-canvas').boundingBox();
  expect(box).not.toBeNull();
  await page.mouse.move(box!.x + box!.width / 2, box!.y + box!.height / 2);
  await page.mouse.wheel(0, 100_000);
  await page.waitForTimeout(200);
  const scroll = await page.evaluate(() => window.__calcplus.getScroll());
  expect(scroll.y).toBeGreaterThan(0);
});

test('paints the headers, the active cell and the fill handle', async ({ page }) => {
  await openFixture(page, gridFixture);
  await clickCell(page, 2, 1);
  const activeRect = await cellRect(page, 2, 1);

  // The column header of the active column carries the accent colour.
  const headerPixel = await pixelAt(page, activeRect.x + 20, 12);
  expect(headerPixel.r).toBe(26);
  expect(headerPixel.g).toBe(115);
  expect(headerPixel.b).toBe(232);

  // The fill handle sits on the bottom-right corner of the active cell.
  const handle = await page.evaluate(() => window.__calcplus.getFillHandle());
  expect(handle).not.toBeNull();
  const handlePixel = await pixelAt(page, (handle?.x ?? 0) + 3, (handle?.y ?? 0) + 3);
  expect(handlePixel.r).toBe(26);
  expect(handlePixel.g).toBe(115);
  expect(handlePixel.b).toBe(232);
});

test('commits the edit when another cell is clicked', async ({ page }) => {
  await openFixture(page, gridFixture);
  await clickCell(page, 2, 1);
  await page.keyboard.type('7');
  await expect(page.getByTestId('cell-editor')).toBeVisible();
  await clickCell(page, 3, 1);
  await expect(page.getByTestId('cell-editor')).toHaveCount(0);
  expect(await page.evaluate(() => window.__calcplus.getCellDisplay(2, 1))).toBe('7.00');
  await expect(page.getByTestId('active-cell')).toHaveText('B4');
});

test('drags the custom scrollbar thumb to scroll the sheet', async ({ page }) => {
  await openFixture(page, gridFixture);
  const thumb = page.getByTestId('grid-vscroll-thumb');
  await expect(thumb).toBeVisible();
  const before = await page.evaluate(() => window.__calcplus.getScroll());
  const thumbBox = await thumb.boundingBox();
  expect(thumbBox).not.toBeNull();
  const centreX = thumbBox!.x + thumbBox!.width / 2;
  const centreY = thumbBox!.y + thumbBox!.height / 2;
  await page.mouse.move(centreX, centreY);
  await page.mouse.down();
  await page.mouse.move(centreX, centreY + 120);
  await page.mouse.up();
  await page.waitForTimeout(100);
  const after = await page.evaluate(() => window.__calcplus.getScroll());
  expect(after.y).toBeGreaterThan(before.y);
});

test('does not commit an edit while an input method is composing', async ({ page }) => {
  await openFixture(page, gridFixture);
  await clickCell(page, 2, 1);
  await page.keyboard.type('4');
  await expect(page.getByTestId('cell-editor')).toBeVisible();

  const client = await page.context().newCDPSession(page);
  await client.send('Input.imeSetComposition', {
    text: 'ニ',
    selectionStart: 1,
    selectionEnd: 1,
  });
  await expect(page.getByTestId('cell-editor')).toHaveValue('4ニ');

  // Enter confirms the composition candidate; it must not commit the cell.
  await page.keyboard.press('Enter');
  await expect(page.getByTestId('cell-editor')).toBeVisible();
  expect(await page.evaluate(() => window.__calcplus.getCellDisplay(2, 1))).toBe('0.25');

  // Ending the composition and pressing Escape leaves the cell untouched.
  await client.send('Input.insertText', { text: '4ニ' });
  await page.keyboard.press('Escape');
  await expect(page.getByTestId('cell-editor')).toHaveCount(0);
  expect(await page.evaluate(() => window.__calcplus.getCellDisplay(2, 1))).toBe('0.25');
});
