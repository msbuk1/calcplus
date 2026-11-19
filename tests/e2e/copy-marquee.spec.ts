import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';
import { cellRect, clickCell, openFixture } from './grid-helpers';

/**
 * Ticket 31 B: after Ctrl+C or Ctrl+X the dashed marquee marks the copied
 * range and is the paste destination. The arrow keys move it (the selection
 * does not move), Enter pastes there in one transaction and selects the pasted
 * range, Esc ends the mode without pasting, and any other key or pointer that
 * moves the selection dismisses it first.
 *
 * The workbook is fixtures/compat/m2-grid.xlsx: B3 holds 0.25 and B4 holds
 * 37.25, with C3 the percentage share of the same row.
 */

const fixture = 'm2-grid.xlsx';

async function cellDisplay(page: Page, row: number, col: number): Promise<string | null> {
  return page.evaluate(
    ([targetRow, targetCol]) =>
      window.__calcplus.getCellDisplay(targetRow as number, targetCol as number),
    [row, col] as const,
  );
}

async function version(page: Page): Promise<number> {
  return page.evaluate(() => window.__calcplus.getState().version);
}

/** Assert the marquee covers exactly the rectangle between two cells. */
async function expectMarquee(
  page: Page,
  start: readonly [number, number],
  end: readonly [number, number],
): Promise<void> {
  const marquee = page.getByTestId('copy-marquee');
  await expect(marquee).toBeVisible();
  const canvas = await page.getByTestId('grid-canvas').boundingBox();
  expect(canvas).not.toBeNull();
  const topLeft = await cellRect(page, start[0], start[1]);
  const bottomRight = await cellRect(page, end[0], end[1]);
  const box = await marquee.boundingBox();
  expect(box).not.toBeNull();
  expect(box!.x).toBeCloseTo(canvas!.x + topLeft.x, 0);
  expect(box!.y).toBeCloseTo(canvas!.y + topLeft.y, 0);
  expect(box!.width).toBeCloseTo(bottomRight.x + bottomRight.width - topLeft.x, 0);
  expect(box!.height).toBeCloseTo(bottomRight.y + bottomRight.height - topLeft.y, 0);
}

test('Ctrl+C marks the range and the arrows move the marquee, then Enter pastes once', async ({
  page,
}) => {
  await openFixture(page, fixture);
  await clickCell(page, 2, 1);
  await page.keyboard.press('Control+c');
  await expectMarquee(page, [2, 1], [2, 1]);

  await page.keyboard.press('ArrowRight');
  await expectMarquee(page, [2, 2], [2, 2]);
  await page.keyboard.press('ArrowDown');
  await expectMarquee(page, [3, 2], [3, 2]);

  // The arrows moved the marquee only: the selection is still the source.
  await expect(page.getByTestId('name-box')).toHaveValue('B3');
  await expect(page.getByTestId('active-cell')).toHaveText('B3');
  expect(await cellDisplay(page, 3, 2)).toBe('13.00%');

  const before = await version(page);
  await page.keyboard.press('Enter');

  // One transaction pasted the snapshot at C4 and selected the pasted range.
  expect(await version(page)).toBe(before + 1);
  await expect.poll(() => cellDisplay(page, 3, 2)).toBe('0.25');
  await expect(page.getByTestId('name-box')).toHaveValue('C4');
  await expect(page.getByTestId('active-cell')).toHaveText('C4');
  await expect(page.getByTestId('copy-marquee')).toHaveCount(0);

  // One undo restores the cell the paste overwrote.
  await page.keyboard.press('Control+z');
  await expect.poll(() => cellDisplay(page, 3, 2)).toBe('13.00%');
});

test('the marquee clamps at the sheet edge instead of wrapping', async ({ page }) => {
  await openFixture(page, fixture);
  // A1 is the first cell: no row above it and no column before it.
  await clickCell(page, 0, 0);
  await page.keyboard.press('Control+c');
  await expectMarquee(page, [0, 0], [0, 0]);

  await page.keyboard.press('ArrowUp');
  await expectMarquee(page, [0, 0], [0, 0]);
  await page.keyboard.press('ArrowLeft');
  await expectMarquee(page, [0, 0], [0, 0]);

  // A step that stays inside the sheet still moves the marquee.
  await page.keyboard.press('ArrowRight');
  await expectMarquee(page, [0, 1], [0, 1]);
  await page.keyboard.press('ArrowDown');
  await expectMarquee(page, [1, 1], [1, 1]);
  await expect(page.getByTestId('name-box')).toHaveValue('A1');
});

test('Esc ends copy mode without pasting', async ({ page }) => {
  await openFixture(page, fixture);
  await clickCell(page, 2, 1);
  await page.keyboard.press('Control+c');
  await page.keyboard.press('ArrowDown');
  await expectMarquee(page, [3, 1], [3, 1]);

  await page.keyboard.press('Escape');
  await expect(page.getByTestId('copy-marquee')).toHaveCount(0);
  expect(await cellDisplay(page, 3, 1)).toBe('37.25');
  await expect(page.getByTestId('name-box')).toHaveValue('B3');
});

test('a pointer that moves the selection dismisses the marquee', async ({ page }) => {
  await openFixture(page, fixture);
  await clickCell(page, 2, 1);
  await page.keyboard.press('Control+c');
  await page.keyboard.press('ArrowDown');
  await expectMarquee(page, [3, 1], [3, 1]);

  await clickCell(page, 5, 1);
  await expect(page.getByTestId('copy-marquee')).toHaveCount(0);
  await expect(page.getByTestId('name-box')).toHaveValue('B6');
  // The marquee's destination was never written.
  expect(await cellDisplay(page, 3, 1)).toBe('37.25');
});

test('another key dismisses copy mode and then does its own work', async ({ page }) => {
  await openFixture(page, fixture);
  await clickCell(page, 2, 1);
  await page.keyboard.press('Control+c');
  await page.keyboard.press('ArrowDown');
  await expectMarquee(page, [3, 1], [3, 1]);

  await page.keyboard.press('Tab');
  await expect(page.getByTestId('copy-marquee')).toHaveCount(0);
  await expect(page.getByTestId('name-box')).toHaveValue('C3');
  expect(await cellDisplay(page, 3, 1)).toBe('37.25');
});

test('Ctrl+V pastes at the moved marquee', async ({ page }) => {
  await openFixture(page, fixture);
  await clickCell(page, 2, 1);
  await page.keyboard.press('Control+c');
  await page.keyboard.press('ArrowRight');
  await expectMarquee(page, [2, 2], [2, 2]);

  await page.keyboard.press('Control+v');
  await expect.poll(() => cellDisplay(page, 2, 2)).toBe('0.25');
  await expect(page.getByTestId('name-box')).toHaveValue('C3');
  await expect(page.getByTestId('copy-marquee')).toHaveCount(0);
});

test('the marquee stays over its cells while the sheet scrolls', async ({ page }) => {
  await openFixture(page, fixture);
  await clickCell(page, 9, 1);
  await page.keyboard.press('Control+c');
  await expectMarquee(page, [9, 1], [9, 1]);

  const canvas = await page.getByTestId('grid-canvas').boundingBox();
  expect(canvas).not.toBeNull();
  await page.mouse.move(canvas!.x + canvas!.width / 2, canvas!.y + canvas!.height / 2);
  await page.mouse.wheel(0, 120);
  await page.waitForTimeout(150);

  expect((await page.evaluate(() => window.__calcplus.getScroll())).y).toBeGreaterThan(0);
  await expectMarquee(page, [9, 1], [9, 1]);
});

test('cut clears the source in the paste transaction and one undo restores both', async ({
  page,
}) => {
  await openFixture(page, fixture);
  await clickCell(page, 2, 1);
  await page.keyboard.press('Control+x');
  await expectMarquee(page, [2, 1], [2, 1]);
  await page.keyboard.press('ArrowDown');
  await expectMarquee(page, [3, 1], [3, 1]);

  const before = await version(page);
  await page.keyboard.press('Enter');
  expect(await version(page)).toBe(before + 1);
  await expect.poll(() => cellDisplay(page, 3, 1)).toBe('0.25');
  expect(await cellDisplay(page, 2, 1)).toBeNull();
  await expect(page.getByTestId('name-box')).toHaveValue('B4');

  // One undo restores the source and the cell the move overwrote.
  await page.keyboard.press('Control+z');
  await expect.poll(() => cellDisplay(page, 2, 1)).toBe('0.25');
  await expect.poll(() => cellDisplay(page, 3, 1)).toBe('37.25');
});
