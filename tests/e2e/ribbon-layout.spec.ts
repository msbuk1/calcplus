import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';
import { openFixture } from './grid-helpers';

const fixture = 'm1-chart-unknown.xlsx';

async function openFormulas(page: Page, width: number): Promise<void> {
  await page.setViewportSize({ width, height: 900 });
  await openFixture(page, fixture);
  await page.getByRole('tab', { name: 'Formulas' }).click();
}

/**
 * The ribbon bands pack their controls into up to two rows. Every row starts
 * at the group's leftmost position (past a tall control like Paste), matching
 * the Excel layout the maintainer captured.
 */
test('each ribbon row starts at the left of its group', async ({ page }) => {
  await openFixture(page, fixture);

  // Font: row 1 starts with the font-family control, row 2 with Bold. The
  // control wrapper is the row start; the select sits beside its icon.
  const fontFamily = page.getByRole('combobox', { name: 'Font', exact: true });
  const familyControl = page.locator('.ribbon-select').filter({ has: fontFamily });
  const bold = page.getByRole('button', { name: 'Bold (Ctrl+B)' });
  const familyBox = await familyControl.boundingBox();
  const boldBox = await bold.boundingBox();
  expect(familyBox).not.toBeNull();
  expect(boldBox).not.toBeNull();
  expect(Math.abs(boldBox!.x - familyBox!.x)).toBeLessThanOrEqual(4);
  // Row 1's wide font select must not widen row 2's first slot.
  expect(boldBox!.width).toBeLessThan(familyBox!.width / 2);

  // Alignment: row 1 starts with Align Left, row 2 with Align Top.
  const alignLeft = page.getByRole('button', { name: 'Align Left' });
  const alignTop = page.getByRole('button', { name: 'Align Top' });
  const leftBox = await alignLeft.boundingBox();
  const topBox = await alignTop.boundingBox();
  expect(leftBox).not.toBeNull();
  expect(topBox).not.toBeNull();
  expect(Math.abs(topBox!.x - leftBox!.x)).toBeLessThanOrEqual(4);
});

/**
 * The Formulas › Function Library tiles follow the Excel 365 capture: one
 * fixed width per tile, two clean columns under the tall Insert Function hero,
 * and a drop-down caret that rides the label's last line.
 */
test('Function Library tiles share one width, line up in columns and carry their carets', async ({
  page,
}) => {
  await openFormulas(page, 1440);
  const group = page.locator('[data-group="Function Library"]');
  const cells = group.locator('.ribbon-cell');

  for (const { width, tileWidth } of [
    { width: 1440, tileWidth: 60 },
    { width: 1000, tileWidth: 52 },
  ]) {
    await page.setViewportSize({ width, height: 900 });
    await expect(cells).toHaveCount(10);

    // Cell 1 is the tall Insert Function hero. The other nine are the two rows,
    // in DOM order: AutoSum, Recently Used, Financial, Logical, Text, then the
    // four second-row tiles.
    const tiles = await cells.evaluateAll((elements) =>
      elements.slice(1).map((element) => {
        const rect = element.getBoundingClientRect();
        return { x: rect.x, y: rect.y, width: rect.width };
      }),
    );
    expect(tiles).toHaveLength(9);
    expect(tiles[0]!.width, `${width}px`).toBe(tileWidth);
    for (const tile of tiles) expect(tile.width).toBeCloseTo(tileWidth, 1);
    // The second row repeats the first four columns instead of flowing freely.
    for (let column = 0; column < 4; column += 1) {
      expect(tiles[5 + column]!.x, `${width}px column ${column + 1}`).toBeCloseTo(
        tiles[column]!.x,
        1,
      );
    }

    for (const name of [
      'Recently Used',
      'Financial',
      'Logical',
      'Text',
      'Date & Time',
      'Lookup & Reference',
      'Math & Trig',
    ]) {
      const caret = group.getByRole('button', { name, exact: true }).locator('.ribbon-menu-caret');
      await expect(caret).toHaveCount(1);
      const placement = await caret.evaluate((element) => {
        const caretRect = element.getBoundingClientRect();
        const label = element.closest('.ribbon-label');
        if (label === null) throw new Error('The caret is outside the label');
        const labelRect = label.getBoundingClientRect();
        const line = element.parentElement;
        if (line === null || line.firstChild === null) {
          throw new Error('The caret has no text line');
        }
        const textRange = document.createRange();
        textRange.selectNodeContents(line.firstChild);
        const textRect = textRange.getBoundingClientRect();
        const iconRect = element
          .closest('.ribbon-button')!
          .querySelector('.ribbon-icon')!
          .getBoundingClientRect();
        return {
          position: getComputedStyle(element).position,
          insideLabel:
            caretRect.left >= labelRect.left - 1 &&
            caretRect.right <= labelRect.right + 2 &&
            caretRect.top >= labelRect.top - 2 &&
            caretRect.top <= labelRect.bottom,
          besideText:
            caretRect.left >= textRect.right - 0.5 && caretRect.left - textRect.right <= 3,
          belowIcon: caretRect.top >= iconRect.bottom - 2,
        };
      });
      expect(placement.position, name).not.toBe('absolute');
      expect(placement.insideLabel, `${width}px ${name}`).toBe(true);
      expect(placement.besideText, `${width}px ${name}`).toBe(true);
      expect(placement.belowIcon, `${width}px ${name}`).toBe(true);
    }
  }
});

test('the Function Library group label keeps air beneath the text at every full-ribbon width', async ({
  page,
}) => {
  await openFormulas(page, 1440);
  const group = page.locator('[data-group="Function Library"]');

  for (const width of [1100, 1280, 1440]) {
    await page.setViewportSize({ width, height: 900 });
    const air = await group.evaluate((element) => {
      const label = element.querySelector('.ribbon-group-label');
      if (label === null) throw new Error('No group label');
      const textRange = document.createRange();
      textRange.selectNodeContents(label);
      const textBottom = textRange.getBoundingClientRect().bottom;
      const groupRect = element.getBoundingClientRect();
      const bandRect = element.closest('.ribbon-body')!.getBoundingClientRect();
      return {
        display: getComputedStyle(label).display,
        groupAir: groupRect.bottom - textBottom,
        bandAir: bandRect.bottom - textBottom,
      };
    });
    expect(air.display, `${width}px`).not.toBe('none');
    expect(air.groupAir, `${width}px`).toBeGreaterThanOrEqual(3);
    expect(air.bandAir, `${width}px`).toBeGreaterThanOrEqual(3);
  }
});
