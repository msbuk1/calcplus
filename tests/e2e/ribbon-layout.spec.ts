import { expect, test } from '@playwright/test';
import type { Locator, Page } from '@playwright/test';
import { openFixture } from './grid-helpers';

const fixture = 'm1-chart-unknown.xlsx';

async function openFormulas(page: Page, width: number): Promise<void> {
  await page.setViewportSize({ width, height: 900 });
  await openFixture(page, fixture);
  await page.getByRole('tab', { name: 'Formulas' }).click();
}

interface Box {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

interface ItemBox {
  readonly button: Box;
  readonly label: Box | null;
  readonly caret: Box | null;
}

/** The nine drop-downs' button, label and caret boxes in DOM order. */
async function itemBoxes(group: Locator): Promise<ItemBox[]> {
  return group.locator('.ribbon-library-item button').evaluateAll((buttons) => {
    const rect = (element: Element | null): Box | null => {
      if (element === null) return null;
      const box = element.getBoundingClientRect();
      return { x: box.x, y: box.y, width: box.width, height: box.height };
    };
    return buttons.map((button) => ({
      button: rect(button)!,
      label: rect(button.querySelector('.ribbon-label')),
      caret: rect(button.querySelector('.ribbon-menu-caret')),
    }));
  });
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
 * Formulas › Function Library, WIDE state (the approved prototype): the nine
 * drop-downs sit in one row at one height, and each caret sits under its
 * label.
 */
test('wide: the nine drop-downs form one row at equal height, caret under the label', async ({
  page,
}) => {
  await openFormulas(page, 1440);
  const group = page.locator('[data-group="Function Library"]');
  const items = await itemBoxes(group);
  expect(items).toHaveLength(9);

  for (const [index, item] of items.entries()) {
    expect(item.button.height, `button ${index}`).toBeCloseTo(items[0]!.button.height, 1);
    expect(item.button.y, `button ${index}`).toBeCloseTo(items[0]!.button.y, 1);
  }
  // One row: the nine run left to right without stacking.
  for (let index = 1; index < items.length; index += 1) {
    expect(items[index]!.button.x, `button ${index}`).toBeGreaterThan(items[index - 1]!.button.x);
  }
  // The caret sits under the label, not beside it.
  for (const [index, item] of items.entries()) {
    expect(item.caret, `caret ${index}`).not.toBeNull();
    expect(item.label, `label ${index}`).not.toBeNull();
    expect(item.caret!.y, `caret ${index}`).toBeGreaterThanOrEqual(
      item.label!.y + item.label!.height - 1,
    );
  }
});

/**
 * NARROW state: the nine drop-downs stack in three columns by three rows with
 * aligned column x values, and each caret hugs the end of its text.
 */
test('narrow: the nine drop-downs land on a 3 x 3 grid with aligned columns, caret beside the text', async ({
  page,
}) => {
  await openFormulas(page, 920);
  const group = page.locator('[data-group="Function Library"]');
  const items = await itemBoxes(group);
  expect(items).toHaveLength(9);

  // Row-major DOM order: three rows of three. Rows share y, columns share x.
  for (let row = 0; row < 3; row += 1) {
    for (let column = 1; column < 3; column += 1) {
      expect(items[row * 3 + column]!.button.y, `row ${row} column ${column}`).toBeCloseTo(
        items[row * 3]!.button.y,
        1,
      );
      expect(items[row * 3 + column]!.button.height, `row ${row} column ${column}`).toBeCloseTo(
        items[0]!.button.height,
        1,
      );
    }
    if (row > 0) {
      expect(items[row * 3]!.button.y, `row ${row}`).toBeGreaterThan(items[0]!.button.y);
    }
  }
  for (let column = 0; column < 3; column += 1) {
    for (let row = 1; row < 3; row += 1) {
      expect(items[row * 3 + column]!.button.x, `row ${row} column ${column}`).toBeCloseTo(
        items[column]!.button.x,
        1,
      );
    }
  }

  // The caret hugs the end of the text and shares its line.
  for (const [index, item] of items.entries()) {
    expect(item.caret, `caret ${index}`).not.toBeNull();
    expect(item.label, `label ${index}`).not.toBeNull();
    const textRight = item.label!.x + item.label!.width;
    expect(item.caret!.x, `caret ${index}`).toBeGreaterThanOrEqual(textRight - 0.5);
    expect(item.caret!.x - textRight, `caret ${index}`).toBeLessThanOrEqual(8);
    expect(item.caret!.y, `caret ${index}`).toBeLessThan(item.label!.y + item.label!.height);
    expect(item.caret!.y + item.caret!.height, `caret ${index}`).toBeGreaterThan(item.label!.y);
  }
});

/** The hero is the tall Insert Function item, and the only one without a caret. */
test('the hero is tall and never carries a caret', async ({ page }) => {
  await openFormulas(page, 1440);
  const group = page.locator('[data-group="Function Library"]');
  const hero = group.locator('.ribbon-library-hero button');

  for (const width of [1440, 920]) {
    await page.setViewportSize({ width, height: 900 });
    const heroBox = await hero.boundingBox();
    const items = await itemBoxes(group);
    expect(heroBox).not.toBeNull();
    expect(items).toHaveLength(9);
    expect(heroBox!.height, `${width}px`).toBeGreaterThan(items[0]!.button.height);
    await expect(hero.locator('.ribbon-menu-caret'), `${width}px`).toHaveCount(0);
    // The nine all carry one.
    for (const [index, item] of items.entries()) {
      expect(item.caret, `${width}px caret ${index}`).not.toBeNull();
    }
  }
});

/**
 * The group label is centred at the bottom of the group in both states. The
 * ui-chrome §5 rule hides ribbon group labels below 1100 px, whatever state
 * the group is in (the maintainer ruled the contract wins over the prototype).
 */
test('the Function Library label shows in the full band and hides below 1100 px', async ({
  page,
}) => {
  await openFormulas(page, 1440);
  const group = page.locator('[data-group="Function Library"]');
  const label = group.locator('.ribbon-group-label');
  await expect(label).toBeVisible();

  await page.setViewportSize({ width: 1000, height: 900 });
  await expect(label).toBeHidden();

  await page.setViewportSize({ width: 920, height: 900 });
  await expect(label).toBeHidden();
});
