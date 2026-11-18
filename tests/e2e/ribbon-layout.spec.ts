import { expect, test } from '@playwright/test';
import { openFixture } from './grid-helpers';

const fixture = 'm1-chart-unknown.xlsx';

/**
 * The ribbon bands pack their controls into up to two rows. Every row starts
 * at the group's leftmost position (past a tall control like Paste), matching
 * the Excel layout the maintainer captured.
 */
test('each ribbon row starts at the left of its group', async ({ page }) => {
  await openFixture(page, fixture);

  // Font: row 1 starts with the font-family select, row 2 with Bold.
  const fontFamily = page.getByRole('combobox', { name: 'Font', exact: true });
  const bold = page.getByRole('button', { name: 'Bold (Ctrl+B)' });
  const familyBox = await fontFamily.boundingBox();
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
