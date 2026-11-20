import { Buffer } from 'node:buffer';
import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';
import { buildXlsx } from '../helpers/build-xlsx';
import { cellRect } from './grid-helpers';

/**
 * Edge auto-scroll while a pointer drag draws a selection.
 *
 * The workbook is built at test time and has the full sparse sheet geometry,
 * so a drag can travel far past any edge without a large fixture. Only the
 * first two rows carry data; the fill drag at the end relies on the empty
 * rows below as its destination.
 */

const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;

function autoscrollWorkbookXlsx(frozenRows = 0): Uint8Array {
  const pane =
    frozenRows === 0
      ? ''
      : `<pane ySplit="${frozenRows}" topLeftCell="A${frozenRows + 1}" activePane="bottomLeft" state="frozen"/>`;
  const worksheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:B80"/>
  <sheetViews><sheetView tabSelected="1" workbookViewId="0">${pane}</sheetView></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Row</t></is></c><c r="B1" t="inlineStr"><is><t>Value</t></is></c></row>
    <row r="2"><c r="A2"><v>1</v></c><c r="B2"><v>2</v></c></row>
  </sheetData>
</worksheet>`;
  return buildXlsx(worksheetXml, { stylesXml });
}

async function openGrid(page: Page, frozenRows = 0): Promise<void> {
  await page.goto('/?fsa=0');
  await page.getByTestId('file-input').setInputFiles({
    name: 'autoscroll.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: Buffer.from(autoscrollWorkbookXlsx(frozenRows)),
  });
  await expect(page.getByTestId('file-name')).toHaveText('autoscroll.xlsx');
  await expect(page.getByTestId('grid-canvas')).toBeVisible();
}

async function selectCell(page: Page, address: string): Promise<void> {
  const nameBox = page.getByTestId('name-box');
  await nameBox.click();
  await page.keyboard.press('Control+a');
  await page.keyboard.type(address);
  await page.keyboard.press('Enter');
  await expect(nameBox).toHaveValue(address);
}

async function canvasBox(page: Page) {
  const box = await page.getByTestId('grid-canvas').boundingBox();
  if (box === null) throw new Error('The grid canvas has no box');
  return box;
}

async function scrollOffset(page: Page) {
  return page.evaluate(() => window.__calcplus.getScroll());
}

async function selectionState(page: Page) {
  return page.evaluate(() => window.__calcplus.getState().selection);
}

/** The rectangle a range selection covers, or `null` for a plain cell selection. */
async function selectionRangeEnd(page: Page) {
  return page.evaluate(() => {
    const selection = window.__calcplus.getState().selection as unknown as {
      extent:
        | { kind: 'cells' }
        | {
            kind: 'range';
            range: { startRow: number; startCol: number; endRow: number; endCol: number };
          };
    };
    return selection.extent.kind === 'range' ? selection.extent.range : null;
  });
}

async function cellDisplay(page: Page, row: number, col: number): Promise<string | null> {
  return page.evaluate(
    ([targetRow, targetCol]) =>
      window.__calcplus.getCellDisplay(targetRow as number, targetCol as number),
    [row, col] as const,
  );
}

/** Press in the middle of one visible cell and leave the button down. */
async function pressCell(page: Page, row: number, col: number) {
  const box = await canvasBox(page);
  const rect = await cellRect(page, row, col);
  const x = box.x + rect.x + rect.width / 2;
  const y = box.y + rect.y + rect.height / 2;
  await page.mouse.move(x, y);
  await page.mouse.down();
  return { box, x, y };
}

/** Page coordinates of the painted fill handle for the active cell. */
async function fillHandlePoint(page: Page, row: number, col: number) {
  const isHandleOfActiveCell = async (): Promise<boolean> => {
    const handle = await page.evaluate(() => window.__calcplus.getFillHandle());
    if (handle === null) return false;
    const rect = await cellRect(page, row, col);
    const centreX = handle.x + handle.width / 2;
    const centreY = handle.y + handle.height / 2;
    return (
      centreX >= rect.x &&
      centreX <= rect.x + rect.width &&
      centreY >= rect.y &&
      centreY <= rect.y + rect.height
    );
  };
  await expect.poll(isHandleOfActiveCell).toBe(true);
  const handle = await page.evaluate(() => window.__calcplus.getFillHandle());
  if (handle === null) throw new Error('The fill handle is not painted');
  const box = await canvasBox(page);
  return {
    x: box.x + handle.x + handle.width / 2,
    y: box.y + handle.y + handle.height / 2,
  };
}

test('dragging above the top edge scrolls up, extends and keeps going while held', async ({
  page,
}) => {
  await openGrid(page);
  await selectCell(page, 'A40');
  const before = await scrollOffset(page);
  expect(before.y).toBeGreaterThan(0);

  const { box, x } = await pressCell(page, 39, 0);
  // The column header is above the first cell row, so the pointer is past the
  // top edge of the grid body.
  await page.mouse.move(x, box.y + 2, { steps: 4 });
  await page.waitForTimeout(150);

  const early = await scrollOffset(page);
  expect(early.y).toBeLessThan(before.y);
  const earlySelection = await selectionState(page);
  expect(earlySelection.anchor).toEqual({ row: 39, col: 0 });
  expect(earlySelection.focus.row).toBeLessThan(39);

  // No further mouse movement: the loop must keep scrolling on its own.
  await page.waitForTimeout(400);
  const later = await scrollOffset(page);
  expect(later.y).toBeLessThan(early.y);
  const laterSelection = await selectionState(page);
  expect(laterSelection.focus.row).toBeLessThan(earlySelection.focus.row);

  await page.mouse.up();
  await page.waitForTimeout(100);
  const released = await scrollOffset(page);
  await page.waitForTimeout(300);
  // The release stopped the loop.
  expect(await scrollOffset(page)).toEqual(released);
});

test('dragging past the bottom edge scrolls down and extends the range', async ({ page }) => {
  await openGrid(page);
  await selectCell(page, 'A3');
  const before = await scrollOffset(page);

  const { box, x } = await pressCell(page, 2, 0);
  await page.mouse.move(x, box.y + box.height - 2, { steps: 4 });
  await page.waitForTimeout(400);

  const after = await scrollOffset(page);
  expect(after.y).toBeGreaterThan(before.y);
  const selection = await selectionState(page);
  expect(selection.anchor).toEqual({ row: 2, col: 0 });
  expect(selection.focus.row).toBeGreaterThan(2);
  await page.mouse.up();
});

test('dragging past the right edge scrolls right and extends the range', async ({ page }) => {
  await openGrid(page);
  await selectCell(page, 'A3');
  const before = await scrollOffset(page);

  const { box, y } = await pressCell(page, 2, 0);
  await page.mouse.move(box.x + box.width - 2, y, { steps: 4 });
  await page.waitForTimeout(400);

  const after = await scrollOffset(page);
  expect(after.x).toBeGreaterThan(before.x);
  const selection = await selectionState(page);
  expect(selection.anchor).toEqual({ row: 2, col: 0 });
  expect(selection.focus.col).toBeGreaterThan(0);
  await page.mouse.up();
});

test('dragging past the left edge scrolls left and extends the range', async ({ page }) => {
  await openGrid(page);
  await selectCell(page, 'Z3');
  const before = await scrollOffset(page);
  expect(before.x).toBeGreaterThan(0);

  const { box, y } = await pressCell(page, 2, 25);
  await page.mouse.move(box.x + 2, y, { steps: 4 });
  await page.waitForTimeout(400);

  const after = await scrollOffset(page);
  expect(after.x).toBeLessThan(before.x);
  const selection = await selectionState(page);
  expect(selection.anchor).toEqual({ row: 2, col: 25 });
  expect(selection.focus.col).toBeLessThan(25);
  await page.mouse.up();
});

test('edge scrolling clamps at the first cell and never wraps', async ({ page }) => {
  await openGrid(page);
  await selectCell(page, 'A1');
  expect(await scrollOffset(page)).toEqual({ x: 0, y: 0 });

  const { box } = await pressCell(page, 0, 0);
  await page.mouse.move(box.x + 2, box.y + 2, { steps: 4 });
  await page.waitForTimeout(350);

  expect(await scrollOffset(page)).toEqual({ x: 0, y: 0 });
  const selection = await selectionState(page);
  expect(selection.anchor).toEqual({ row: 0, col: 0 });
  expect(selection.focus).toEqual({ row: 0, col: 0 });
  await page.mouse.up();
});

test('edge scrolling clamps at the last cell and never wraps', async ({ page }) => {
  await openGrid(page);
  await selectCell(page, 'XFD1048576');
  const before = await scrollOffset(page);
  expect(before.x).toBeGreaterThan(0);
  expect(before.y).toBeGreaterThan(0);

  const { box } = await pressCell(page, 1_048_575, 16_383);
  await page.mouse.move(box.x + box.width - 2, box.y + box.height - 2, { steps: 4 });
  await page.waitForTimeout(350);

  expect(await scrollOffset(page)).toEqual(before);
  const selection = await selectionState(page);
  expect(selection.anchor).toEqual({ row: 1_048_575, col: 16_383 });
  expect(selection.focus).toEqual({ row: 1_048_575, col: 16_383 });
  await page.mouse.up();
});

test('a frozen row stays put while the rows below scroll', async ({ page }) => {
  await openGrid(page, 1);
  expect(await page.evaluate(() => window.__calcplus.getFreeze())).toEqual({ rows: 1, cols: 0 });
  await selectCell(page, 'A40');

  const frozen = await cellRect(page, 0, 1);
  const before = await scrollOffset(page);
  expect(before.y).toBeGreaterThan(0);

  const { box, x } = await pressCell(page, 39, 0);
  await page.mouse.move(x, box.y + 2, { steps: 4 });
  await page.waitForTimeout(400);

  const after = await scrollOffset(page);
  expect(after.y).toBeLessThan(before.y);
  // The frozen row keeps its painted position while the body scrolls under it.
  const frozenAfter = await cellRect(page, 0, 1);
  expect(frozenAfter).toEqual(frozen);
  const selection = await selectionState(page);
  expect(selection.anchor).toEqual({ row: 39, col: 0 });
  expect(selection.focus.row).toBeLessThan(39);
  await page.mouse.up();
});

test('a Format Painter drag past the bottom edge paints the revealed rows', async ({ page }) => {
  await openGrid(page);
  await selectCell(page, 'A1');
  await page.keyboard.press('Control+b');
  await page.getByRole('button', { name: 'Format Painter' }).click();
  await expect(page.getByTestId('status-message')).toContainText('Format Painter picked up');

  const { box, x } = await pressCell(page, 2, 0);
  await page.mouse.move(x, box.y + box.height - 2, { steps: 4 });
  await page.waitForTimeout(400);
  expect((await scrollOffset(page)).y).toBeGreaterThan(0);
  await page.mouse.up();

  await expect(page.getByTestId('status-message')).toContainText('Applied formatting');
  // The target begins at A3 and reaches the rows edge scrolling revealed.
  await expect
    .poll(async () => (await page.evaluate(() => window.__calcplus.getCellStyle(20, 0)))?.font.bold)
    .toBe(true);

  // The whole drag was one transaction; one undo removes it.
  await page.keyboard.press('Control+z');
  await expect
    .poll(async () => (await page.evaluate(() => window.__calcplus.getCellStyle(20, 0)))?.font.bold)
    .not.toBe(true);
});

test('a fill dragged past the bottom edge is one undo step', async ({ page }) => {
  await openGrid(page);
  await selectCell(page, 'B2');
  const versionBefore = await page.evaluate(() => window.__calcplus.getState().version);
  const handle = await fillHandlePoint(page, 1, 1);
  const box = await canvasBox(page);

  await page.mouse.move(handle.x, handle.y);
  await page.mouse.down();
  await page.mouse.move(handle.x, box.y + box.height - 2, { steps: 6 });
  await page.waitForTimeout(500);

  // The preview alone changes no document state.
  expect(await page.evaluate(() => window.__calcplus.getState().version)).toBe(versionBefore);
  expect((await scrollOffset(page)).y).toBeGreaterThan(0);

  await page.mouse.up();
  const versionAfter = await page.evaluate(() => window.__calcplus.getState().version);
  expect(versionAfter - versionBefore).toBe(1);

  await expect.poll(async () => (await selectionRangeEnd(page))?.endRow ?? 0).toBeGreaterThan(1);
  const destination = await selectionRangeEnd(page);
  if (destination === null) throw new Error('The fill left no range selection');
  expect(destination.endRow).toBeGreaterThan(1);
  expect(await cellDisplay(page, destination.endRow, destination.startCol)).not.toBeNull();

  await page.keyboard.press('Control+z');
  await expect.poll(() => cellDisplay(page, destination.endRow, destination.startCol)).toBeNull();
});
