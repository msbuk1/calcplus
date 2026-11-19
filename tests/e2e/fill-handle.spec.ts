import { Buffer } from 'node:buffer';
import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';
import { buildXlsx } from '../helpers/build-xlsx';

/**
 * The fill handle on the real grid (feature matrix §1 "Fill handle: copy,
 * relative-reference adjustment, double-click to adjacent data"; command-system
 * §4: one gesture is one transaction).
 *
 * The workbook is built at test time, exactly like the data-tools fixtures, so
 * the walk never depends on a binary that could drift from the model:
 *
 *   A  Qty      1, 2, 3, 4, 5      (the adjacent data the double-click follows)
 *   B  Double   =A2*2 with a fill and a bold font
 *   C  Tenfold  =A2*10
 *   F  Base     10
 *   G  Grow     =F2*2
 *   A9:B9       a merged range the blocked-fill test drags onto
 */

function columnLetters(index: number): string {
  let letters = '';
  let value = index + 1;
  while (value > 0) {
    const remainder = (value - 1) % 26;
    letters = String.fromCharCode(65 + remainder) + letters;
    value = Math.floor((value - 1) / 26);
  }
  return letters;
}

function textCell(ref: string, text: string, style?: number): string {
  const styleAttribute = style === undefined ? '' : ` s="${style}"`;
  return `<c r="${ref}"${styleAttribute} t="inlineStr"><is><t>${text}</t></is></c>`;
}

function numberCell(ref: string, value: number, style?: number): string {
  const styleAttribute = style === undefined ? '' : ` s="${style}"`;
  return `<c r="${ref}"${styleAttribute}><v>${value}</v></c>`;
}

function formulaCell(ref: string, formula: string, cached: number, style?: number): string {
  const styleAttribute = style === undefined ? '' : ` s="${style}"`;
  return `<c r="${ref}"${styleAttribute}><f>${formula}</f><v>${cached}</v></c>`;
}

export function fillWorkbookXlsx(): Uint8Array {
  const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><name val="Calibri"/></font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFC000"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="1" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;

  const rows = [
    `<row r="1">${textCell('A1', 'Qty')}${textCell('B1', 'Double')}${textCell('C1', 'Tenfold')}${textCell('F1', 'Base')}${textCell('G1', 'Grow')}</row>`,
    `<row r="2">${numberCell('A2', 1)}${formulaCell('B2', 'A2*2', 2, 1)}${formulaCell('C2', 'A2*10', 10)}${numberCell('F2', 10)}${formulaCell('G2', 'F2*2', 20)}</row>`,
    `<row r="3">${numberCell('A3', 2)}</row>`,
    `<row r="4">${numberCell('A4', 3)}</row>`,
    `<row r="5">${numberCell('A5', 4)}</row>`,
    `<row r="6">${numberCell('A6', 5)}</row>`,
    `<row r="8">${numberCell('A8', 1)}</row>`,
    `<row r="9">${textCell('A9', 'Merged')}</row>`,
  ].join('');

  const worksheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:${columnLetters(7)}9"/>
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetData>${rows}</sheetData>
  <mergeCells count="1"><mergeCell ref="A9:B9"/></mergeCells>
</worksheet>`;
  return buildXlsx(worksheetXml, { stylesXml });
}

async function openFillSheet(page: Page): Promise<void> {
  await page.goto('/?fsa=0');
  await page.getByTestId('file-input').setInputFiles({
    name: 'fill.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: Buffer.from(fillWorkbookXlsx()),
  });
  await expect(page.getByTestId('file-name')).toHaveText('fill.xlsx');
  await expect(page.getByTestId('grid-canvas')).toBeVisible();
}

async function selectRange(page: Page, address: string): Promise<void> {
  const nameBox = page.getByTestId('name-box');
  await nameBox.click();
  await page.keyboard.press('Control+a');
  await page.keyboard.type(address);
  await page.keyboard.press('Enter');
  await expect(nameBox).toHaveValue(address);
}

async function cellRect(page: Page, row: number, col: number) {
  const rect = await page.evaluate(
    ([targetRow, targetCol]) =>
      window.__calcplus.getCellRect(targetRow as number, targetCol as number),
    [row, col] as const,
  );
  if (rect === null) throw new Error(`No rectangle for row ${row}, column ${col}`);
  return rect;
}

async function cellDisplay(page: Page, row: number, col: number): Promise<string | null> {
  return page.evaluate(
    ([targetRow, targetCol]) =>
      window.__calcplus.getCellDisplay(targetRow as number, targetCol as number),
    [row, col] as const,
  );
}

/** Page coordinates of the painted fill handle for the given active cell. */
async function fillHandlePoint(page: Page, row: number, col: number) {
  const isHandleOfActiveCell = async (): Promise<boolean> => {
    const handle = await page.evaluate(() => window.__calcplus.getFillHandle());
    if (handle === null) return false;
    const rect = await page.evaluate(
      ([targetRow, targetCol]) =>
        window.__calcplus.getCellRect(targetRow as number, targetCol as number),
      [row, col] as const,
    );
    if (rect === null) return false;
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
  const box = await page.getByTestId('grid-canvas').boundingBox();
  if (box === null) throw new Error('The grid canvas has no box');
  return {
    x: box.x + handle.x + handle.width / 2,
    y: box.y + handle.y + handle.height / 2,
  };
}

/** Drag the fill handle from the active cell to the given cell. */
async function dragFillHandle(
  page: Page,
  fromRow: number,
  fromCol: number,
  row: number,
  col: number,
) {
  const handle = await fillHandlePoint(page, fromRow, fromCol);
  const target = await cellRect(page, row, col);
  const box = await page.getByTestId('grid-canvas').boundingBox();
  if (box === null) throw new Error('The grid canvas has no box');
  await page.mouse.move(handle.x, handle.y);
  await page.mouse.down();
  await page.mouse.move(box.x + target.x + target.width / 2, box.y + target.y + target.height / 2, {
    steps: 8,
  });
  await page.mouse.up();
}

async function doubleClickFillHandle(page: Page, row: number, col: number): Promise<void> {
  const handle = await fillHandlePoint(page, row, col);
  await page.mouse.dblclick(handle.x, handle.y);
}

test('dragging the fill handle copies the block, adjusts formulas and is one undo step', async ({
  page,
}) => {
  await openFillSheet(page);
  await selectRange(page, 'B2');

  await dragFillHandle(page, 1, 1, 3, 1);

  // B3 and B4 are copies of B2 with the relative row shifted by one and two.
  await expect.poll(() => cellDisplay(page, 2, 1)).toBe('4');
  await expect.poll(() => cellDisplay(page, 3, 1)).toBe('6');
  expect(await cellDisplay(page, 4, 1)).toBeNull();
  // Formatting travels with the source cell.
  const style = await page.evaluate(() => window.__calcplus.getCellStyle(2, 1));
  expect(style?.fill.pattern).toBe('solid');
  expect(style?.font.bold).toBe(true);
  // The filled rectangle becomes the selection, as Excel leaves it.
  await expect(page.getByTestId('name-box')).toHaveValue('B2:B4');
  await expect(page.getByTestId('status-message')).toContainText('Filled B2:B4');

  // One gesture is one undo step: a single Ctrl+Z removes both filled rows.
  await page.keyboard.press('Control+z');
  await expect.poll(() => cellDisplay(page, 2, 1)).toBeNull();
  expect(await cellDisplay(page, 3, 1)).toBeNull();
  expect(await cellDisplay(page, 1, 1)).toBe('2');

  await page.keyboard.press('Control+y');
  await expect.poll(() => cellDisplay(page, 2, 1)).toBe('4');
  await expect.poll(() => cellDisplay(page, 3, 1)).toBe('6');
});

test('double-clicking the fill handle fills down to the adjacent data edge', async ({ page }) => {
  await openFillSheet(page);
  await selectRange(page, 'B2');

  await doubleClickFillHandle(page, 1, 1);

  // Column A runs to row 6, so the fill stops there: 2, 3, 4 and 5 doubled.
  await expect.poll(() => cellDisplay(page, 5, 1)).toBe('10');
  await expect.poll(() => cellDisplay(page, 2, 1)).toBe('4');
  expect(await cellDisplay(page, 3, 1)).toBe('6');
  expect(await cellDisplay(page, 4, 1)).toBe('8');
  expect(await cellDisplay(page, 6, 1)).toBeNull();
  await expect(page.getByTestId('name-box')).toHaveValue('B2:B6');

  // The handle consumed the gesture: no in-cell editor opened, so Ctrl+Z
  // reaches the undo command instead of being swallowed by the textarea.
  // The fill moved the painted handle away from the double-click point, so
  // this also covers the layout race the recorded press point closes.
  await expect(page.getByTestId('cell-editor')).toHaveCount(0);

  await page.keyboard.press('Control+z');
  await expect.poll(() => cellDisplay(page, 2, 1)).toBeNull();
  expect(await cellDisplay(page, 5, 1)).toBeNull();
});

test('a double-click away from the handle still opens the in-cell editor', async ({ page }) => {
  await openFillSheet(page);
  await selectRange(page, 'B2');

  const cell = await cellRect(page, 1, 1);
  const box = await page.getByTestId('grid-canvas').boundingBox();
  if (box === null) throw new Error('The grid canvas has no box');
  await page.mouse.dblclick(box.x + cell.x + cell.width / 2, box.y + cell.y + cell.height / 2);

  // The guard belongs to the handle alone; an ordinary cell double-click edits.
  await expect(page.getByTestId('cell-editor')).toBeVisible();
  await expect(page.getByTestId('cell-editor')).toHaveValue('A2*2');
  await page.keyboard.press('Escape');
  await expect(page.getByTestId('cell-editor')).toHaveCount(0);
});

test('Ctrl+D fills down and Ctrl+R fills right, one transaction each', async ({ page }) => {
  await openFillSheet(page);
  await selectRange(page, 'C2:C6');

  await page.keyboard.press('Control+d');
  await expect.poll(() => cellDisplay(page, 2, 2)).toBe('20');
  await expect.poll(() => cellDisplay(page, 5, 2)).toBe('50');
  await page.keyboard.press('Control+z');
  await expect.poll(() => cellDisplay(page, 2, 2)).toBeNull();
  expect(await cellDisplay(page, 5, 2)).toBeNull();

  await selectRange(page, 'G2:H2');
  await page.keyboard.press('Control+r');
  await expect.poll(() => cellDisplay(page, 1, 7)).toBe('40');

  // One undo removes the Ctrl+R fill.
  await page.keyboard.press('Control+z');
  await expect.poll(() => cellDisplay(page, 1, 7)).toBeNull();
});

test('the destination outline grows with the pointer and applies on release', async ({ page }) => {
  await openFillSheet(page);
  await selectRange(page, 'B2');
  const handle = await fillHandlePoint(page, 1, 1);
  const target = await cellRect(page, 3, 1);
  const box = await page.getByTestId('grid-canvas').boundingBox();
  if (box === null) throw new Error('The grid canvas has no box');

  await page.mouse.move(handle.x, handle.y);
  await page.mouse.down();
  await page.mouse.move(box.x + target.x + target.width / 2, box.y + target.y + target.height / 2, {
    steps: 5,
  });

  // The live outline covers the whole destination, source included.
  const preview = page.getByTestId('fill-preview');
  await expect(preview).toBeVisible();
  const start = await cellRect(page, 1, 1);
  const end = await cellRect(page, 3, 1);
  const previewBox = await preview.boundingBox();
  if (previewBox === null) throw new Error('The preview has no box');
  expect(previewBox.x).toBeCloseTo(box.x + start.x, 0);
  expect(previewBox.y).toBeCloseTo(box.y + start.y, 0);
  expect(previewBox.width).toBeCloseTo(end.x + end.width - start.x, 0);
  expect(previewBox.height).toBeCloseTo(end.y + end.height - start.y, 0);
  // Nothing is written until the release.
  expect(await cellDisplay(page, 2, 1)).toBeNull();

  await page.mouse.up();
  await expect.poll(() => cellDisplay(page, 2, 1)).toBe('4');
});

test("a fill onto a merged range is blocked with Excel's message and changes nothing", async ({
  page,
}) => {
  await openFillSheet(page);
  await selectRange(page, 'A8');

  await dragFillHandle(page, 7, 0, 8, 0);

  await expect(page.getByTestId('status-message')).toContainText('merged cell');
  expect(await cellDisplay(page, 8, 0)).toBe('Merged');
  expect(await page.evaluate(() => window.__calcplus.getState().canUndo)).toBe(false);
});
