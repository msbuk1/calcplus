import { Buffer } from 'node:buffer';
import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';
import { buildXlsx } from '../helpers/build-xlsx';

/**
 * The drag gestures on the real grid (ticket 49; feature matrix §1
 * "Drag-and-drop move; Shift-drag insert", §7 file drop). The workbook is
 * built at test time like the other gesture specs, so the walk never depends
 * on a binary that could drift from the model:
 *
 *   A  1, 2, 3, 4        the block the moves carry
 *   B  =A1*10            a formula whose reference rides with its target
 *   C  30, 32            column data outside the bands
 */

function numberCell(ref: string, value: number): string {
  return `<c r="${ref}"><v>${value}</v></c>`;
}

function formulaCell(ref: string, formula: string, cached: number): string {
  return `<c r="${ref}"><f>${formula}</f><v>${cached}</v></c>`;
}

function dragWorkbookXlsx(): Uint8Array {
  const rows = [
    `<row r="1">${numberCell('A1', 1)}${formulaCell('B1', 'A1*10', 10)}${numberCell('C1', 30)}</row>`,
    `<row r="2">${numberCell('A2', 2)}${numberCell('C2', 32)}</row>`,
    `<row r="3">${numberCell('A3', 3)}</row>`,
    `<row r="4">${numberCell('A4', 4)}</row>`,
  ].join('');
  const worksheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:C4"/>
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetData>${rows}</sheetData>
</worksheet>`;
  return buildXlsx(worksheetXml);
}

async function openDragSheet(page: Page): Promise<void> {
  await page.goto('/?fsa=0');
  await page.getByTestId('file-input').setInputFiles({
    name: 'drag.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: Buffer.from(dragWorkbookXlsx()),
  });
  await expect(page.getByTestId('file-name')).toHaveText('drag.xlsx');
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

async function canvasPoint(
  page: Page,
  row: number,
  col: number,
): Promise<{ x: number; y: number }> {
  const box = await page.getByTestId('grid-canvas').boundingBox();
  if (box === null) throw new Error('The grid canvas has no box');
  const rect = await page.evaluate(
    ([targetRow, targetCol]) =>
      window.__calcplus.getCellRect(targetRow as number, targetCol as number),
    [row, col] as const,
  );
  if (rect === null) throw new Error(`No rectangle for row ${row}, column ${col}`);
  return { x: box.x + rect.x + rect.width / 2, y: box.y + rect.y + rect.height / 2 };
}

async function cellDisplay(page: Page, row: number, col: number): Promise<string | null> {
  return page.evaluate(
    ([targetRow, targetCol]) =>
      window.__calcplus.getCellDisplay(targetRow as number, targetCol as number),
    [row, col] as const,
  );
}

/** Drop one file on the app root, exactly as the browser would deliver it. */
async function dropFile(page: Page, name: string, bytes: Uint8Array): Promise<void> {
  await page.evaluate(
    ([fileName, base64]) => {
      const binary = atob(base64 as string);
      const buffer = new Uint8Array(binary.length);
      for (let index = 0; index < binary.length; index += 1) {
        buffer[index] = binary.charCodeAt(index);
      }
      const file = new File([buffer], fileName as string);
      const transfer = new DataTransfer();
      transfer.items.add(file);
      const root = document.querySelector('.app');
      if (root === null) throw new Error('The app root is missing');
      root.dispatchEvent(new DragEvent('dragover', { dataTransfer: transfer, bubbles: true }));
      root.dispatchEvent(new DragEvent('drop', { dataTransfer: transfer, bubbles: true }));
    },
    [name, Buffer.from(bytes).toString('base64')] as const,
  );
}

// ---------------------------------------------------------------------------
// Drag-and-drop move
// ---------------------------------------------------------------------------

test('a drag from inside the selection moves the range and is one undo step', async ({ page }) => {
  await openDragSheet(page);
  await selectRange(page, 'A1:B1');

  const from = await canvasPoint(page, 0, 0);
  const to = await canvasPoint(page, 3, 0);
  await page.mouse.move(from.x, from.y);
  await page.mouse.down();
  await page.mouse.move(to.x, to.y, { steps: 6 });
  // The destination outline follows the pointer (ticket 49).
  await expect(page.getByTestId('move-preview')).toBeVisible();
  await page.mouse.up();

  // The block lands at A4 and the formula's reference to the carried A1
  // follows it, so B4 still shows the carried value.
  await expect(page.getByTestId('name-box')).toHaveValue('A4:B4');
  expect(await cellDisplay(page, 0, 0)).toBeNull();
  expect(await cellDisplay(page, 3, 0)).toBe('1');
  expect(await cellDisplay(page, 3, 1)).toBe('10');

  // One undo removes the whole move: the gesture was one transaction.
  await page.keyboard.press('Control+z');
  expect(await cellDisplay(page, 0, 0)).toBe('1');
  expect(await cellDisplay(page, 0, 1)).toBe('10');
  expect(await cellDisplay(page, 3, 0)).toBe('4');
  expect(await page.evaluate(() => window.__calcplus.getState().canUndo)).toBe(false);
});

test('Esc cancels a drag-and-drop move', async ({ page }) => {
  await openDragSheet(page);
  await selectRange(page, 'A1:B1');

  const from = await canvasPoint(page, 0, 0);
  const to = await canvasPoint(page, 3, 0);
  await page.mouse.move(from.x, from.y);
  await page.mouse.down();
  await page.mouse.move(to.x, to.y, { steps: 6 });
  await expect(page.getByTestId('move-preview')).toBeVisible();
  await page.keyboard.press('Escape');
  await expect(page.getByTestId('move-preview')).toHaveCount(0);
  await page.mouse.up();

  expect(await cellDisplay(page, 0, 0)).toBe('1');
  expect(await cellDisplay(page, 3, 0)).toBe('4');
  // Nothing was recorded: the cancelled gesture is not in the undo history.
  expect(await page.evaluate(() => window.__calcplus.getState().canUndo)).toBe(false);
});

test('a drop back on the source changes nothing', async ({ page }) => {
  await openDragSheet(page);
  await selectRange(page, 'A1:B1');

  // Move the pointer far enough to start the gesture, but release over the
  // pressed cell: the destination is the source again (ticket 49).
  const from = await canvasPoint(page, 0, 0);
  await page.mouse.move(from.x, from.y);
  await page.mouse.down();
  await page.mouse.move(from.x + 4, from.y, { steps: 2 });
  await page.mouse.up();

  await expect(page.getByTestId('name-box')).toHaveValue('A1:B1');
  expect(await cellDisplay(page, 0, 0)).toBe('1');
  expect(await page.evaluate(() => window.__calcplus.getState().canUndo)).toBe(false);
});

// ---------------------------------------------------------------------------
// Shift-drag insert
// ---------------------------------------------------------------------------

test('a Shift-drag from the selection edge inserts and shifts cells down', async ({ page }) => {
  await openDragSheet(page);
  await selectRange(page, 'A1:A2');

  const from = await canvasPoint(page, 1, 0);
  const to = await canvasPoint(page, 3, 0);
  await page.keyboard.down('Shift');
  await page.mouse.move(from.x, from.y);
  await page.mouse.down();
  await page.mouse.move(to.x, to.y, { steps: 6 });
  await expect(page.getByTestId('insert-preview')).toBeVisible();
  await page.mouse.up();
  await page.keyboard.up('Shift');

  // The block lands at A3; the old A3 and A4 shift down inside column A only.
  expect(await cellDisplay(page, 2, 0)).toBe('1');
  expect(await cellDisplay(page, 3, 0)).toBe('2');
  expect(await cellDisplay(page, 4, 0)).toBe('3');
  expect(await cellDisplay(page, 5, 0)).toBe('4');
  // C2 is outside the band and keeps its place.
  expect(await cellDisplay(page, 1, 2)).toBe('32');
  // B1 kept its text but its target A1 carried to A3, so it still shows 10.
  expect(await cellDisplay(page, 0, 1)).toBe('10');

  // One undo removes the whole insert: the gesture was one transaction.
  await page.keyboard.press('Control+z');
  expect(await cellDisplay(page, 0, 0)).toBe('1');
  expect(await cellDisplay(page, 2, 0)).toBe('3');
  expect(await cellDisplay(page, 4, 0)).toBeNull();
  expect(await page.evaluate(() => window.__calcplus.getState().canUndo)).toBe(false);
});

test('a Shift-drag from the right edge inserts and shifts cells right', async ({ page }) => {
  await openDragSheet(page);
  await selectRange(page, 'A1:B1');

  const from = await canvasPoint(page, 0, 1);
  const to = await canvasPoint(page, 0, 5);
  await page.keyboard.down('Shift');
  await page.mouse.move(from.x, from.y);
  await page.mouse.down();
  await page.mouse.move(to.x, to.y, { steps: 6 });
  await expect(page.getByTestId('insert-preview')).toBeVisible();
  await page.mouse.up();
  await page.keyboard.up('Shift');

  // The block lands at E1; its formula carried along and still reads its own
  // carried A1, now at E1.
  expect(await cellDisplay(page, 0, 4)).toBe('1');
  expect(await cellDisplay(page, 0, 5)).toBe('10');
  // The old C1 is left of the gap and keeps its place.
  expect(await cellDisplay(page, 0, 2)).toBe('30');
});

// ---------------------------------------------------------------------------
// File drop
// ---------------------------------------------------------------------------

test('dropping an .xlsx opens it like File → Open', async ({ page }) => {
  await page.goto('/?fsa=0');
  await dropFile(page, 'dropped.xlsx', dragWorkbookXlsx());
  await expect(page.getByTestId('file-name')).toHaveText('dropped.xlsx');
  await expect(page.getByTestId('grid-canvas')).toBeVisible();
  expect(await cellDisplay(page, 0, 0)).toBe('1');
});

test('dropping a .csv opens it as a one-sheet workbook', async ({ page }) => {
  await page.goto('/?fsa=0');
  await dropFile(page, 'dropped.csv', new TextEncoder().encode('1,2\n3,hello'));
  await expect(page.getByTestId('file-name')).toHaveText('dropped.csv');
  await expect(page.getByTestId('grid-canvas')).toBeVisible();
  expect(await cellDisplay(page, 0, 0)).toBe('1');
  expect(await cellDisplay(page, 0, 1)).toBe('2');
  expect(await cellDisplay(page, 1, 0)).toBe('3');
  expect(await cellDisplay(page, 1, 1)).toBe('hello');
});

test('dropping an unsupported type shows an error and opens nothing', async ({ page }) => {
  await openDragSheet(page);

  await dropFile(page, 'notes.txt', new TextEncoder().encode('plain text'));
  await expect(page.getByTestId('status-message')).toContainText('Nothing was opened');
  await expect(page.getByTestId('file-name')).toHaveText('drag.xlsx');
  expect(await cellDisplay(page, 0, 0)).toBe('1');
  expect(await page.evaluate(() => window.__calcplus.getState().canUndo)).toBe(false);
});

test('a drop with unsaved changes asks before replacing the workbook', async ({ page }) => {
  await openDragSheet(page);
  await page.getByTestId('grid-shell').click();
  await page.keyboard.type('9');
  await page.keyboard.press('Enter');
  await expect
    .poll(async () => page.evaluate(() => window.__calcplus.getState().status))
    .toContain('Edited');

  const prompts: string[] = [];
  page.on('dialog', (dialog) => {
    prompts.push(dialog.message());
    void dialog.accept();
  });
  await dropFile(page, 'dropped.xlsx', dragWorkbookXlsx());
  await expect(page.getByTestId('file-name')).toHaveText('dropped.xlsx');
  expect(prompts.some((message) => message.includes('unsaved changes'))).toBe(true);
  // The typed edit belonged to the replaced workbook and is gone.
  expect(await cellDisplay(page, 2, 0)).toBe('3');
});
