import { Buffer } from 'node:buffer';
import { expect, test, type Page } from '@playwright/test';
import { buildXlsx } from '../helpers/build-xlsx';
import { cellRect, clickCell, regionSignature, type CellRect } from './grid-helpers';

/**
 * The frozen strips are strict clip bands (ticket 41).
 *
 * With frozen rows, scrolling up used to paint the incoming row's number over
 * the frozen strip's number, leaving two numbers stacked in the gutter. When a
 * scrolling row straddles the divider, its header label was painted in full,
 * centred half above the divider; the frozen columns clip their labels at the
 * divider, the frozen rows did not.
 *
 * Everything here measures painted pixels: the strip is a paint result, and the
 * scroll direction is what used to decide whether it was correct. The frozen
 * strip must look exactly the same at every scroll offset.
 */
function scrollXlsx(frozenRows: number, frozenCols: number): Uint8Array {
  const rows: string[] = [];
  for (let row = 1; row <= 60; row += 1) {
    const cells: string[] = [];
    for (let col = 0; col < 16; col += 1) {
      const ref = `${String.fromCharCode(65 + col)}${row}`;
      // Every row and column paints distinct ink, so a label or a value from
      // another row changes the signature of the strip.
      cells.push(`<c r="${ref}"><v>${row * 100 + col}</v></c>`);
    }
    rows.push(`<row r="${row}">${cells.join('')}</row>`);
  }
  const pane =
    frozenRows > 0 || frozenCols > 0
      ? `<pane xSplit="${frozenCols}" ySplit="${frozenRows}" topLeftCell="${String.fromCharCode(65 + frozenCols)}${frozenRows + 1}" activePane="bottomRight" state="frozen"/>`
      : '';
  return buildXlsx(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews><sheetView tabSelected="1" workbookViewId="0">${pane}</sheetView></sheetViews>
  <sheetData>${rows.join('')}</sheetData>
</worksheet>`);
}

/** A pixel region of the canvas, in canvas CSS pixels. */
type Region = CellRect;

async function openGrid(
  page: Page,
  freeze: { readonly rows: number; readonly cols: number },
): Promise<void> {
  await page.goto('/?fsa=0');
  await page.addStyleTag({
    content: '.grid-area { flex: 0 0 auto; height: 640px; padding: 12px; }',
  });
  await page.getByTestId('file-input').setInputFiles({
    name: 'freeze-scroll.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: Buffer.from(scrollXlsx(freeze.rows, freeze.cols)),
  });
  await page.waitForFunction(() => window.__calcplus.getCellDisplay(0, 0) === '100');
  await page.evaluate(() => document.fonts.ready.then(() => true));
  // Park the active cell where it is on screen at rest but never lands inside
  // a frozen strip at the offsets under test: a cell box that scrolls under a
  // strip paints inside it, which is a separate defect (ticket 41 follow-up).
  await clickCell(page, 15, 12);
  await settle(page);
}

/** Wait for the scheduled canvas paint, not just the projection update. */
async function settle(page: Page): Promise<void> {
  await page.evaluate(
    () =>
      new Promise<void>((resolve) => {
        requestAnimationFrame(() => requestAnimationFrame(() => resolve()));
      }),
  );
}

async function wheel(page: Page, deltaX: number, deltaY: number): Promise<void> {
  const box = await page.getByTestId('grid-canvas').boundingBox();
  if (box === null) throw new Error('The grid canvas has no box');
  await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
  await page.mouse.wheel(deltaX, deltaY);
  await settle(page);
}

/** The painted signature of a region, for equality between two scroll offsets. */
async function signature(page: Page, region: Region): Promise<string> {
  return regionSignature(page, region.x, region.y, region.width, region.height);
}

/** The dark pixels in a region: label and cell text. */
async function ink(page: Page, region: Region): Promise<number> {
  return page.evaluate((area: Region) => {
    const canvas = document.querySelector<HTMLCanvasElement>('[data-testid="grid-canvas"]');
    const context = canvas?.getContext('2d') ?? null;
    if (context === null) throw new Error('No canvas context');
    const ratio = canvas === null ? 1 : canvas.width / canvas.getBoundingClientRect().width;
    const data = context.getImageData(
      Math.round(area.x * ratio),
      Math.round(area.y * ratio),
      Math.max(1, Math.round(area.width * ratio)),
      Math.max(1, Math.round(area.height * ratio)),
    ).data;
    let dark = 0;
    for (let index = 0; index < data.length; index += 4) {
      if ((data[index] ?? 255) < 120) dark += 1;
    }
    return dark;
  }, region);
}

/**
 * The frozen row strip: the gutter numbers and the cell contents of the frozen
 * rows, from the top of the body to the freeze divider.
 */
async function frozenRowBand(page: Page, frozenRows: number): Promise<Region> {
  const box = await page.getByTestId('grid-canvas').boundingBox();
  if (box === null) throw new Error('The grid canvas has no box');
  const first = await cellRect(page, 0, 0);
  const last = await cellRect(page, frozenRows - 1, 0);
  return {
    x: 0,
    y: first.y,
    // Clear of the scrollbar strip at the right edge.
    width: box.width - 16,
    height: last.y + last.height - first.y,
  };
}

/** The frozen column strip, from the top of the body down to the bottom. */
async function frozenColumnBand(page: Page, frozenCols: number): Promise<Region> {
  const box = await page.getByTestId('grid-canvas').boundingBox();
  if (box === null) throw new Error('The grid canvas has no box');
  const first = await cellRect(page, 0, 0);
  const last = await cellRect(page, 0, frozenCols - 1);
  return {
    x: first.x,
    y: first.y,
    width: last.x + last.width - first.x,
    // Clear of the scrollbar strip at the bottom edge.
    height: box.height - first.y - 16,
  };
}

/** The row height of the fixture's uniform rows, in canvas pixels. */
async function rowHeight(page: Page): Promise<number> {
  return (await cellRect(page, 1, 0)).height;
}

/** The column width of the fixture's uniform columns, in canvas pixels. */
async function columnWidth(page: Page): Promise<number> {
  return (await cellRect(page, 0, 1)).width;
}

test('the frozen row strip stays clean at every vertical scroll offset', async ({ page }) => {
  // Two frozen rows and one frozen column: the four-pane case, as in the
  // maintainer's capture.
  await openGrid(page, { rows: 2, cols: 1 });
  const band = await frozenRowBand(page, 2);
  const atRest = await signature(page, band);

  // Down: the strip never moves, so its pixels never change.
  await wheel(page, 0, 400);
  expect(await signature(page, band), 'scrolling down changed the frozen strip').toBe(atRest);

  // Up: the scroll lands between two row boundaries, so the row above the
  // divider straddles it. Its label used to paint half above the divider, over
  // the frozen row's own number. The fixture's rows are uniform, so the sheet
  // offset divided by the row height gives the top scrolling row.
  await wheel(page, 0, -190);
  const scrollY = (await page.evaluate(() => window.__calcplus.getScroll())).y;
  const topRow = Math.floor(scrollY / (await rowHeight(page)));
  const straddling = await cellRect(page, topRow, 0);
  const divider = band.y + band.height;
  // The case is real: the top scrolling row starts above the divider and ends
  // below it.
  expect(straddling.y).toBeLessThan(divider);
  expect(straddling.y + straddling.height).toBeGreaterThan(divider);

  // The clip cuts the straddling label at the divider; it does not silence the
  // scrolling labels. The first fully visible row keeps its number.
  const firstFull = await cellRect(page, topRow + 1, 0);
  expect(
    await ink(page, { x: 2, y: firstFull.y + 2, width: 44, height: firstFull.height - 4 }),
    'the scrolling row numbers vanished below the divider',
  ).toBeGreaterThan(5);

  expect(await signature(page, band), 'scrolling up painted into the frozen strip').toBe(atRest);
});

test('the frozen column strip stays clean at every horizontal scroll offset', async ({ page }) => {
  await openGrid(page, { rows: 1, cols: 2 });
  const band = await frozenColumnBand(page, 2);
  const atRest = await signature(page, band);

  // Right: the strip never moves, so its pixels never change.
  await wheel(page, 400, 0);
  expect(await signature(page, band), 'scrolling right changed the frozen strip').toBe(atRest);

  // Left: the first scrolling column straddles the divider. Its label must
  // clip at the divider, the way the row labels now do.
  await wheel(page, -190, 0);
  const scrollX = (await page.evaluate(() => window.__calcplus.getScroll())).x;
  const straddling = await cellRect(page, 0, Math.floor(scrollX / (await columnWidth(page))));
  const divider = band.x + band.width;
  expect(straddling.x).toBeLessThan(divider);
  expect(straddling.x + straddling.width).toBeGreaterThan(divider);

  // The clip cuts the straddling label at the divider; it does not silence the
  // scrolling labels. The first fully visible column keeps its own label.
  const firstFull = await cellRect(page, 0, Math.floor(scrollX / (await columnWidth(page))) + 1);
  expect(
    await ink(page, { x: firstFull.x + 2, y: 2, width: firstFull.width - 4, height: 20 }),
    'the scrolling column labels vanished past the divider',
  ).toBeGreaterThan(5);

  expect(await signature(page, band), 'scrolling left painted into the frozen strip').toBe(atRest);
});
