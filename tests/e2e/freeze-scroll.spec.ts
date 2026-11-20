import { Buffer } from 'node:buffer';
import { createHash } from 'node:crypto';
import { expect, test, type Page } from '@playwright/test';
import { buildXlsx } from '../helpers/build-xlsx';
import {
  cellRect,
  clickCell,
  dragSelect,
  pixelAt,
  regionSignature,
  type CellRect,
} from './grid-helpers';

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

/**
 * A hash of the composited page pixels of a canvas region. Unlike the canvas
 * signature this includes DOM overlays (the marquee is DOM chrome), so it can
 * prove that an overlay does not paint inside a frozen strip.
 */
async function regionHash(page: Page, region: Region): Promise<string> {
  const box = await page.getByTestId('grid-canvas').boundingBox();
  if (box === null) throw new Error('The grid canvas has no box');
  const png = await page.screenshot({
    clip: {
      x: box.x + region.x,
      y: box.y + region.y,
      width: region.width,
      height: region.height,
    },
  });
  return createHash('sha256').update(png).digest('hex');
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

test('a selection scrolled under the frozen strip leaves no ink in it', async ({ page }) => {
  // The repro from ticket 45, with frozen rows only: the active cell and its
  // fill handle are painted with the selection overlays.
  await openGrid(page, { rows: 2, cols: 0 });
  const band = await frozenRowBand(page, 2);
  const divider = band.y + band.height;

  // Select a cell below the strip and park it there: the band carries no
  // selection ink to begin with, because the cell is outside it.
  await clickCell(page, 20, 12);
  await settle(page);
  const atRest = await signature(page, band);

  // Scroll the sheet down so the active cell slides under the strip.
  await wheel(page, 0, 400);
  const hidden = await cellRect(page, 20, 12);
  expect(hidden.y).toBeLessThan(divider);
  expect(hidden.y + hidden.height).toBeGreaterThan(band.y);

  expect(await signature(page, band), 'the selection painted inside the frozen strip').toBe(atRest);
  // Nothing is left to press either: the handle is not painted and not a target.
  expect(await page.evaluate(() => window.__calcplus.getFillHandle())).toBeNull();

  // Scroll back until the cell is in view again: the box and handle come back,
  // clipped to the pane rather than hidden for good.
  await wheel(page, 0, -190);
  const visible = await cellRect(page, 20, 12);
  expect(visible.y).toBeGreaterThanOrEqual(divider);
  const handle = await page.evaluate(() => window.__calcplus.getFillHandle());
  expect(handle).not.toBeNull();
  expect(handle?.x).toBeGreaterThanOrEqual(visible.x);
  expect(handle?.y).toBeGreaterThanOrEqual(visible.y);
  // The active-cell border is the accent colour, painted inside the cell edge.
  const border = await pixelAt(page, visible.x + 1, visible.y + visible.height / 2);
  expect(border.r, JSON.stringify(border)).toBe(26);
  expect(border.g).toBe(115);
  expect(border.b).toBe(232);
  expect(await signature(page, band)).toBe(atRest);
});

test('a range scrolled under the frozen strip leaves no tint or border in it', async ({ page }) => {
  await openGrid(page, { rows: 2, cols: 1 });
  const band = await frozenRowBand(page, 2);
  const divider = band.y + band.height;

  // A multi-cell selection whose active cell sits below the strip, as in the
  // selection-marquee case (ticket 31).
  await dragSelect(page, 20, 0, 24, 2);
  await settle(page);
  const atRest = await signature(page, band);

  // Stop the scroll 10px through a row, so the part of the range the pane shows
  // starts above the divider: without a clip its tint and top border would
  // paint over the frozen rows.
  await wheel(page, 0, 400);
  await wheel(page, 0, -10);
  const scrollY = (await page.evaluate(() => window.__calcplus.getScroll())).y;
  const top = await cellRect(page, 20, 0);
  expect(top.y).toBeLessThan(divider);
  const partial = await cellRect(page, Math.floor(scrollY / (await rowHeight(page))), 0);
  expect(partial.y).toBeLessThan(divider);

  expect(
    await signature(page, band),
    'the range tint or border painted inside the frozen strip',
  ).toBe(atRest);

  // The handle hangs on the range's own corner, below the strip.
  const handle = await page.evaluate(() => window.__calcplus.getFillHandle());
  expect(handle).not.toBeNull();
  expect(handle?.y).toBeGreaterThanOrEqual(divider);
});

test('a selection across the divider still tints both panes round the active cell', async ({
  page,
}) => {
  // Ticket 31 A stays as it was: the range tints in every pane it touches, and
  // the active cell keeps the grid background (the hole).
  await openGrid(page, { rows: 2, cols: 1 });
  await dragSelect(page, 0, 0, 4, 2);
  await settle(page);

  const inside = async (row: number, col: number) => {
    const rect = await cellRect(page, row, col);
    return pixelAt(page, rect.x + rect.width / 2, rect.y + rect.height / 2);
  };

  // B1 is the frozen row in a scrolling column; B5 scrolls under no strip.
  const frozenPart = await inside(0, 1);
  expect(frozenPart.r, JSON.stringify(frozenPart)).toBeGreaterThan(225);
  expect(frozenPart.r).toBeLessThan(240);
  expect(frozenPart.g).toBeGreaterThan(235);
  expect(frozenPart.b).toBeGreaterThan(248);

  const scrollingPart = await inside(4, 1);
  expect(scrollingPart.r, JSON.stringify(scrollingPart)).toBeGreaterThan(225);
  expect(scrollingPart.r).toBeLessThan(240);
  expect(scrollingPart.g).toBeGreaterThan(235);
  expect(scrollingPart.b).toBeGreaterThan(248);

  // The active cell A1 stays unfilled, as ticket 31 A rules.
  const active = await inside(0, 0);
  expect(active.r, JSON.stringify(active)).toBeGreaterThan(245);
  expect(active.g).toBeGreaterThan(245);
  expect(active.b).toBeGreaterThan(245);
});

test('the copy marquee leaves no ink in the frozen strip', async ({ page }) => {
  // The marquee is a DOM overlay, so its ink is measured from composited page
  // pixels, not from the canvas (ticket 46).
  await openGrid(page, { rows: 2, cols: 0 });
  const band = await frozenRowBand(page, 2);
  const divider = band.y + band.height;
  const canvas = await page.getByTestId('grid-canvas').boundingBox();
  if (canvas === null) throw new Error('The grid canvas has no box');

  // Copy a cell just below the strip: the marquee sits clear of the band.
  await clickCell(page, 21, 3);
  await page.keyboard.press('Control+c');
  const marquee = page.getByTestId('copy-marquee');
  await expect(marquee).toBeVisible();
  await settle(page);
  const cleanBand = await regionHash(page, band);

  // Scroll so the marquee's rectangle straddles the divider: its top is above
  // it and its bottom below.
  await wheel(page, 0, 400);
  await wheel(page, 0, -10);
  const straddling = await cellRect(page, 21, 3);
  expect(straddling.y).toBeLessThan(divider);
  expect(straddling.y + straddling.height).toBeGreaterThan(divider);
  expect(await regionHash(page, band), 'the marquee painted inside the frozen strip').toBe(
    cleanBand,
  );

  // The clip cuts the paint, not the geometry: the element is still the range
  // the paste targets, so its box matches the cells.
  const clipped = await marquee.boundingBox();
  expect(clipped).not.toBeNull();
  expect(clipped?.x).toBeCloseTo(canvas.x + straddling.x, 0);
  expect(clipped?.y).toBeCloseTo(canvas.y + straddling.y, 0);
  const clipPath = await marquee.evaluate((element) => getComputedStyle(element).clipPath);
  expect(clipPath, 'the marquee has no pane clip').not.toBe('none');

  // Scroll back into view: the marquee still tracks its cell and the band is
  // untouched.
  await wheel(page, 0, -190);
  const visible = await cellRect(page, 21, 3);
  expect(visible.y).toBeGreaterThanOrEqual(divider);
  const tracked = await marquee.boundingBox();
  expect(tracked?.x).toBeCloseTo(canvas.x + visible.x, 0);
  expect(tracked?.y).toBeCloseTo(canvas.y + visible.y, 0);
  await settle(page);
  expect(await regionHash(page, band)).toBe(cleanBand);
});

test('the Format Painter highlight leaves no ink in the frozen strip', async ({ page }) => {
  await openGrid(page, { rows: 2, cols: 0 });
  const band = await frozenRowBand(page, 2);
  const divider = band.y + band.height;
  const canvas = await page.getByTestId('grid-canvas').boundingBox();
  if (canvas === null) throw new Error('The grid canvas has no box');
  const cleanBand = await regionHash(page, band);

  // Pick up the painter, then hover a scrolling cell: the target highlight is
  // painted below the strip.
  await clickCell(page, 8, 3);
  await page.getByRole('button', { name: 'Format Painter' }).click();
  await expect(page.getByTestId('status-message')).toContainText('Format Painter picked up');
  const target = await cellRect(page, 21, 3);
  await page.mouse.move(
    canvas.x + target.x + target.width / 2,
    canvas.y + target.y + target.height / 2,
  );
  await settle(page);
  const tint = await pixelAt(page, target.x + target.width / 2, target.y + target.height / 2);
  expect(tint.r, JSON.stringify(tint)).toBeGreaterThan(245);
  expect(tint.r).toBeGreaterThan(tint.b + 20);

  // Scroll without moving the pointer: the target slides under the strip. Its
  // tint and dashed border must be cut at the divider.
  await page.mouse.wheel(0, 400);
  await page.mouse.wheel(0, -10);
  await settle(page);
  const hidden = await cellRect(page, 21, 3);
  expect(hidden.y).toBeLessThan(divider);
  expect(hidden.y + hidden.height).toBeGreaterThan(band.y);
  expect(
    await regionHash(page, band),
    'the painter highlight painted inside the frozen strip',
  ).toBe(cleanBand);

  // Scroll back: the same cell is highlighted again, and the strip is still
  // clean.
  await page.mouse.wheel(0, -190);
  await settle(page);
  const visible = await cellRect(page, 21, 3);
  expect(visible.y).toBeGreaterThanOrEqual(divider);
  const again = await pixelAt(page, visible.x + visible.width / 2, visible.y + visible.height / 2);
  expect(again.r, JSON.stringify(again)).toBeGreaterThan(245);
  expect(again.r).toBeGreaterThan(again.b + 20);
  expect(await regionHash(page, band)).toBe(cleanBand);
});

test('the fill destination outline leaves no ink in the frozen strip', async ({ page }) => {
  await openGrid(page, { rows: 2, cols: 0 });
  const band = await frozenRowBand(page, 2);
  const divider = band.y + band.height;
  const canvas = await page.getByTestId('grid-canvas').boundingBox();
  if (canvas === null) throw new Error('The grid canvas has no box');
  const cleanBand = await regionHash(page, band);

  // Grab the fill handle of B4 and drag to B7: the live destination covers
  // B4:B7.
  await clickCell(page, 3, 1);
  const handle = await page.evaluate(() => window.__calcplus.getFillHandle());
  expect(handle).not.toBeNull();
  const target = await cellRect(page, 6, 1);
  await page.mouse.move(canvas.x + (handle?.x ?? 0) + 3, canvas.y + (handle?.y ?? 0) + 3);
  await page.mouse.down();
  await page.mouse.move(
    canvas.x + target.x + target.width / 2,
    canvas.y + target.y + target.height / 2,
    { steps: 5 },
  );
  const preview = page.getByTestId('fill-preview');
  await expect(preview).toBeVisible();

  // Wheel-scroll mid-drag without moving the pointer: the destination slides
  // under the strip, and the outline follows the cells as they move.
  await page.mouse.wheel(0, 60);
  await settle(page);
  const top = await cellRect(page, 3, 1);
  const bottom = await cellRect(page, 6, 1);
  expect(top.y).toBeLessThan(divider);
  expect(bottom.y + bottom.height).toBeGreaterThan(divider);
  const box = await preview.boundingBox();
  expect(box).not.toBeNull();
  expect(box?.x).toBeCloseTo(canvas.x + top.x, 0);
  expect(box?.y).toBeCloseTo(canvas.y + top.y, 0);
  expect(box?.height).toBeCloseTo(bottom.y + bottom.height - top.y, 0);
  expect(
    await regionHash(page, band),
    'the fill destination outline painted inside the frozen strip',
  ).toBe(cleanBand);

  // The release still fills the destination B5:B7 from the source.
  await page.mouse.up();
  await expect(preview).toHaveCount(0);
  await expect.poll(() => page.evaluate(() => window.__calcplus.getCellDisplay(4, 1))).toBe('401');
  await expect.poll(() => page.evaluate(() => window.__calcplus.getCellDisplay(6, 1))).toBe('401');
});
