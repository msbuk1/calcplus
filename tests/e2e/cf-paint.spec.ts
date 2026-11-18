import { Buffer } from 'node:buffer';
import { expect, test, type Page } from '@playwright/test';
import { buildXlsx } from '../helpers/build-xlsx';
import { clickCell, pixelAt, type CellRect } from './grid-helpers';

/**
 * Conditional formatting paints (CF12 `cf-merged`, CF13 `cf-panes-zoom`).
 *
 * The fixture freezes column A, so the freeze divider sits at the right edge of
 * column A. Column A carries a colour-scale band, column B carries data bars
 * with mixed signs (axis at zero), column C carries icon-only icon sets, and the
 * merged A1:D1 row carries a rule that crosses the divider.
 *
 * Everything below measures painted pixels: the canvas is the only place a
 * render-only feature can be checked (§2: CF never changes the cell).
 */
function cfPaintXlsx(): Uint8Array {
  const worksheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0">
      <pane xSplit="1" ySplit="0" topLeftCell="B1" activePane="topRight" state="frozen"/>
    </sheetView>
  </sheetViews>
  <sheetData>
    <row r="1"><c r="A1"><v>0</v></c><c r="B1"/><c r="C1"/><c r="D1"/></row>
    <row r="2"><c r="A2"><v>100</v></c></row>
    <row r="4"><c r="A4"><v>0</v></c><c r="B4"><v>-10</v></c><c r="C4"><v>0</v></c></row>
    <row r="5"><c r="A5"><v>25</v></c><c r="B5"><v>-5</v></c><c r="C5"><v>50</v></c></row>
    <row r="6"><c r="A6"><v>50</v></c><c r="B6"><v>0</v></c><c r="C6"><v>100</v></c></row>
    <row r="7"><c r="A7"><v>75</v></c><c r="B7"><v>5</v></c></row>
    <row r="8"><c r="A8"><v>100</v></c><c r="B8"><v>10</v></c></row>
    <row r="10"><c r="A10"><v>50</v></c><c r="B10"/></row>
  </sheetData>
  <mergeCells count="1"><mergeCell ref="A1:D1"/></mergeCells>
  <conditionalFormatting sqref="A1:A2">
    <cfRule type="colorScale" priority="1">
      <colorScale>
        <cfvo type="min"/><cfvo type="max"/>
        <color rgb="FFFF0000"/><color rgb="FFFFFFFF"/>
      </colorScale>
    </cfRule>
  </conditionalFormatting>
  <conditionalFormatting sqref="A4:A8">
    <cfRule type="colorScale" priority="1">
      <colorScale>
        <cfvo type="min"/><cfvo type="max"/>
        <color rgb="FFFF0000"/><color rgb="FFFFFFFF"/>
      </colorScale>
    </cfRule>
  </conditionalFormatting>
  <conditionalFormatting sqref="B4:B8">
    <cfRule type="dataBar" priority="1">
      <dataBar showValue="0">
        <cfvo type="min"/><cfvo type="max"/>
        <color rgb="FF638EC6"/>
      </dataBar>
    </cfRule>
  </conditionalFormatting>
  <conditionalFormatting sqref="A10">
    <cfRule type="dataBar" priority="1">
      <dataBar showValue="0">
        <cfvo type="num" val="0"/><cfvo type="num" val="100"/>
        <color rgb="FF638EC6"/>
      </dataBar>
    </cfRule>
  </conditionalFormatting>
  <conditionalFormatting sqref="C4:C6">
    <cfRule type="iconSet" priority="1">
      <iconSet iconSet="3Arrows" showValue="0">
        <cfvo type="percent" val="33"/><cfvo type="percent" val="67"/>
      </iconSet>
    </cfRule>
  </conditionalFormatting>
</worksheet>`;
  return buildXlsx(worksheet);
}

/** The debug surface this spec uses beyond the shared `DebugApi` typing. */
interface CfDebugApi {
  setZoom(percent: number): void;
  getViewState(): { zoom: number };
}

async function openFixture(page: Page): Promise<void> {
  await page.goto('/?fsa=0');
  await page.addStyleTag({
    content: '.grid-area { flex: 0 0 auto; height: 640px; padding: 12px; }',
  });
  await page.getByTestId('file-input').setInputFiles({
    name: 'm5-cf-paint.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: Buffer.from(cfPaintXlsx()),
  });
  await page.waitForFunction(() => window.__calcplus.getCellDisplay(0, 0) === '0');
  await page.evaluate(() => document.fonts.ready.then(() => true));
  // Park the selection away from the measured cells: the selection overlay is
  // painted after every cell, and it would tint the pixels under test.
  await clickCell(page, 3, 3);
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

async function rect(page: Page, row: number, col: number): Promise<CellRect> {
  const found = await page.evaluate(
    ([targetRow, targetCol]) =>
      window.__calcplus.getCellRect(targetRow as number, targetCol as number),
    [row, col] as const,
  );
  if (found === null) throw new Error(`Missing rectangle for ${row}:${col}`);
  return found;
}

interface PixelCounts {
  readonly red: number;
  readonly blue: number;
  readonly white: number;
  readonly dark: number;
  readonly total: number;
}

/** Count painted pixel classes inside a rectangle of the canvas. */
async function scan(page: Page, region: CellRect): Promise<PixelCounts> {
  return page.evaluate((area: CellRect) => {
    const canvas = document.querySelector<HTMLCanvasElement>('[data-testid="grid-canvas"]');
    const context = canvas?.getContext('2d') ?? null;
    if (context === null) throw new Error('No canvas context');
    const ratio = canvas === null ? 1 : canvas.width / canvas.getBoundingClientRect().width;
    const left = Math.round(area.x * ratio);
    const top = Math.round(area.y * ratio);
    const width = Math.max(1, Math.round(area.width * ratio));
    const height = Math.max(1, Math.round(area.height * ratio));
    const data = context.getImageData(left, top, width, height).data;
    let red = 0;
    let blue = 0;
    let white = 0;
    let dark = 0;
    for (let index = 0; index < data.length; index += 4) {
      const r = data[index] ?? 0;
      const g = data[index + 1] ?? 0;
      const b = data[index + 2] ?? 0;
      if (r > 180 && g < 140 && b < 140) red += 1;
      if (b > 150 && b > r + 30) blue += 1;
      if (r > 240 && g > 240 && b > 240) white += 1;
      if (r < 60 && g < 60 && b < 60) dark += 1;
    }
    return { red, blue, white, dark, total: width * height };
  }, region);
}

/** The interior of a cell: one pixel inside each edge, away from gridlines. */
function inner(area: CellRect): CellRect {
  return {
    x: area.x + 2,
    y: area.y + 2,
    width: Math.max(1, area.width - 4),
    height: Math.max(1, area.height - 4),
  };
}

test.describe('conditional formatting paints the grid', () => {
  test.use({
    viewport: { width: 1280, height: 920 },
    deviceScaleFactor: 1,
    colorScheme: 'light',
    reducedMotion: 'reduce',
  });

  test('paints a colour-scale band and a data bar', async ({ page }) => {
    await openFixture(page);

    const low = await rect(page, 3, 0);
    const high = await rect(page, 7, 0);
    const negative = await rect(page, 3, 1);
    const positive = await rect(page, 7, 1);

    // The colour scale runs red at the range minimum to white at its maximum.
    const lowFill = await pixelAt(page, low.x + 8, low.y + low.height / 2);
    expect(lowFill.r, JSON.stringify(lowFill)).toBeGreaterThan(200);
    expect(lowFill.g).toBeLessThan(70);
    expect(lowFill.b).toBeLessThan(70);
    const highFill = await pixelAt(page, high.x + 8, high.y + high.height / 2);
    expect(highFill.r).toBeGreaterThan(245);
    expect(highFill.g).toBeGreaterThan(245);
    expect(highFill.b).toBeGreaterThan(245);

    // A data bar paints behind the (hidden) value with its axis at zero: a
    // negative bar grows left from the axis, a positive one grows right.
    const negativeInside = await pixelAt(
      page,
      negative.x + negative.width * 0.25,
      negative.y + negative.height / 2,
    );
    expect(negativeInside.b, JSON.stringify(negativeInside)).toBeGreaterThan(150);
    expect(negativeInside.b).toBeGreaterThan(negativeInside.r + 30);
    const negativePast = await pixelAt(
      page,
      negative.x + negative.width * 0.8,
      negative.y + negative.height / 2,
    );
    expect(negativePast.r).toBeGreaterThan(245);
    expect(negativePast.g).toBeGreaterThan(245);
    expect(negativePast.b).toBeGreaterThan(245);

    const positiveInside = await pixelAt(
      page,
      positive.x + positive.width * 0.6,
      positive.y + positive.height / 2,
    );
    expect(positiveInside.b, JSON.stringify(positiveInside)).toBeGreaterThan(150);
    expect(positiveInside.b).toBeGreaterThan(positiveInside.r + 30);
    const positivePast = await pixelAt(
      page,
      positive.x + positive.width * 0.25,
      positive.y + positive.height / 2,
    );
    expect(positivePast.r).toBeGreaterThan(245);
    expect(positivePast.g).toBeGreaterThan(245);
    expect(positivePast.b).toBeGreaterThan(245);

    // The zero axis is a grey line at the middle of a mixed-sign range.
    const axis = await pixelAt(
      page,
      positive.x + positive.width / 2,
      positive.y + positive.height / 2,
    );
    expect(Math.abs(axis.r - axis.g)).toBeLessThan(14);
    expect(Math.abs(axis.g - axis.b)).toBeLessThan(14);
  });

  test('paints a merged rule across the merged rectangle', async ({ page }) => {
    await openFixture(page);
    const anchor = await rect(page, 0, 0);
    const last = await rect(page, 0, 3);
    const divider = anchor.x + anchor.width;

    // The merge really does cross the freeze divider.
    expect(divider).toBeLessThan(last.x + last.width);

    // The CF fill paints from the anchor's value across the whole merged
    // rectangle: inside the frozen pane and inside the scrolling pane (§5.2,
    // §5.3). A fill cut at either pane edge fails one of these.
    const frozenPart = await pixelAt(page, anchor.x + 8, anchor.y + anchor.height / 2);
    expect(frozenPart.r, JSON.stringify(frozenPart)).toBeGreaterThan(200);
    expect(frozenPart.g).toBeLessThan(70);
    expect(frozenPart.b).toBeLessThan(70);

    // Three quarters across the merge: past the divider, clear of the
    // right-aligned value's ink.
    const scrollingPart = await pixelAt(
      page,
      anchor.x + (last.x + last.width - anchor.x) * 0.75,
      last.y + last.height / 2,
    );
    expect(scrollingPart.r, JSON.stringify(scrollingPart)).toBeGreaterThan(200);
    expect(scrollingPart.g).toBeLessThan(70);
    expect(scrollingPart.b).toBeLessThan(70);

    // The rule covers the merged anchor only: the row below is manual white.
    const below = await rect(page, 1, 3);
    const belowPixel = await pixelAt(page, below.x + 8, below.y + below.height / 2);
    expect(belowPixel.r).toBeGreaterThan(245);
    expect(belowPixel.g).toBeGreaterThan(245);
    expect(belowPixel.b).toBeGreaterThan(245);
  });

  test('does not bleed across a freeze divider', async ({ page }) => {
    await openFixture(page);
    const frozen = await rect(page, 9, 0);
    const scrolling = await rect(page, 9, 1);
    // Column A is the frozen pane; column B starts on the scrolling side.
    expect(scrolling.x).toBeGreaterThanOrEqual(frozen.x + frozen.width);

    // The bar fills half of the frozen cell ...
    const bar = await pixelAt(page, frozen.x + frozen.width * 0.25, frozen.y + frozen.height / 2);
    expect(bar.b, JSON.stringify(bar)).toBeGreaterThan(150);
    expect(bar.b).toBeGreaterThan(bar.r + 30);

    // ... and the neighbouring scrolling pane keeps its manual white.
    const neighbour = await scan(page, inner(scrolling));
    expect(neighbour.blue, JSON.stringify(neighbour)).toBe(0);
    expect(neighbour.red).toBe(0);
    expect(neighbour.white).toBeGreaterThan(neighbour.total * 0.9);
  });

  test('suppresses icons below 25% zoom while bars still paint', async ({ page }) => {
    await openFixture(page);

    // C4 holds the lowest value, so its icon is the first (red) glyph. The
    // icon-only rule hides the value, so red ink can only be the icon.
    const iconAtFull = await scan(page, inner(await rect(page, 3, 2)));
    expect(iconAtFull.red, JSON.stringify(iconAtFull)).toBeGreaterThan(2);

    await page.evaluate(() => (window.__calcplus as unknown as CfDebugApi).setZoom(20));
    await page.waitForFunction(
      () => (window.__calcplus as unknown as CfDebugApi).getViewState().zoom === 20,
    );
    await settle(page);
    // The scroll offset is kept in pixels across a zoom change, so the viewport
    // has slid right; bring the columns back into view before measuring.
    const box = await page.getByTestId('grid-canvas').boundingBox();
    if (box === null) throw new Error('The grid canvas has no box');
    await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
    await page.mouse.wheel(-400, 0);
    await settle(page);

    const zoomed = await rect(page, 3, 2);
    expect(zoomed.width).toBeLessThan(20);
    const iconAtLow = await scan(page, inner(zoomed));
    expect(iconAtLow.red, JSON.stringify(iconAtLow)).toBe(0);

    // Fills and bars still paint at 20% (§5.4).
    const band = await scan(page, inner(await rect(page, 3, 0)));
    expect(band.red, JSON.stringify(band)).toBeGreaterThan(0);
    const barAtLow = await scan(page, inner(await rect(page, 7, 1)));
    expect(barAtLow.blue, JSON.stringify(barAtLow)).toBeGreaterThan(0);
  });
});
