import { Buffer } from 'node:buffer';
import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';
import { buildMultiSheetXlsx } from '../helpers/build-xlsx';
import { regionSignature, type CellRect } from './grid-helpers';

/**
 * Interaction-layer follow-ups for hyperlinks and comments (feature matrix §8:
 * "Hyperlinks: rendered and followable (Ctrl+click)" and "Legacy notes and
 * threaded comments: indicators and popups render").
 *
 * The workbook is built at test time so the walk never depends on a binary:
 *
 *   Links sheet
 *     A1 plain "Report" / B1 the same text with an external link and a tooltip
 *     A2 a javascript: link that must be refused
 *     A3 an internal link to 'Target Sheet'!B3
 *     A4 an internal link to the range 'Target Sheet'!B3:C4
 *     A5 an internal link to a sheet that does not exist
 *     C6 a legacy note, C8 a threaded comment, C9 its reply
 *   Target Sheet
 *     B3 and C4 hold the jump targets
 */

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const SPREADSHEET_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PACKAGE_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

function textCell(ref: string, text: string): string {
  return `<c r="${ref}" t="inlineStr"><is><t>${text}</t></is></c>`;
}

const linksSheetXml = `${XML_HEADER}
<worksheet xmlns="${SPREADSHEET_NS}" xmlns:r="${REL_NS}">
  <sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetData>
    <row r="1">${textCell('A1', 'Report')}${textCell('B1', 'Report')}</row>
    <row r="2">${textCell('A2', 'Blocked')}</row>
    <row r="3">${textCell('A3', 'Internal')}</row>
    <row r="4">${textCell('A4', 'Range')}</row>
    <row r="5">${textCell('A5', 'Broken')}</row>
    <row r="6">${textCell('A6', 'Far')}${textCell('C6', 'Legacy cell')}</row>
    <row r="8">${textCell('C8', 'Thread cell')}</row>
    <row r="9">${textCell('C9', 'Reply cell')}</row>
  </sheetData>
  <hyperlinks>
    <hyperlink ref="B1" r:id="rId1" tooltip="Open the report"/>
    <hyperlink ref="A2" r:id="rId2" tooltip="Blocked link"/>
    <hyperlink ref="A3" location="'Target Sheet'!B3"/>
    <hyperlink ref="A4" location="'Target Sheet'!B3:C4"/>
    <hyperlink ref="A5" location="Missing!A1"/>
    <hyperlink ref="A6" location="'Target Sheet'!A40"/>
  </hyperlinks>
</worksheet>`;

const targetSheetXml = `${XML_HEADER}
<worksheet xmlns="${SPREADSHEET_NS}">
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetData>
    <row r="3">${textCell('A3', 'target label')}${textCell('B3', 'target cell')}${textCell('C3', 'target col')}</row>
    <row r="4">${textCell('B4', 'second row')}${textCell('C4', 'range end')}</row>
  </sheetData>
</worksheet>`;

const linksSheetRels = `${XML_HEADER}
<Relationships xmlns="${PACKAGE_REL_NS}">
  <Relationship Id="rId1" Type="${REL_NS}/hyperlink" Target="https://example.invalid/report" TargetMode="External"/>
  <Relationship Id="rId2" Type="${REL_NS}/hyperlink" Target="javascript:alert(1)" TargetMode="External"/>
  <Relationship Id="rId3" Type="${REL_NS}/comments" Target="../comments1.xml"/>
  <Relationship Id="rId4" Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment" Target="../threadedComments/threadedComment1.xml"/>
</Relationships>`;

const commentsXml = `${XML_HEADER}
<comments xmlns="${SPREADSHEET_NS}">
  <authors><author>Ana Silva</author></authors>
  <commentList>
    <comment ref="C6" authorId="0"><text><t>Legacy note for the popup.</t></text></comment>
  </commentList>
</comments>`;

const threadedCommentsXml = `${XML_HEADER}
<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <threadedComment ref="C8" dT="2026-11-18T09:15:00Z" personId="{11111111-1111-1111-1111-111111111111}" id="{22222222-2222-2222-2222-222222222222}">
    <text>Threaded note for the popup.</text>
  </threadedComment>
  <threadedComment ref="C9" dT="2026-11-18T10:02:00Z" personId="{33333333-3333-3333-3333-333333333333}" id="{44444444-4444-4444-4444-444444444444}" parentId="{22222222-2222-2222-2222-222222222222}">
    <text>Reply with its own author.</text>
  </threadedComment>
</ThreadedComments>`;

const personsXml = `${XML_HEADER}
<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <person displayName="Ana Silva" id="{11111111-1111-1111-1111-111111111111}" userId="ana@example.invalid" providerId="None"/>
  <person displayName="Bo Chen" id="{33333333-3333-3333-3333-333333333333}" userId="bo@example.invalid" providerId="None"/>
</personList>`;

export function linksWorkbookXlsx(): Uint8Array {
  return buildMultiSheetXlsx(
    [
      { name: 'Links', worksheetXml: linksSheetXml },
      { name: 'Target Sheet', worksheetXml: targetSheetXml },
    ],
    {
      extraParts: {
        'xl/worksheets/_rels/sheet1.xml.rels': linksSheetRels,
        'xl/comments1.xml': commentsXml,
        'xl/threadedComments/threadedComment1.xml': threadedCommentsXml,
        'xl/persons/person.xml': personsXml,
      },
      workbookRelsExtras: `<Relationship Id="rId9" Type="http://schemas.microsoft.com/office/2017/10/relationships/person" Target="persons/person.xml"/>`,
    },
  );
}

interface OpenedLink {
  readonly url: string;
  readonly target: string;
  readonly features: string;
}

/** Record every `window.open` call instead of letting the browser act on it. */
async function interceptWindowOpen(page: Page): Promise<void> {
  await page.addInitScript(() => {
    const calls: { url: string; target: string; features: string }[] = [];
    (window as unknown as { __openedLinks: typeof calls }).__openedLinks = calls;
    window.open = ((url?: string | URL, target?: string, features?: string) => {
      calls.push({ url: String(url ?? ''), target: target ?? '', features: features ?? '' });
      return null;
    }) as typeof window.open;
  });
}

async function openedLinks(page: Page): Promise<readonly OpenedLink[]> {
  return page.evaluate(
    () => (window as unknown as { __openedLinks: OpenedLink[] }).__openedLinks ?? [],
  );
}

async function openLinksSheet(page: Page): Promise<void> {
  await interceptWindowOpen(page);
  await page.goto('/?fsa=0');
  await page.getByTestId('file-input').setInputFiles({
    name: 'links.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: Buffer.from(linksWorkbookXlsx()),
  });
  await expect(page.getByTestId('file-name')).toHaveText('links.xlsx');
  await expect(page.getByTestId('grid-canvas')).toBeVisible();
}

async function cellRect(page: Page, row: number, col: number): Promise<CellRect> {
  const rect = await page.evaluate(
    ([targetRow, targetCol]) =>
      window.__calcplus.getCellRect(targetRow as number, targetCol as number),
    [row, col] as const,
  );
  if (rect === null) throw new Error(`No rectangle for row ${row}, column ${col}`);
  return rect;
}

async function cellPoint(page: Page, row: number, col: number) {
  const box = await page.getByTestId('grid-canvas').boundingBox();
  if (box === null) throw new Error('The grid canvas has no box');
  const rect = await cellRect(page, row, col);
  return {
    x: box.x + rect.x + rect.width / 2,
    y: box.y + rect.y + rect.height / 2,
  };
}

/** The comment indicator is the small triangle at the cell's top-right corner. */
async function indicatorPoint(page: Page, row: number, col: number) {
  const box = await page.getByTestId('grid-canvas').boundingBox();
  if (box === null) throw new Error('The grid canvas has no box');
  const rect = await cellRect(page, row, col);
  return { x: box.x + rect.x + rect.width - 3, y: box.y + rect.y + 3 };
}

async function ctrlClickCell(page: Page, row: number, col: number): Promise<void> {
  const point = await cellPoint(page, row, col);
  await page.keyboard.down('Control');
  await page.mouse.click(point.x, point.y);
  await page.keyboard.up('Control');
}

interface SheetDebugApi {
  getState(): { sheetId: string | null };
  getSheets(): { id: string; name: string }[];
}

interface SelectionDebugApi {
  getState(): {
    selection: {
      anchor: { row: number; col: number };
      focus: { row: number; col: number };
      extent: {
        kind: string;
        range?: { startRow: number; startCol: number; endRow: number; endCol: number };
      };
    };
  };
}

async function activeSheetName(page: Page): Promise<string | null> {
  return page.evaluate(() => {
    const api = window.__calcplus as unknown as SheetDebugApi;
    const state = api.getState();
    const sheet = api.getSheets().find((entry) => entry.id === state.sheetId);
    return sheet?.name ?? null;
  });
}

/** Count pixels close to a colour inside a cell's rectangle. */
async function countColourPixels(
  page: Page,
  rect: CellRect,
  rgb: readonly [number, number, number],
  tolerance = 30,
): Promise<number> {
  return page.evaluate(
    ({ x, y, width, height, target, slack }) => {
      const canvas = document.querySelector<HTMLCanvasElement>('[data-testid="grid-canvas"]');
      if (canvas === null) throw new Error('No grid canvas');
      const context = canvas.getContext('2d');
      if (context === null) throw new Error('No canvas context');
      const ratio = canvas.width / canvas.getBoundingClientRect().width;
      const data = context.getImageData(
        Math.round(x * ratio),
        Math.round(y * ratio),
        Math.max(1, Math.round(width * ratio)),
        Math.max(1, Math.round(height * ratio)),
      ).data;
      let count = 0;
      for (let index = 0; index < data.length; index += 4) {
        if ((data[index + 3] ?? 0) < 200) continue;
        if (
          Math.abs((data[index] ?? 0) - target[0]) <= slack &&
          Math.abs((data[index + 1] ?? 0) - target[1]) <= slack &&
          Math.abs((data[index + 2] ?? 0) - target[2]) <= slack
        ) {
          count += 1;
        }
      }
      return count;
    },
    {
      x: rect.x + 2,
      y: rect.y + 2,
      width: rect.width - 4,
      height: rect.height - 4,
      target: rgb,
      slack: tolerance,
    },
  );
}

test('a linked cell paints in the link colour with an underline', async ({ page }) => {
  await openLinksSheet(page);
  // Move the selection away so the active-cell tint cannot confuse the
  // comparison; A1 and B1 then hold the same text and the same style.
  const away = await cellPoint(page, 19, 5);
  await page.mouse.click(away.x, away.y);
  await expect(page.getByTestId('name-box')).toHaveValue('F20');

  const plain = await cellRect(page, 0, 0);
  const linked = await cellRect(page, 0, 1);
  // The theme accent (#107c41) paints the glyphs and the underline, so the
  // linked cell carries accent pixels and the plain one carries none.
  await expect.poll(() => countColourPixels(page, linked, [16, 124, 65])).toBeGreaterThan(0);
  expect(await countColourPixels(page, plain, [16, 124, 65])).toBe(0);
  // The underline changes the painted region, so the two cells differ even
  // before the colour comparison.
  expect(await regionSignature(page, linked.x, linked.y, linked.width, linked.height)).not.toBe(
    await regionSignature(page, plain.x, plain.y, plain.width, plain.height),
  );
});

test('Ctrl+click opens an allowed external link in a new tab', async ({ page }) => {
  await openLinksSheet(page);
  const before = await page.evaluate(() => window.__calcplus.getState().selection);

  await ctrlClickCell(page, 0, 1);

  await expect.poll(async () => (await openedLinks(page)).length).toBe(1);
  const [call] = await openedLinks(page);
  expect(call?.url).toBe('https://example.invalid/report');
  expect(call?.target).toBe('_blank');
  expect(call?.features).toContain('noopener');
  // Following a link never moves the selection.
  expect(await page.evaluate(() => window.__calcplus.getState().selection)).toEqual(before);
  await expect(page.getByTestId('status-message')).toContainText('new tab');
});

test('a refused scheme shows a message and never navigates the app', async ({ page }) => {
  await openLinksSheet(page);
  const url = page.url();
  const dialogs: string[] = [];
  page.on('dialog', async (dialog) => {
    dialogs.push(dialog.message());
    await dialog.dismiss();
  });

  await ctrlClickCell(page, 1, 0);

  await expect(page.getByTestId('status-message')).toContainText('not allowed');
  await expect(page.getByTestId('status-message')).toContainText('only http, https and mailto');
  expect(await openedLinks(page)).toHaveLength(0);
  expect(dialogs).toHaveLength(0);
  expect(page.url()).toBe(url);
  await expect(page.getByTestId('grid-canvas')).toBeVisible();
});

test('Ctrl+click an internal link activates the sheet and selects the cell', async ({ page }) => {
  await openLinksSheet(page);
  expect(await activeSheetName(page)).toBe('Links');

  await ctrlClickCell(page, 2, 0);

  await expect.poll(() => activeSheetName(page)).toBe('Target Sheet');
  const state = await page.evaluate(() => window.__calcplus.getState());
  expect(state.selection.anchor).toEqual({ row: 2, col: 1 });
  expect(state.selection.focus).toEqual({ row: 2, col: 1 });
  await expect(page.getByTestId('name-box')).toHaveValue('B3');
  await expect(page.getByTestId('status-message')).toContainText(
    'Followed the link to Target Sheet!B3',
  );
});

test('Ctrl+click an internal range link selects the whole range', async ({ page }) => {
  await openLinksSheet(page);

  await ctrlClickCell(page, 3, 0);

  await expect.poll(() => activeSheetName(page)).toBe('Target Sheet');
  const state = await page.evaluate(() =>
    (window.__calcplus as unknown as SelectionDebugApi).getState(),
  );
  // Ticket 31 keeps the active cell at the gesture origin; the extent carries
  // the whole jumped range.
  expect(state.selection.anchor).toEqual({ row: 2, col: 1 });
  expect(state.selection.extent).toEqual({
    kind: 'range',
    range: { startRow: 2, startCol: 1, endRow: 3, endCol: 2 },
  });
  await expect(page.getByTestId('name-box')).toHaveValue('B3:C4');
});

test('an unresolvable internal location reports instead of failing silently', async ({ page }) => {
  await openLinksSheet(page);

  await ctrlClickCell(page, 4, 0);

  await expect(page.getByTestId('status-message')).toContainText('could not be resolved');
  await expect(page.getByTestId('status-message')).toContainText('Missing!A1');
  expect(await activeSheetName(page)).toBe('Links');
  // The selection stays where it was: A1 unless a click moved it.
  const state = await page.evaluate(() => window.__calcplus.getState());
  expect(state.selection.anchor).toEqual({ row: 0, col: 0 });
});

test('an internal jump reveals a cell far down the target sheet', async ({ page }) => {
  await openLinksSheet(page);

  await ctrlClickCell(page, 5, 0);

  await expect.poll(() => activeSheetName(page)).toBe('Target Sheet');
  const state = await page.evaluate(() => window.__calcplus.getState());
  expect(state.selection.anchor).toEqual({ row: 39, col: 0 });
  await expect(page.getByTestId('name-box')).toHaveValue('A40');
  // The jump scrolls the target into view instead of leaving it off screen.
  await expect.poll(() => page.evaluate(() => window.__calcplus.getScroll().y)).toBeGreaterThan(0);
  const rect = await cellRect(page, 39, 0);
  const box = await page.getByTestId('grid-canvas').boundingBox();
  if (box === null) throw new Error('The grid canvas has no box');
  expect(rect.y).toBeGreaterThan(0);
  expect(rect.y + rect.height).toBeLessThanOrEqual(box.height);
});

test('a plain click on a link cell selects it like any other cell', async ({ page }) => {
  await openLinksSheet(page);
  const point = await cellPoint(page, 0, 1);

  await page.mouse.click(point.x, point.y);

  await expect
    .poll(() => page.evaluate(() => window.__calcplus.getState().selection.anchor))
    .toEqual({ row: 0, col: 1 });
  await expect(page.getByTestId('name-box')).toHaveValue('B1');
  expect(await openedLinks(page)).toHaveLength(0);
});

test('hovering a linked cell shows its tooltip and moving away hides it', async ({ page }) => {
  await openLinksSheet(page);
  const link = await cellPoint(page, 0, 1);

  await page.mouse.move(link.x, link.y);

  const tooltip = page.getByTestId('link-tooltip');
  await expect(tooltip).toBeVisible();
  await expect(tooltip).toHaveText('Open the report');

  const away = await cellPoint(page, 12, 4);
  await page.mouse.move(away.x, away.y);
  await expect(tooltip).toBeHidden();
});

test('hovering the comment indicator shows the legacy note and its author', async ({ page }) => {
  await openLinksSheet(page);
  const indicator = await indicatorPoint(page, 5, 2);

  await page.mouse.move(indicator.x, indicator.y);

  const popup = page.getByTestId('comment-popup');
  await expect(popup).toBeVisible();
  await expect(page.getByTestId('comment-popup-author')).toHaveText('Ana Silva');
  await expect(page.getByTestId('comment-popup-text')).toHaveText('Legacy note for the popup.');
  await expect(page.getByTestId('comment-popup-badge')).toHaveCount(0);
});

test('a threaded comment and its reply each show their own text and author', async ({ page }) => {
  await openLinksSheet(page);

  const root = await indicatorPoint(page, 7, 2);
  await page.mouse.move(root.x, root.y);
  await expect(page.getByTestId('comment-popup-text')).toHaveText('Threaded note for the popup.');
  await expect(page.getByTestId('comment-popup-author')).toHaveText('Ana Silva');
  await expect(page.getByTestId('comment-popup-badge')).toHaveText('Threaded comment');

  const reply = await indicatorPoint(page, 8, 2);
  await page.mouse.move(reply.x, reply.y);
  await expect(page.getByTestId('comment-popup-text')).toHaveText('Reply with its own author.');
  await expect(page.getByTestId('comment-popup-author')).toHaveText('Bo Chen');
});

test('the comment popup dismisses on Esc and on mouse-away', async ({ page }) => {
  await openLinksSheet(page);
  // Focus the grid first, so the Esc key reaches the shell.
  const focusPoint = await cellPoint(page, 19, 5);
  await page.mouse.click(focusPoint.x, focusPoint.y);

  const indicator = await indicatorPoint(page, 5, 2);
  await page.mouse.move(indicator.x, indicator.y);
  const popup = page.getByTestId('comment-popup');
  await expect(popup).toBeVisible();

  await page.keyboard.press('Escape');
  await expect(popup).toBeHidden();

  await page.mouse.move(indicator.x, indicator.y);
  await expect(popup).toBeVisible();
  const away = await cellPoint(page, 12, 4);
  await page.mouse.move(away.x, away.y);
  await expect(popup).toBeHidden();
});

test('the comment popup tracks the sheet while it scrolls', async ({ page }) => {
  await openLinksSheet(page);
  const indicator = await indicatorPoint(page, 7, 2);
  await page.mouse.move(indicator.x, indicator.y);

  const popup = page.getByTestId('comment-popup');
  await expect(popup).toBeVisible();
  const before = await popup.boundingBox();
  const box = await page.getByTestId('grid-canvas').boundingBox();
  if (before === null || box === null) throw new Error('Missing box');

  await page.mouse.wheel(0, 40);

  // The cell moves up by the scroll offset and the popup follows it exactly.
  await expect
    .poll(async () => {
      const rect = await cellRect(page, 7, 2);
      const after = await popup.boundingBox();
      if (after === null) return Number.NaN;
      return Math.round(after.y - box.y - rect.y);
    })
    .toBe(0);
  const after = await popup.boundingBox();
  if (after === null) throw new Error('The popup closed while scrolling');
  expect(after.y).toBeLessThan(before.y);
  expect(after.y).toBeLessThan(indicator.y);
});
