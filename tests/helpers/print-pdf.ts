import type {
  PrintBox,
  PrintItem,
  PrintPage,
  PrintPaper,
  PrintPlan,
  PrintTextItem,
} from '../../src/print/types';

/** Shared plan fixtures for the M7 lane 3 renderer tests. */

export const A4_PORTRAIT: PrintPaper = {
  name: 'A4',
  code: 9,
  orientation: 'portrait',
  width: 595.28,
  height: 841.89,
};

export const LETTER_LANDSCAPE: PrintPaper = {
  name: 'Letter',
  code: 1,
  orientation: 'landscape',
  width: 792,
  height: 612,
};

export const BODY: PrintBox = { x: 50, y: 54, width: 495, height: 734 };

export function box(x: number, y: number, width: number, height: number): PrintBox {
  return { x, y, width, height };
}

export function textItem(overrides: Partial<PrintTextItem> = {}): PrintTextItem {
  return {
    kind: 'text',
    text: 'Hello',
    x: 60,
    y: 100,
    anchor: 'start',
    font: { family: 'Calibri', size: 11, bold: false, italic: false, color: '#202020' },
    rotation: 0,
    clip: BODY,
    ...overrides,
  };
}

export function makePage(
  index: number,
  items: readonly PrintItem[] = [],
  overrides: Partial<PrintPage> = {},
): PrintPage {
  const paper = overrides.paper ?? A4_PORTRAIT;
  return {
    index,
    setIndex: 0,
    pageInSetIndex: index,
    paper,
    box: { x: 0, y: 0, width: paper.width, height: paper.height },
    body: BODY,
    band: { startRow: 0, endRow: 0, startCol: 0, endCol: 0 },
    columns: [],
    rows: [],
    items,
    clipped: false,
    ...overrides,
  };
}

export function makePlan(
  pages: readonly PrintPage[],
  overrides: Partial<PrintPlan> = {},
): PrintPlan {
  return {
    sheetId: 'sheet-1',
    sheetName: 'Report',
    paper: A4_PORTRAIT,
    margins: { top: 54, right: 50, bottom: 54, left: 50 },
    scale: 1,
    fitToPage: false,
    showGridLines: true,
    showHeadings: false,
    horizontalCentered: false,
    verticalCentered: false,
    pageSets: [
      {
        index: 0,
        range: { startRow: 0, endRow: 0, startCol: 0, endCol: 0 },
        firstPage: 0,
        pageCount: pages.length,
      },
    ],
    pages,
    ...overrides,
  };
}

/** A 1x1 transparent PNG, enough for PDFKit to place a real image XObject. */
export const ONE_PIXEL_PNG = Uint8Array.from(
  Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
    'base64',
  ),
);
