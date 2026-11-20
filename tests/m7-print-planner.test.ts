import { describe, expect, it } from 'vitest';
import { DEFAULT_ALIGNMENT, DEFAULT_BORDER, DEFAULT_FILL, DEFAULT_FONT } from '@model/style';
import type {
  BorderEdgeStyle,
  BorderStyle,
  GridRange,
  PrintOptionsRecord,
  PrintSetupRecord,
} from '@model/types';
import type { CellValue } from '@model/value';
import type { MergeSpan, ProjectedStyle } from '@view/types';
import { createPrintPlan } from '../src/print/planner';
import type {
  PrintCell,
  PrintFillItem,
  PrintItem,
  PrintLineItem,
  PrintSheetSnapshot,
  PrintTextItem,
} from '../src/print/types';

/**
 * M7 lane 1 — the pure print planner (print-pdf §2, §4).
 *
 * The snapshots here are hand-built so every expected number follows from the
 * fixture's own constants; the adapter tests exercise the real projection and
 * number formats.
 */

const LETTER_WIDTH = 612;
const LETTER_HEIGHT = 792;
const DEFAULT_LEFT = 0.7 * 72;
const DEFAULT_TOP = 0.75 * 72;

function styleOf(overrides: Partial<ProjectedStyle> = {}): ProjectedStyle {
  return {
    font: DEFAULT_FONT,
    fill: DEFAULT_FILL,
    border: DEFAULT_BORDER,
    alignment: DEFAULT_ALIGNMENT,
    numberFormat: null,
    numberFormatColor: null,
    ...overrides,
  };
}

function setupOf(overrides: Partial<PrintSetupRecord> = {}): PrintSetupRecord {
  return {
    margins: null,
    options: null,
    orientation: null,
    fitToWidth: null,
    fitToHeight: null,
    scale: null,
    fitToPage: null,
    paperSize: null,
    ...overrides,
  };
}

function printOptionsOf(overrides: Partial<PrintOptionsRecord> = {}): PrintOptionsRecord {
  return {
    gridLines: false,
    gridLinesSet: false,
    headings: false,
    horizontalCentered: false,
    verticalCentered: false,
    ...overrides,
  };
}

interface CellSpec {
  readonly row: number;
  readonly col: number;
  readonly text?: string;
  readonly value?: CellValue;
  readonly style?: Partial<ProjectedStyle>;
  readonly hyperlink?: boolean;
}

interface SnapshotSpec {
  readonly rows?: number;
  readonly cols?: number;
  readonly rowHeight?: number | ((row: number) => number);
  readonly columnWidth?: number | ((col: number) => number);
  readonly cells?: readonly CellSpec[];
  readonly merges?: readonly MergeSpan[];
  readonly printAreas?: readonly GridRange[];
  readonly usedRange?: GridRange | null;
  readonly setup?: Partial<PrintSetupRecord> | null;
}

function createSnapshot(spec: SnapshotSpec): PrintSheetSnapshot {
  const rows = spec.rows ?? 10;
  const cols = spec.cols ?? 3;
  const merges = spec.merges ?? [];
  const mergeByCell = new Map<string, MergeSpan>();
  for (const merge of merges) {
    for (let row = merge.anchorRow; row < merge.anchorRow + merge.rowSpan; row += 1) {
      for (let col = merge.anchorCol; col < merge.anchorCol + merge.colSpan; col += 1) {
        mergeByCell.set(`${row}:${col}`, merge);
      }
    }
  }
  const cells = new Map<string, PrintCell>();
  let minRow = Number.POSITIVE_INFINITY;
  let maxRow = -1;
  let minCol = Number.POSITIVE_INFINITY;
  let maxCol = -1;
  for (const cell of spec.cells ?? []) {
    const text = cell.text ?? '';
    cells.set(`${cell.row}:${cell.col}`, {
      row: cell.row,
      col: cell.col,
      value: cell.value ?? (text === '' ? { type: 'blank' } : { type: 'text', value: text }),
      displayText: text,
      style: styleOf(cell.style),
      hyperlink: cell.hyperlink === true,
      merge: mergeByCell.get(`${cell.row}:${cell.col}`) ?? null,
      layoutText: () => ({ text, overflow: false }),
    });
    minRow = Math.min(minRow, cell.row);
    maxRow = Math.max(maxRow, cell.row);
    minCol = Math.min(minCol, cell.col);
    maxCol = Math.max(maxCol, cell.col);
  }
  // The sheet's used range defaults to the whole test grid; a test that wants
  // an empty or partial sheet passes `usedRange` explicitly.
  const usedRange =
    spec.usedRange === undefined
      ? { startRow: 0, startCol: 0, endRow: rows - 1, endCol: cols - 1 }
      : spec.usedRange;
  const rowHeight = spec.rowHeight ?? 15;
  const columnWidth = spec.columnWidth ?? 60;

  return {
    sheetId: 'sheet-1',
    sheetName: 'Report',
    printAreas: spec.printAreas ?? [],
    usedRange,
    printSetup: spec.setup === undefined || spec.setup === null ? null : setupOf(spec.setup),
    rowHeight: (row) => (typeof rowHeight === 'function' ? rowHeight(row) : rowHeight),
    columnWidth: (col) => (typeof columnWidth === 'function' ? columnWidth(col) : columnWidth),
    columnLabel: (col) => columnLabel(col),
    rowLabel: (row) => String(row + 1),
    headingFont: DEFAULT_FONT,
    cellAt: (row, col) => cells.get(`${row}:${col}`) ?? null,
    cellsIn: (startRow, endRow, startCol, endCol) =>
      [...cells.values()].filter(
        (cell) =>
          cell.row >= startRow && cell.row <= endRow && cell.col >= startCol && cell.col <= endCol,
      ),
    mergesIn: (startRow, endRow, startCol, endCol) =>
      merges.filter(
        (merge) =>
          merge.anchorRow + merge.rowSpan - 1 >= startRow &&
          merge.anchorRow <= endRow &&
          merge.anchorCol + merge.colSpan - 1 >= startCol &&
          merge.anchorCol <= endCol,
      ),
  };
}

function columnLabel(col: number): string {
  let label = '';
  let value = col + 1;
  while (value > 0) {
    const remainder = (value - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    value = Math.floor((value - 1) / 26);
  }
  return label;
}

function texts(items: readonly PrintItem[]): PrintTextItem[] {
  return items.filter((item): item is PrintTextItem => item.kind === 'text');
}

function fills(items: readonly PrintItem[]): PrintFillItem[] {
  return items.filter((item): item is PrintFillItem => item.kind === 'fill');
}

function lines(items: readonly PrintItem[]): PrintLineItem[] {
  return items.filter((item): item is PrintLineItem => item.kind === 'line');
}

function edge(style: string): BorderEdgeStyle {
  return { style, color: '#000000', weight: style === 'thick' ? 'thick' : 'thin' };
}

function thickBorder(): BorderStyle {
  return {
    left: edge('thick'),
    right: edge('thick'),
    top: edge('thick'),
    bottom: edge('thick'),
    diagonal: { style: 'none', color: null, weight: null },
    diagonalUp: false,
    diagonalDown: false,
  };
}

// ---------------------------------------------------------------------------
// Page box, margins and bands
// ---------------------------------------------------------------------------

describe('page box, margins and bands', () => {
  const plan = createPrintPlan(
    createSnapshot({
      rows: 10,
      cols: 3,
      cells: [{ row: 0, col: 0, text: 'Name' }],
    }),
  );

  it('uses the default Letter paper and Excel margins in points', () => {
    expect(plan.paper).toEqual({
      name: 'Letter',
      code: null,
      orientation: 'portrait',
      width: LETTER_WIDTH,
      height: LETTER_HEIGHT,
    });
    expect(plan.margins).toEqual({
      top: DEFAULT_TOP,
      right: DEFAULT_LEFT,
      bottom: DEFAULT_TOP,
      left: DEFAULT_LEFT,
    });
  });

  it('emits one page whose box and body are in points', () => {
    expect(plan.pages).toHaveLength(1);
    const page = plan.pages[0]!;
    expect(page.box).toEqual({ x: 0, y: 0, width: LETTER_WIDTH, height: LETTER_HEIGHT });
    expect(page.body).toEqual({
      x: DEFAULT_LEFT,
      y: DEFAULT_TOP,
      width: LETTER_WIDTH - DEFAULT_LEFT * 2,
      height: LETTER_HEIGHT - DEFAULT_TOP * 2,
    });
    expect(page.band).toEqual({ startRow: 0, endRow: 9, startCol: 0, endCol: 2 });
    expect(page.clipped).toBe(false);
  });

  it('positions the row and column bands from the body origin', () => {
    const page = plan.pages[0]!;
    expect(page.columns.map((column) => column.x)).toEqual([
      DEFAULT_LEFT,
      DEFAULT_LEFT + 60,
      DEFAULT_LEFT + 120,
    ]);
    expect(page.rows[0]!.y).toBe(DEFAULT_TOP);
    expect(page.rows[1]!.y).toBe(DEFAULT_TOP + 15);
    expect(page.rows[9]!.y + page.rows[9]!.height).toBe(DEFAULT_TOP + 150);
  });

  it('swaps the paper box for landscape', () => {
    const landscape = createPrintPlan(
      createSnapshot({ rows: 2, cols: 2, setup: { orientation: 'landscape', paperSize: 9 } }),
    );
    expect(landscape.paper).toMatchObject({
      name: 'A4',
      orientation: 'landscape',
      width: 841.89,
      height: 595.28,
    });
    expect(landscape.pages[0]!.box.width).toBe(841.89);
  });

  it('pages down then over', () => {
    const paged = createPrintPlan(
      createSnapshot({
        rows: 80,
        cols: 30,
        setup: {
          margins: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
        },
      }),
    );
    expect(paged.pages.length).toBeGreaterThan(2);
    // Page 2 is directly below page 1: the column band is unchanged.
    expect(paged.pages[1]!.band.startCol).toBe(paged.pages[0]!.band.startCol);
    expect(paged.pages[1]!.band.startRow).toBe(paged.pages[0]!.band.endRow + 1);
    // After the first column, the next page starts at the top of the sheet.
    const firstColumnPageCount = paged.pages.filter(
      (page) => page.band.startCol === paged.pages[0]!.band.startCol,
    ).length;
    const next = paged.pages[firstColumnPageCount]!;
    expect(next.band.startCol).toBeGreaterThan(paged.pages[0]!.band.endCol);
    expect(next.band.startRow).toBe(0);
  });
});

// ---------------------------------------------------------------------------
// Scale
// ---------------------------------------------------------------------------

describe('uniform scale and fit-to-width', () => {
  const wide = { rows: 10, cols: 40 } as const;

  it('keeps 100% when the sheet declares no fit-to-page', () => {
    const plan = createPrintPlan(createSnapshot(wide));
    expect(plan.scale).toBe(1);
    expect(plan.fitToPage).toBe(false);
    // 40 x 60pt = 2400pt against a 511.2pt Letter body: five page widths.
    expect(new Set(plan.pages.map((page) => page.band.startCol)).size).toBe(5);
  });

  it('shrinks every column onto one page width when fit-to-width is on', () => {
    const plan = createPrintPlan(
      createSnapshot({
        ...wide,
        setup: {
          paperSize: 9,
          fitToPage: true,
          fitToWidth: 1,
          fitToHeight: 0,
        },
      }),
    );
    const bodyWidth = 595.28 - DEFAULT_LEFT * 2;
    expect(plan.fitToPage).toBe(true);
    expect(plan.scale).toBeLessThanOrEqual(bodyWidth / 2400);
    expect(plan.scale).toBeGreaterThan((bodyWidth / 2400) * 0.95);
    expect(plan.pages).toHaveLength(1);
    expect(plan.pages[0]!.band).toMatchObject({ startCol: 0, endCol: 39 });
    // The scale is applied to the geometry, not just recorded.
    expect(plan.pages[0]!.columns[39]!.x + plan.pages[0]!.columns[39]!.width).toBeLessThanOrEqual(
      bodyWidth + DEFAULT_LEFT + 0.01,
    );
  });

  it('does not stretch a small sheet when fit-to-width is on', () => {
    const plan = createPrintPlan(
      createSnapshot({
        rows: 3,
        cols: 2,
        setup: { fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
      }),
    );
    expect(plan.scale).toBe(1);
  });

  it('honours fit-to-height and the authored scale when fit-to-page is off', () => {
    const tall = createPrintPlan(
      createSnapshot({
        rows: 100,
        cols: 2,
        setup: { fitToPage: true, fitToWidth: 0, fitToHeight: 1 },
      }),
    );
    expect(tall.pages).toHaveLength(1);
    expect(tall.pages[0]!.band).toMatchObject({ startRow: 0, endRow: 99 });

    const scaled = createPrintPlan(
      createSnapshot({ rows: 10, cols: 2, setup: { scale: 150, fitToPage: false } }),
    );
    expect(scaled.scale).toBe(1.5);
    expect(scaled.pages[0]!.rows[1]!.height).toBe(22.5);
  });

  it('clamps the authored scale to Excel 10%-400%', () => {
    expect(createPrintPlan(createSnapshot({ setup: { scale: 5 } })).scale).toBe(0.1);
    expect(createPrintPlan(createSnapshot({ setup: { scale: 900 } })).scale).toBe(4);
  });

  it('clips a column wider than the body and says so', () => {
    const plan = createPrintPlan(
      createSnapshot({
        rows: 2,
        cols: 1,
        columnWidth: 900,
        cells: [{ row: 0, col: 0, text: 'x' }],
      }),
    );
    expect(plan.pages).toHaveLength(1);
    expect(plan.pages[0]!.clipped).toBe(true);
    expect(plan.pages[0]!.columns[0]!.width).toBe(900);
  });

  it('refuses a plan beyond the page cap instead of hanging', () => {
    expect(() => createPrintPlan(createSnapshot({ rows: 200_000, cols: 1 }))).toThrow(
      /calcplus prints at most 2000/,
    );
  });
});

// ---------------------------------------------------------------------------
// Margins
// ---------------------------------------------------------------------------

describe('margins', () => {
  it('move the body box and change how much fits', () => {
    const tight = createPrintPlan(
      createSnapshot({
        rows: 10,
        cols: 2,
        columnWidth: 300,
        setup: {
          margins: { left: 2, right: 2, top: 2, bottom: 2, header: 0.3, footer: 0.3 },
        },
      }),
    );
    expect(tight.pages[0]!.body).toEqual({
      x: 144,
      y: 144,
      width: LETTER_WIDTH - 288,
      height: LETTER_HEIGHT - 288,
    });
    // 300pt columns do not fit side by side in 324pt of body.
    expect(tight.pages.every((page) => page.band.startCol === page.band.endCol)).toBe(true);
    expect(tight.pages).toHaveLength(2);
  });

  it('centres the content block when printOptions asks', () => {
    const plan = createPrintPlan(
      createSnapshot({
        rows: 2,
        cols: 2,
        setup: {
          options: printOptionsOf({ horizontalCentered: true, verticalCentered: true }),
        },
      }),
    );
    const page = plan.pages[0]!;
    const blockWidth = 120;
    const blockHeight = 30;
    expect(page.columns[0]!.x).toBeCloseTo(
      DEFAULT_LEFT + (LETTER_WIDTH - DEFAULT_LEFT * 2 - blockWidth) / 2,
      6,
    );
    expect(page.rows[0]!.y).toBeCloseTo(
      DEFAULT_TOP + (LETTER_HEIGHT - DEFAULT_TOP * 2 - blockHeight) / 2,
      6,
    );
  });
});

// ---------------------------------------------------------------------------
// Print areas
// ---------------------------------------------------------------------------

describe('print area resolution', () => {
  it('prints each defined area as its own page set', () => {
    const plan = createPrintPlan(
      createSnapshot({
        rows: 10,
        cols: 8,
        cells: [
          { row: 0, col: 0, text: 'One' },
          { row: 4, col: 4, text: 'Two' },
        ],
        printAreas: [
          { startRow: 0, startCol: 0, endRow: 1, endCol: 1 },
          { startRow: 4, startCol: 4, endRow: 5, endCol: 5 },
        ],
      }),
    );
    expect(plan.pageSets).toEqual([
      {
        index: 0,
        range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 },
        firstPage: 0,
        pageCount: 1,
      },
      {
        index: 1,
        range: { startRow: 4, startCol: 4, endRow: 5, endCol: 5 },
        firstPage: 1,
        pageCount: 1,
      },
    ]);
    expect(plan.pages.map((page) => page.band)).toEqual([
      { startRow: 0, endRow: 1, startCol: 0, endCol: 1 },
      { startRow: 4, endRow: 5, startCol: 4, endCol: 5 },
    ]);
    expect(plan.pages[0]!.setIndex).toBe(0);
    expect(plan.pages[1]!.setIndex).toBe(1);
    expect(plan.pages[1]!.pageInSetIndex).toBe(0);
  });

  it('falls back to the used range when no area is defined', () => {
    const plan = createPrintPlan(
      createSnapshot({
        rows: 10,
        cols: 4,
        cells: [
          { row: 2, col: 1, text: 'A' },
          { row: 4, col: 2, text: 'B' },
        ],
        usedRange: { startRow: 2, startCol: 1, endRow: 4, endCol: 2 },
      }),
    );
    expect(plan.pageSets).toHaveLength(1);
    expect(plan.pageSets[0]!.range).toEqual({
      startRow: 2,
      startCol: 1,
      endRow: 4,
      endCol: 2,
    });
    expect(plan.pages[0]!.band).toEqual({ startRow: 2, endRow: 4, startCol: 1, endCol: 2 });
    expect(plan.pages[0]!.columns.map((column) => column.col)).toEqual([1, 2]);
  });

  it('is empty when the sheet is empty', () => {
    const plan = createPrintPlan(createSnapshot({ usedRange: null }));
    expect(plan.pages).toEqual([]);
    expect(plan.pageSets).toEqual([]);
    expect(plan.scale).toBe(1);
  });

  it('drops hidden rows and columns from the flow', () => {
    const plan = createPrintPlan(
      createSnapshot({
        rows: 5,
        cols: 3,
        rowHeight: (row) => (row === 1 ? 0 : 15),
        columnWidth: (col) => (col === 0 ? 0 : 60),
        usedRange: { startRow: 0, startCol: 0, endRow: 4, endCol: 2 },
      }),
    );
    const page = plan.pages[0]!;
    expect(page.rows.map((row) => row.row)).toEqual([0, 2, 3, 4]);
    expect(page.columns.map((column) => column.col)).toEqual([1, 2]);
    // The band is the printed (visible) span: the hidden column is not in it.
    expect(page.band).toEqual({ startRow: 0, endRow: 4, startCol: 1, endCol: 2 });
  });
});

// ---------------------------------------------------------------------------
// Gridlines and headings
// ---------------------------------------------------------------------------

describe('gridlines and headings follow printOptions', () => {
  it('prints gridlines only when gridLines and gridLinesSet are both true', () => {
    const off = createPrintPlan(
      createSnapshot({ setup: { options: printOptionsOf({ gridLines: true }) } }),
    );
    expect(off.showGridLines).toBe(false);
    expect(lines(off.pages[0]!.items)).toEqual([]);

    const on = createPrintPlan(
      createSnapshot({
        setup: { options: printOptionsOf({ gridLines: true, gridLinesSet: true }) },
      }),
    );
    expect(on.showGridLines).toBe(true);
    // 3 columns and 10 rows: four vertical and eleven horizontal lines.
    expect(lines(on.pages[0]!.items)).toHaveLength(4 + 11);
    expect(lines(on.pages[0]!.items)[0]).toMatchObject({
      color: '#d0d0d0',
      width: 0.75,
    });
  });

  it('reserves a heading band and labels each row and column', () => {
    const plan = createPrintPlan(
      createSnapshot({
        rows: 3,
        cols: 2,
        cells: [{ row: 0, col: 0, text: 'A1' }],
        setup: { options: printOptionsOf({ headings: true }) },
      }),
    );
    const page = plan.pages[0]!;
    expect(plan.showHeadings).toBe(true);
    expect(page.columns[0]!.x).toBe(DEFAULT_LEFT + 25);
    expect(page.rows[0]!.y).toBe(DEFAULT_TOP + 15);
    const labels = texts(page.items).map((item) => item.text);
    expect(labels).toEqual(expect.arrayContaining(['A', 'B', '1', '2', '3', 'A1']));
    const letter = texts(page.items).find((item) => item.text === 'A');
    expect(letter!.y).toBeLessThan(page.rows[0]!.y);
  });
});

// ---------------------------------------------------------------------------
// Merged cells
// ---------------------------------------------------------------------------

describe('merged cells clip per page with content on the anchor page', () => {
  const merge: MergeSpan = { anchorRow: 20, anchorCol: 0, rowSpan: 7, colSpan: 2 };
  const plan = createPrintPlan(
    createSnapshot({
      rows: 48,
      cols: 3,
      // 3in top and bottom margins on Letter paper leave 360pt of body: 24 rows.
      setup: {
        margins: { left: 0.7, right: 0.7, top: 3, bottom: 3, header: 0.3, footer: 0.3 },
      },
      merges: [merge],
      cells: [
        {
          row: 20,
          col: 0,
          text: 'Merged',
          style: { fill: { pattern: 'solid', color: '#FFFF00', backgroundColor: null } },
        },
      ],
      printAreas: [{ startRow: 0, startCol: 0, endRow: 47, endCol: 2 }],
    }),
  );

  it('splits the range at the page break', () => {
    expect(plan.pages.length).toBeGreaterThan(1);
    expect(plan.pages[0]!.band).toEqual({ startRow: 0, endRow: 23, startCol: 0, endCol: 2 });
    expect(plan.pages[1]!.band).toMatchObject({ startRow: 24 });
  });

  it('keeps the text on the anchor page only', () => {
    expect(texts(plan.pages[0]!.items).map((item) => item.text)).toContain('Merged');
    expect(texts(plan.pages[1]!.items).map((item) => item.text)).not.toContain('Merged');
  });

  it('clips the merged region to each page and keeps its fill', () => {
    const firstPage = fills(plan.pages[0]!.items).find((item) => item.color === '#FFFF00');
    const secondPage = fills(plan.pages[1]!.items).find((item) => item.color === '#FFFF00');
    // Rows 20-23 on the first page: four 15pt rows.
    expect(firstPage!.box.height).toBe(60);
    // Rows 24-26 on the second page: the rest of the merge.
    expect(secondPage!.box.height).toBe(45);
    // The region is two columns wide.
    expect(firstPage!.box.width).toBe(120);
  });
});

// ---------------------------------------------------------------------------
// Display list
// ---------------------------------------------------------------------------

describe('the display list carries fonts, fills, borders and text', () => {
  const plan = createPrintPlan(
    createSnapshot({
      rows: 3,
      cols: 3,
      cells: [
        {
          row: 0,
          col: 0,
          text: 'Total',
          style: {
            font: { ...DEFAULT_FONT, bold: true, italic: true, size: 14 },
            fill: { pattern: 'solid', color: '#FFFF00', backgroundColor: null },
            border: thickBorder(),
            alignment: { ...DEFAULT_ALIGNMENT, vertical: 'top' },
          },
        },
        {
          row: 1,
          col: 1,
          text: '£1,200.00',
          value: { type: 'number', value: 1200 },
        },
      ],
    }),
  );
  const page = plan.pages[0]!;

  it('paints the fill underneath the text', () => {
    const fill = fills(page.items)[0];
    expect(fill).toMatchObject({
      pattern: 'solid',
      color: '#FFFF00',
      box: { x: DEFAULT_LEFT, y: DEFAULT_TOP, width: 60, height: 15 },
    });
    expect(page.items.indexOf(fill!)).toBeLessThan(
      page.items.indexOf(texts(page.items).find((item) => item.text === 'Total')!),
    );
  });

  it('carries the resolved font and baseline', () => {
    const text = texts(page.items).find((item) => item.text === 'Total')!;
    expect(text.font).toEqual({
      family: 'Calibri',
      size: 14,
      bold: true,
      italic: true,
      color: '#202020',
    });
    expect(text.x).toBe(DEFAULT_LEFT + 3);
    expect(text.y).toBe(DEFAULT_TOP + 1.5 + 14);
    expect(text.anchor).toBe('start');
    expect(text.rotation).toBe(0);
    expect(text.clip).toEqual({ x: DEFAULT_LEFT, y: DEFAULT_TOP, width: 60, height: 15 });
  });

  it('emits one line per declared border edge with the thick weight', () => {
    const border = lines(page.items).filter((line) => line.color === '#000000');
    expect(border).toHaveLength(4);
    expect(border.every((line) => line.width === 2.25)).toBe(true);
  });

  it('right-aligns a number by General alignment', () => {
    const text = texts(page.items).find((item) => item.text === '£1,200.00')!;
    expect(text.anchor).toBe('end');
    expect(text.x).toBe(DEFAULT_LEFT + 120 - 3);
  });
});

describe('borders', () => {
  it('draws a shared edge once', () => {
    const thin = { style: 'thin', color: '#FF0000', weight: 'thin' } as const;
    const border = (left: typeof thin | typeof DEFAULT_BORDER.left): BorderStyle => ({
      ...DEFAULT_BORDER,
      right: left,
      left,
    });
    const plan = createPrintPlan(
      createSnapshot({
        rows: 1,
        cols: 2,
        cells: [
          { row: 0, col: 0, text: 'a', style: { border: border(thin) } },
          { row: 0, col: 1, text: 'b', style: { border: border(thin) } },
        ],
      }),
    );
    const boundary = DEFAULT_LEFT + 60;
    const atBoundary = lines(plan.pages[0]!.items).filter(
      (line) => line.x1 === boundary && line.x2 === boundary,
    );
    expect(atBoundary).toHaveLength(1);
  });

  it('lets the heavier edge win on a shared boundary', () => {
    const thin = { style: 'thin', color: '#FF0000', weight: 'thin' } as const;
    const thick = { style: 'thick', color: '#0000FF', weight: 'thick' } as const;
    const plan = createPrintPlan(
      createSnapshot({
        rows: 1,
        cols: 2,
        cells: [
          { row: 0, col: 0, text: 'a', style: { border: { ...DEFAULT_BORDER, right: thin } } },
          { row: 0, col: 1, text: 'b', style: { border: { ...DEFAULT_BORDER, left: thick } } },
        ],
      }),
    );
    const boundary = DEFAULT_LEFT + 60;
    const atBoundary = lines(plan.pages[0]!.items).filter(
      (line) => line.x1 === boundary && line.x2 === boundary,
    );
    expect(atBoundary).toHaveLength(1);
    expect(atBoundary[0]).toMatchObject({ color: '#0000FF', width: 2.25 });
  });
});
