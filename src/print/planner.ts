import type { BorderEdgeStyle, BorderStyle, GridRange, PageMarginsRecord } from '@model/types';
import type { MergeSpan, ProjectedStyle } from '@view/types';
import {
  LINE_SPACING,
  horizontalFor,
  rotationFor,
  textInsets,
  verticalFor,
  wrapTextLines,
} from '@grid/text-layout';
import type {
  PrintBand,
  PrintBox,
  PrintCell,
  PrintFillItem,
  PrintFont,
  PrintItem,
  PrintMargins,
  PrintPage,
  PrintPageColumn,
  PrintPageRow,
  PrintPageSet,
  PrintPaper,
  PrintPlan,
  PrintSheetSnapshot,
  PrintTextAnchor,
  PrintTextItem,
  PrintTextMeasure,
} from './types';

/**
 * The print planner (print-pdf §2, §4).
 *
 * Pure by construction: it takes a plain snapshot of the display records and
 * the sheet's print setup, and returns a `PrintPlan`. It never touches the DOM,
 * a canvas, the network or the workbook model, and the plan is plain data, so
 * an SVG renderer, a PDFKit writer and a 300 DPI canvas fallback all consume
 * the same list.
 *
 * §4 rules owned here:
 *
 * - the print area (`_xlnm.Print_Area`, several areas = several page sets, no
 *   area = the used range);
 * - orientation, fit-to-width (`fitToPage`, `fitToWidth`, `fitToHeight`), a
 *   single uniform scale, and margins in points;
 * - gridlines and headings from `printOptions` only, never the view state;
 * - a merged range clipped per page, with its text on the anchor's page;
 * - text, fills and borders from the same display records the grid paints.
 *
 * The scale is uniform and never distorts the sheet. A row or column larger
 * than the body is clipped at the page edge and the page reports `clipped`.
 */

/** Excel clamps print scaling to 10%-400% (print-pdf research §1.2). */
export const MIN_PRINT_SCALE = 0.1;
export const MAX_PRINT_SCALE = 4;

/** A plan larger than this is refused instead of freezing the tab. */
export const MAX_PRINT_PAGES = 2000;

/** Excel's minimum font size, in points (6 CSS pixels). */
const MIN_FONT_POINTS = 4.5;
/** CSS pixels the grid's text layout uses, converted to points at 96 dpi. */
const PT_PER_PIXEL = 0.75;
const TEXT_PADDING = 4 * PT_PER_PIXEL;
const TOP_PADDING = 2 * PT_PER_PIXEL;
const BOTTOM_PADDING = 1.5 * PT_PER_PIXEL;
/** The row-number column and the column-letter row, in sheet points. */
const HEADING_COLUMN_POINTS = 25;
const HEADING_ROW_POINTS = 15;
const GRIDLINE_COLOR = '#d0d0d0';
const GRIDLINE_WIDTH = 0.75;
/** The app's theme accent, matching the grid's hyperlink colour. */
const LINK_COLOR = '#107c41';
const DEFAULT_TEXT_COLOR = '#202020';
const AUTOMATIC_BORDER_COLOR = '#000000';

/** Excel's defaults for a sheet that declares no `<pageMargins>`. */
const DEFAULT_MARGINS: PrintMargins = {
  top: 0.75 * 72,
  right: 0.7 * 72,
  bottom: 0.75 * 72,
  left: 0.7 * 72,
};

/** OOXML `paperSize` codes; the width and height are portrait points. */
const PAPER_SIZES: Readonly<
  Record<number, { readonly name: string; readonly width: number; readonly height: number }>
> = {
  1: { name: 'Letter', width: 612, height: 792 },
  3: { name: 'Tabloid', width: 792, height: 1224 },
  4: { name: 'Ledger', width: 1224, height: 792 },
  5: { name: 'Legal', width: 612, height: 1008 },
  6: { name: 'Statement', width: 396, height: 612 },
  7: { name: 'Executive', width: 522, height: 756 },
  8: { name: 'A3', width: 841.89, height: 1190.55 },
  9: { name: 'A4', width: 595.28, height: 841.89 },
  10: { name: 'A5', width: 419.53, height: 595.28 },
  11: { name: 'B4', width: 708.66, height: 1000.63 },
  12: { name: 'B5', width: 498.9, height: 708.66 },
  13: { name: 'Folio', width: 612, height: 936 },
  14: { name: 'Quarto', width: 610, height: 780 },
  15: { name: 'Standard 10x14', width: 720, height: 1008 },
  16: { name: 'Standard 11x17', width: 792, height: 1224 },
  17: { name: 'Note', width: 540, height: 720 },
  18: { name: 'Envelope 9', width: 684, height: 918 },
  19: { name: 'Envelope 10', width: 684, height: 792 },
};

const DEFAULT_PAPER_SIZE = 1;

/** A simple average-width estimate, used when the caller supplies no measure. */
export const approximateTextMeasure: PrintTextMeasure = (text, font) =>
  text.length * font.size * 0.5;

export interface PrintPlannerOptions {
  /**
   * Text measurement in sheet points. Omit for an approximation: it affects
   * wrapping and `#` overflow inside a cell, never pagination (row heights are
   * fixed by the sheet).
   */
  readonly measure?: PrintTextMeasure;
}

export class PrintPlanError extends Error {
  readonly code: 'too-many-pages';

  constructor(message: string) {
    super(message);
    this.name = 'PrintPlanError';
    this.code = 'too-many-pages';
  }
}

interface AxisEntry {
  readonly index: number;
  readonly size: number;
}

interface PageGeometry {
  readonly originX: number;
  readonly originY: number;
  readonly scale: number;
  readonly body: PrintBox;
}

/**
 * Resolve the sheet's print setup into a plan. The plan is empty when the
 * sheet has no print area and no used range, or when every row and column in
 * the area is hidden.
 */
export function createPrintPlan(
  sheet: PrintSheetSnapshot,
  options: PrintPlannerOptions = {},
): PrintPlan {
  const measure = options.measure ?? approximateTextMeasure;
  const setup = sheet.printSetup;
  const paper = resolvePaper(setup?.paperSize ?? null, setup?.orientation ?? 'portrait');
  const margins = resolveMargins(setup?.margins ?? null);
  const printOptions = setup?.options ?? null;
  const showGridLines = printOptions?.gridLines === true && printOptions.gridLinesSet === true;
  const showHeadings = printOptions?.headings === true;
  const horizontalCentered = printOptions?.horizontalCentered === true;
  const verticalCentered = printOptions?.verticalCentered === true;
  const fitToPage = setup?.fitToPage === true;
  const ranges = resolveRanges(sheet);

  const box: PrintBox = { x: 0, y: 0, width: paper.width, height: paper.height };
  const body: PrintBox = {
    x: round(margins.left),
    y: round(margins.top),
    width: round(Math.max(0, paper.width - margins.left - margins.right)),
    height: round(Math.max(0, paper.height - margins.top - margins.bottom)),
  };
  const headingWidth = showHeadings ? HEADING_COLUMN_POINTS : 0;
  const headingHeight = showHeadings ? HEADING_ROW_POINTS : 0;

  if (ranges.length === 0) {
    return {
      sheetId: sheet.sheetId,
      sheetName: sheet.sheetName,
      paper,
      margins,
      scale: 1,
      fitToPage,
      showGridLines,
      showHeadings,
      horizontalCentered,
      verticalCentered,
      pageSets: [],
      pages: [],
    };
  }

  // Each range packs independently, but Excel applies one scale to the sheet,
  // so take the smallest scale that satisfies every print area.
  const columnsByRange = ranges.map((range) =>
    axisEntries(range.startCol, range.endCol, (col) => sheet.columnWidth(col)),
  );
  const rowsByRange = ranges.map((range) =>
    axisEntries(range.startRow, range.endRow, (row) => sheet.rowHeight(row)),
  );
  const scale = resolveScale({
    setup,
    fitToPage,
    body,
    headingWidth,
    headingHeight,
    columns: columnsByRange,
    rows: rowsByRange,
  });

  const availableByRange = columnsByRange.map((columns, index) => ({
    columns,
    rows: rowsByRange[index] ?? [],
    columnBands: packBands(columns, contentCapacity(body.width, headingWidth, scale)),
    rowBands: packBands(
      rowsByRange[index] ?? [],
      contentCapacity(body.height, headingHeight, scale),
    ),
  }));

  let pageCount = 0;
  for (const entry of availableByRange) {
    pageCount += entry.columnBands.length * entry.rowBands.length;
  }
  if (pageCount > MAX_PRINT_PAGES) {
    throw new PrintPlanError(
      `The print area needs ${pageCount} pages; calcplus prints at most ${MAX_PRINT_PAGES} at a time. Narrow the print area or fit it to fewer pages.`,
    );
  }

  const pages: PrintPage[] = [];
  const pageSets: PrintPageSet[] = [];
  for (const [setIndex, range] of ranges.entries()) {
    const layout = availableByRange[setIndex];
    if (layout === undefined) continue;
    const firstPage = pages.length;
    for (const columnBand of layout.columnBands) {
      for (const rowBand of layout.rowBands) {
        pages.push(
          buildPage({
            sheet,
            measure,
            paper,
            box,
            body,
            scale,
            setIndex,
            pageInSetIndex: pages.length - firstPage,
            index: pages.length,
            band: bandOf(range, columnBand, rowBand),
            columnBand,
            rowBand,
            showGridLines,
            showHeadings,
            horizontalCentered,
            verticalCentered,
            headingWidth,
            headingHeight,
            headingFont: sheet.headingFont,
          }),
        );
      }
    }
    pageSets.push({
      index: setIndex,
      range,
      firstPage,
      pageCount: pages.length - firstPage,
    });
  }

  return {
    sheetId: sheet.sheetId,
    sheetName: sheet.sheetName,
    paper,
    margins,
    scale,
    fitToPage,
    showGridLines,
    showHeadings,
    horizontalCentered,
    verticalCentered,
    pageSets,
    pages,
  };
}

// ---------------------------------------------------------------------------
// Page setup
// ---------------------------------------------------------------------------

function resolvePaper(code: number | null, orientation: 'portrait' | 'landscape'): PrintPaper {
  const known = code === null ? undefined : PAPER_SIZES[code];
  const size = known ?? PAPER_SIZES[DEFAULT_PAPER_SIZE]!;
  const portraitWidth = size.width;
  const portraitHeight = size.height;
  return {
    name: code !== null && known === undefined ? 'Custom' : size.name,
    code,
    orientation,
    width: orientation === 'landscape' ? portraitHeight : portraitWidth,
    height: orientation === 'landscape' ? portraitWidth : portraitHeight,
  };
}

function resolveMargins(record: PageMarginsRecord | null): PrintMargins {
  if (record === null) return DEFAULT_MARGINS;
  return {
    top: record.top * 72,
    right: record.right * 72,
    bottom: record.bottom * 72,
    left: record.left * 72,
  };
}

/** The content capacity in unscaled sheet points for a body axis. */
function contentCapacity(bodySize: number, headingSize: number, scale: number): number {
  return bodySize / scale - headingSize;
}

// ---------------------------------------------------------------------------
// Ranges, axes and pagination
// ---------------------------------------------------------------------------

/** The authored print areas, else the used range, clamped to the sheet. */
function resolveRanges(sheet: PrintSheetSnapshot): GridRange[] {
  const source =
    sheet.printAreas.length > 0
      ? sheet.printAreas
      : sheet.usedRange === null
        ? []
        : [sheet.usedRange];
  const ranges: GridRange[] = [];
  for (const range of source) {
    const normalised = normaliseRange(range);
    if (normalised !== null) ranges.push(normalised);
  }
  return ranges;
}

function normaliseRange(range: GridRange): GridRange | null {
  const startRow = Math.max(0, Math.min(range.startRow, range.endRow));
  const endRow = Math.max(range.startRow, range.endRow);
  const startCol = Math.max(0, Math.min(range.startCol, range.endCol));
  const endCol = Math.max(range.startCol, range.endCol);
  if (!Number.isFinite(startRow) || !Number.isFinite(endRow)) return null;
  if (!Number.isFinite(startCol) || !Number.isFinite(endCol)) return null;
  return { startRow, startCol, endRow, endCol };
}

function axisEntries(start: number, end: number, sizeOf: (index: number) => number): AxisEntry[] {
  const entries: AxisEntry[] = [];
  const first = Math.max(0, Math.min(start, end));
  const last = Math.max(start, end);
  for (let index = first; index <= last; index += 1) {
    const size = sizeOf(index);
    if (!Number.isFinite(size) || size <= 0) continue;
    entries.push({ index, size });
  }
  return entries;
}

/** Split entries into bands no larger than the capacity. Empty when none fit. */
function packBands(entries: readonly AxisEntry[], capacity: number): AxisEntry[][] {
  const bands: AxisEntry[][] = [];
  let current: AxisEntry[] = [];
  let used = 0;
  for (const entry of entries) {
    if (current.length > 0 && used + entry.size > capacity + 1e-6) {
      bands.push(current);
      current = [];
      used = 0;
    }
    current.push(entry);
    used += entry.size;
  }
  if (current.length > 0) bands.push(current);
  return bands;
}

function countBands(entries: readonly AxisEntry[], capacity: number): number {
  let bands = entries.length === 0 ? 0 : 1;
  let used = 0;
  for (const entry of entries) {
    if (used > 0 && used + entry.size > capacity + 1e-6) {
      bands += 1;
      used = 0;
    }
    used += entry.size;
  }
  return bands;
}

interface ScaleInput {
  readonly setup: PrintSheetSnapshot['printSetup'];
  readonly fitToPage: boolean;
  readonly body: PrintBox;
  readonly headingWidth: number;
  readonly headingHeight: number;
  readonly columns: readonly AxisEntry[][];
  readonly rows: readonly AxisEntry[][];
}

/**
 * §4 scale: one uniform factor for the whole sheet. Fit-to-page searches for
 * the largest factor at or below 100% that keeps every print area inside its
 * page limits (0 = automatic in Excel's mapping). Without fit-to-page the
 * authored `scale` applies, clamped to 10%-400%.
 */
function resolveScale(input: ScaleInput): number {
  if (!input.fitToPage) {
    const percent = input.setup?.scale ?? 100;
    return clamp(percent / 100, MIN_PRINT_SCALE, MAX_PRINT_SCALE);
  }
  const widthLimit = Math.max(0, input.setup?.fitToWidth ?? 1);
  const heightLimit = Math.max(0, input.setup?.fitToHeight ?? 0);
  if (widthLimit === 0 && heightLimit === 0) return 1;

  const fits = (scale: number): boolean => {
    const columnCapacity = contentCapacity(input.body.width, input.headingWidth, scale);
    const rowCapacity = contentCapacity(input.body.height, input.headingHeight, scale);
    for (const [index, columns] of input.columns.entries()) {
      const rows = input.rows[index] ?? [];
      if (widthLimit > 0 && countBands(columns, columnCapacity) > widthLimit) return false;
      if (heightLimit > 0 && countBands(rows, rowCapacity) > heightLimit) return false;
    }
    return true;
  };

  if (fits(1)) return 1;
  if (!fits(MIN_PRINT_SCALE)) return MIN_PRINT_SCALE;
  let low = MIN_PRINT_SCALE;
  let high = 1;
  for (let step = 0; step < 40; step += 1) {
    const middle = (low + high) / 2;
    if (fits(middle)) low = middle;
    else high = middle;
  }
  // Excel offers whole percentages; round down so the fit can never break.
  return Math.max(MIN_PRINT_SCALE, Math.floor(low * 100) / 100);
}

function bandOf(
  range: GridRange,
  columnBand: readonly AxisEntry[],
  rowBand: readonly AxisEntry[],
): PrintBand {
  return {
    startRow: rowBand[0]?.index ?? range.startRow,
    endRow: rowBand[rowBand.length - 1]?.index ?? range.endRow,
    startCol: columnBand[0]?.index ?? range.startCol,
    endCol: columnBand[columnBand.length - 1]?.index ?? range.endCol,
  };
}

// ---------------------------------------------------------------------------
// One page
// ---------------------------------------------------------------------------

interface PageInput {
  readonly sheet: PrintSheetSnapshot;
  readonly measure: PrintTextMeasure;
  readonly paper: PrintPaper;
  readonly box: PrintBox;
  readonly body: PrintBox;
  readonly scale: number;
  readonly index: number;
  readonly setIndex: number;
  readonly pageInSetIndex: number;
  readonly band: PrintBand;
  readonly columnBand: readonly AxisEntry[];
  readonly rowBand: readonly AxisEntry[];
  readonly showGridLines: boolean;
  readonly showHeadings: boolean;
  readonly horizontalCentered: boolean;
  readonly verticalCentered: boolean;
  readonly headingWidth: number;
  readonly headingHeight: number;
  readonly headingFont: PrintSheetSnapshot['headingFont'];
}

interface ColumnPosition {
  readonly col: number;
  /** The unscaled offset from the content block's left edge. */
  readonly offset: number;
  readonly size: number;
}

interface RowPosition {
  readonly row: number;
  readonly offset: number;
  readonly size: number;
}

function buildPage(input: PageInput): PrintPage {
  const { scale, body, columnBand, rowBand } = input;
  const columns: ColumnPosition[] = [];
  let localX = 0;
  for (const entry of columnBand) {
    columns.push({ col: entry.index, offset: localX, size: entry.size });
    localX += entry.size;
  }
  const rows: RowPosition[] = [];
  let localY = 0;
  for (const entry of rowBand) {
    rows.push({ row: entry.index, offset: localY, size: entry.size });
    localY += entry.size;
  }
  const contentWidth = localX * scale;
  const contentHeight = localY * scale;
  const headingWidth = input.headingWidth * scale;
  const headingHeight = input.headingHeight * scale;
  const blockWidth = contentWidth + headingWidth;
  const blockHeight = contentHeight + headingHeight;
  const offsetX = input.horizontalCentered ? Math.max(0, (body.width - blockWidth) / 2) : 0;
  const offsetY = input.verticalCentered ? Math.max(0, (body.height - blockHeight) / 2) : 0;
  const originX = body.x + offsetX + headingWidth;
  const originY = body.y + offsetY + headingHeight;
  const geometry: PageGeometry = { originX, originY, scale, body };

  const pageColumns: PrintPageColumn[] = columns.map((column) => {
    const x = round(originX + column.offset * scale);
    return {
      col: column.col,
      x,
      width: round(originX + (column.offset + column.size) * scale) - x,
    };
  });
  const pageRows: PrintPageRow[] = rows.map((row) => {
    const y = round(originY + row.offset * scale);
    return {
      row: row.row,
      y,
      height: round(originY + (row.offset + row.size) * scale) - y,
    };
  });

  const columnByIndex = new Map<number, ColumnPosition>();
  for (const column of columns) columnByIndex.set(column.col, column);
  const rowByIndex = new Map<number, RowPosition>();
  for (const row of rows) rowByIndex.set(row.row, row);

  const gridlines: PrintItem[] = [];
  const fills: PrintFillItem[] = [];
  const headings: PrintItem[] = [];
  const texts: PrintItem[] = [];
  const borders: PrintItem[] = [];
  const band = input.band;
  const merges = input.sheet.mergesIn(band.startRow, band.endRow, band.startCol, band.endCol);

  if (input.showGridLines) {
    emitGridlines(columns, rows, columnByIndex, rowByIndex, merges, geometry, gridlines);
  }

  if (input.showHeadings) {
    emitHeadings(input, pageColumns, pageRows, geometry, headings);
  }

  const cells = input.sheet.cellsIn(band.startRow, band.endRow, band.startCol, band.endCol);
  for (const cell of cells) {
    if (cell.merge !== null) continue;
    const column = columnByIndex.get(cell.col);
    const row = rowByIndex.get(cell.row);
    if (column === undefined || row === undefined) continue;
    const local: PrintBox = {
      x: column.offset,
      y: row.offset,
      width: column.size,
      height: row.size,
    };
    const pageBox = mapBox(local, geometry);
    const clip = intersectBox(pageBox, body);
    if (clip !== null) {
      const fill = fillItemFor(cell.style, clip);
      if (fill !== null) fills.push(fill);
    }
    emitCellText(cell, local, geometry, input.measure, texts, clip ?? pageBox);
    emitCellBorder(
      cell.style.border,
      local,
      geometry,
      input.sheet,
      cell.row,
      cell.col,
      band,
      borders,
    );
  }

  for (const merge of merges) {
    paintMerge(input, merge, columnByIndex, rowByIndex, geometry, body, fills, texts, borders);
  }

  return {
    index: input.index,
    setIndex: input.setIndex,
    pageInSetIndex: input.pageInSetIndex,
    paper: input.paper,
    box: input.box,
    body,
    band,
    columns: pageColumns,
    rows: pageRows,
    items: [...gridlines, ...fills, ...headings, ...texts, ...borders],
    clipped:
      columnBand.some(
        (entry) => entry.size > contentCapacity(body.width, input.headingWidth, scale) + 1e-6,
      ) ||
      rowBand.some(
        (entry) => entry.size > contentCapacity(body.height, input.headingHeight, scale) + 1e-6,
      ),
  };
}

/** Local (unscaled, content-origin) box → page points. */
function mapBox(local: PrintBox, geometry: PageGeometry): PrintBox {
  const x = round(geometry.originX + local.x * geometry.scale);
  const y = round(geometry.originY + local.y * geometry.scale);
  return {
    x,
    y,
    width: round(geometry.originX + (local.x + local.width) * geometry.scale) - x,
    height: round(geometry.originY + (local.y + local.height) * geometry.scale) - y,
  };
}

function intersectBox(left: PrintBox, right: PrintBox): PrintBox | null {
  const x = Math.max(left.x, right.x);
  const y = Math.max(left.y, right.y);
  const right1 = Math.min(left.x + left.width, right.x + right.width);
  const bottom = Math.min(left.y + left.height, right.y + right.height);
  if (right1 <= x + 1e-6 || bottom <= y + 1e-6) return null;
  return { x: round(x), y: round(y), width: round(right1 - x), height: round(bottom - y) };
}

// ---------------------------------------------------------------------------
// Gridlines and headings
// ---------------------------------------------------------------------------

function emitGridlines(
  columns: readonly ColumnPosition[],
  rows: readonly RowPosition[],
  columnByIndex: ReadonlyMap<number, ColumnPosition>,
  rowByIndex: ReadonlyMap<number, RowPosition>,
  merges: readonly MergeSpan[],
  geometry: PageGeometry,
  out: PrintItem[],
): void {
  if (columns.length === 0 || rows.length === 0) return;
  const width = Math.max(0.25, GRIDLINE_WIDTH * geometry.scale);
  const firstRow = rows[0]!;
  const lastRow = rows[rows.length - 1]!;
  const firstColumn = columns[0]!;
  const lastColumn = columns[columns.length - 1]!;
  const top = geometry.originY + firstRow.offset * geometry.scale;
  const bottom = geometry.originY + (lastRow.offset + lastRow.size) * geometry.scale;
  const left = geometry.originX + firstColumn.offset * geometry.scale;
  const right = geometry.originX + (lastColumn.offset + lastColumn.size) * geometry.scale;
  const spacing = 0.01;

  // Vertical lines at every column boundary, minus the interior of merged
  // ranges (Excel does not print gridlines through a merge).
  const verticals = new Set<number>();
  for (const column of columns) {
    verticals.add(round(geometry.originX + column.offset * geometry.scale));
    verticals.add(round(geometry.originX + (column.offset + column.size) * geometry.scale));
  }
  for (const x of verticals) {
    const blockers = mergeBlocks(merges, columnByIndex, rowByIndex, geometry, 'x', x);
    for (const segment of subtractIntervals(top + spacing, bottom - spacing, blockers)) {
      out.push({
        kind: 'line',
        x1: x,
        y1: segment.start,
        x2: x,
        y2: segment.end,
        color: GRIDLINE_COLOR,
        width,
        dash: null,
      });
    }
  }

  const horizontals = new Set<number>();
  for (const row of rows) {
    horizontals.add(round(geometry.originY + row.offset * geometry.scale));
    horizontals.add(round(geometry.originY + (row.offset + row.size) * geometry.scale));
  }
  for (const y of horizontals) {
    const blockers = mergeBlocks(merges, columnByIndex, rowByIndex, geometry, 'y', y);
    for (const segment of subtractIntervals(left + spacing, right - spacing, blockers)) {
      out.push({
        kind: 'line',
        x1: segment.start,
        y1: y,
        x2: segment.end,
        y2: y,
        color: GRIDLINE_COLOR,
        width,
        dash: null,
      });
    }
  }
}

function round(value: number): number {
  return Math.round(value * 100) / 100;
}

/**
 * The y-intervals (for a vertical line at `x`) or x-intervals (for a
 * horizontal line at `y`) that a merged range covers.
 */
function mergeBlocks(
  merges: readonly MergeSpan[],
  columns: ReadonlyMap<number, ColumnPosition>,
  rows: ReadonlyMap<number, RowPosition>,
  geometry: PageGeometry,
  axis: 'x' | 'y',
  position: number,
): { start: number; end: number }[] {
  const blocks: { start: number; end: number }[] = [];
  for (const merge of merges) {
    const box = mergePageBox(merge, columns, rows, geometry);
    if (box === null) continue;
    if (axis === 'x') {
      if (position > box.x + 0.01 && position < box.x + box.width - 0.01) {
        blocks.push({ start: box.y, end: box.y + box.height });
      }
    } else if (position > box.y + 0.01 && position < box.y + box.height - 0.01) {
      blocks.push({ start: box.x, end: box.x + box.width });
    }
  }
  return blocks;
}

/** A merge's page box clipped to the band, or `null` when it misses the page. */
function mergePageBox(
  merge: MergeSpan,
  columns: ReadonlyMap<number, ColumnPosition>,
  rows: ReadonlyMap<number, RowPosition>,
  geometry: PageGeometry,
): PrintBox | null {
  const firstVisibleCol = columns.get(merge.anchorCol);
  const lastVisibleCol = columns.get(merge.anchorCol + merge.colSpan - 1);
  const firstVisibleRow = rows.get(merge.anchorRow);
  const lastVisibleRow = rows.get(merge.anchorRow + merge.rowSpan - 1);
  // A merge that starts or ends outside the band has no entry at its true
  // edge; the gridline pass only needs the part that is on this page, so use
  // the nearest visible edge instead.
  const left = firstVisibleCol ?? nearestVisible(columns, merge.anchorCol);
  const right = lastVisibleCol ?? nearestVisible(columns, merge.anchorCol + merge.colSpan - 1);
  const top = firstVisibleRow ?? nearestVisible(rows, merge.anchorRow);
  const bottom = lastVisibleRow ?? nearestVisible(rows, merge.anchorRow + merge.rowSpan - 1);
  if (left === null || right === null || top === null || bottom === null) return null;
  return {
    x: geometry.originX + left.offset * geometry.scale,
    y: geometry.originY + top.offset * geometry.scale,
    width: Math.max(0, (right.offset + right.size - left.offset) * geometry.scale),
    height: Math.max(0, (bottom.offset + bottom.size - top.offset) * geometry.scale),
  };
}

/**
 * The band position at `index`, or the nearest visible one: a merge that
 * starts or ends outside the band is clipped to the band's first or last
 * visible row/column.
 */
function nearestVisible<T>(positions: ReadonlyMap<number, T>, index: number): T | null {
  let best: T | null = null;
  let first: T | null = null;
  for (const [candidate, position] of positions) {
    if (first === null) first = position;
    if (candidate <= index) best = position;
    else break;
  }
  return best ?? first;
}

/** Subtract the blocker intervals from [start, end], keeping the gaps. */
function subtractIntervals(
  start: number,
  end: number,
  blockers: readonly { start: number; end: number }[],
): { start: number; end: number }[] {
  let segments: { start: number; end: number }[] = [{ start, end }];
  for (const blocker of blockers) {
    const next: { start: number; end: number }[] = [];
    for (const segment of segments) {
      if (blocker.end <= segment.start || blocker.start >= segment.end) {
        next.push(segment);
        continue;
      }
      if (blocker.start > segment.start) next.push({ start: segment.start, end: blocker.start });
      if (blocker.end < segment.end) next.push({ start: blocker.end, end: segment.end });
    }
    segments = next;
  }
  return segments.filter((segment) => segment.end - segment.start > 0.01);
}

function emitHeadings(
  input: PageInput,
  columns: readonly PrintPageColumn[],
  rows: readonly PrintPageRow[],
  geometry: PageGeometry,
  out: PrintItem[],
): void {
  const font = headingFont(input.headingFont, geometry.scale);
  const clip = input.body;
  const letterTop =
    (rows.length > 0 ? rows[0]!.y : geometry.originY) - input.headingHeight * geometry.scale;
  for (const column of columns) {
    const text = input.sheet.columnLabel(column.col);
    if (text === '') continue;
    out.push(
      textItem(
        text,
        column.x + column.width / 2,
        letterTop + font.size + 2,
        'middle',
        font,
        0,
        clip,
      ),
    );
  }
  for (const row of rows) {
    const text = input.sheet.rowLabel(row.row);
    if (text === '') continue;
    const baseline = row.y + row.height / 2 + font.size * 0.35;
    out.push(textItem(text, geometry.originX - 3, baseline, 'end', font, 0, clip));
  }
}

function headingFont(font: PrintSheetSnapshot['headingFont'], scale: number): PrintFont {
  return {
    family: font.name,
    size: Math.max(MIN_FONT_POINTS, Math.max(1, font.size) * scale),
    bold: false,
    italic: false,
    color: AUTOMATIC_BORDER_COLOR,
  };
}

// ---------------------------------------------------------------------------
// Fills, text and borders
// ---------------------------------------------------------------------------

function fillItemFor(style: ProjectedStyle, clip: PrintBox): PrintFillItem | null {
  const fill = style.fill;
  if (fill.pattern === 'none') return null;
  if (fill.pattern === '' || fill.pattern === 'solid') {
    if (fill.color === null) return null;
    return { kind: 'fill', box: clip, pattern: 'solid', color: fill.color, background: null };
  }
  const background = fill.backgroundColor ?? null;
  if (fill.color === null && background === null) return null;
  return { kind: 'fill', box: clip, pattern: fill.pattern, color: fill.color, background };
}

function emitCellText(
  cell: PrintCell,
  local: PrintBox,
  geometry: PageGeometry,
  measure: PrintTextMeasure,
  out: PrintItem[],
  clip: PrintBox,
): void {
  const style = cell.style;
  const rotation = rotationFor(style);
  const horizontal = horizontalFor(style, cell.value);
  const insets = textInsets(style, horizontal, 1);
  const size = Math.max(MIN_FONT_POINTS, Math.max(1, style.font.size));
  const link = cell.hyperlink;
  const color = link ? LINK_COLOR : textColorFor(style);
  const maximum = Math.max(
    1,
    local.width - insets.left * PT_PER_PIXEL - insets.right * PT_PER_PIXEL,
  );

  if (rotation.kind === 'stacked') {
    paintStacked(cell, local, geometry, measure, out, clip, color);
    return;
  }

  if (rotation.kind === 'angle') {
    const available = Math.max(1, local.height - TEXT_PADDING * 2);
    const laidOut = cell.layoutText(available, measure);
    const text = laidOut.text;
    if (text === '') return;
    let fontSize = size;
    if (style.alignment.shrinkToFit) {
      fontSize = fitFontSize(text, available, style, size, measure);
    }
    const font = printFont(style, fontSize, geometry.scale, color);
    const width = measure(text, printFont(style, fontSize, 1, color));
    const degrees = (rotation.radians * 180) / Math.PI;
    const radians = rotation.radians;
    const centreX = local.x + local.width / 2;
    const centreY = local.y + local.height / 2;
    const shift = fontSize * 0.35;
    const anchorX = centreX - Math.sin(radians) * shift;
    const anchorY = centreY + Math.cos(radians) * shift;
    const point = mapPoint(anchorX, anchorY, geometry);
    out.push(textItem(text, point.x, point.y, 'middle', font, degrees, clip));
    if (!link && style.font.underlineStyle === 'none' && !style.font.strike) return;
    const half = width / 2;
    const direction = { x: Math.cos(radians), y: Math.sin(radians) };
    const normal = { x: -Math.sin(radians), y: Math.cos(radians) };
    const start = { x: anchorX - direction.x * half, y: anchorY - direction.y * half };
    emitDecorations(style, link, start, direction, normal, width, fontSize, geometry, out);
    return;
  }

  const laidOut = cell.layoutText(maximum, measure);
  if (laidOut.text === '') return;
  let fontSize = size;
  if (style.alignment.shrinkToFit && !style.alignment.wrapText) {
    fontSize = fitFontSize(laidOut.text, maximum, style, size, measure);
  }
  const fontForMeasure: PrintFont = printFont(style, fontSize, 1, color);
  const text = laidOut.text;
  const lines = style.alignment.wrapText
    ? wrapTextLines(text, maximum, (value) => measure(value, fontForMeasure))
    : [text];
  const lineHeight = fontSize * LINE_SPACING;
  const capacity = Math.max(1, Math.floor(local.height / lineHeight));
  const visible = lines.slice(0, capacity);
  if (visible.length === 0) return;

  const blockHeight = Math.max(lineHeight, visible.length * lineHeight);
  const vertical = verticalFor(style);
  let top: number;
  if (vertical === 'top') top = local.y + TOP_PADDING;
  else if (vertical === 'center') top = local.y + (local.height - blockHeight) / 2;
  else top = local.y + local.height - blockHeight - BOTTOM_PADDING;
  if (
    (vertical === 'bottom' || vertical === 'center') &&
    top + blockHeight > local.y + local.height
  ) {
    top = local.y + TOP_PADDING;
  }

  const left = local.x + insets.left * PT_PER_PIXEL;
  const right = local.x + local.width - insets.right * PT_PER_PIXEL;
  const x = horizontal === 'center' ? (left + right) / 2 : horizontal === 'right' ? right : left;
  const anchor: PrintTextAnchor =
    horizontal === 'center' ? 'middle' : horizontal === 'right' ? 'end' : 'start';
  const font = printFont(style, fontSize, geometry.scale, color);
  for (const [index, line] of visible.entries()) {
    if (line === '') continue;
    const baseline = top + index * lineHeight + fontSize;
    const point = mapPoint(x, baseline, geometry);
    out.push(textItem(line, point.x, point.y, anchor, font, 0, clip));
    emitDecorations(
      style,
      link,
      { x, y: baseline },
      { x: 1, y: 0 },
      { x: 0, y: 1 },
      measure(line, fontForMeasure),
      fontSize,
      geometry,
      out,
      horizontal,
      local,
      insets,
    );
  }
}

/**
 * Underline and strike as line items: the plan's text primitive carries no
 * decoration flags, so both renderers draw the same two lines.
 */
function emitDecorations(
  style: ProjectedStyle,
  link: boolean,
  origin: { x: number; y: number },
  direction: { x: number; y: number },
  normal: { x: number; y: number },
  width: number,
  size: number,
  geometry: PageGeometry,
  out: PrintItem[],
  horizontal?: 'left' | 'center' | 'right',
  local?: PrintBox,
  insets?: { left: number; right: number },
): void {
  const underline = link ? 'single' : style.font.underlineStyle;
  const accounting = underline === 'singleAccounting' || underline === 'doubleAccounting';
  const stroke = Math.max(0.5, (size / 14) * geometry.scale);
  const color = link ? LINK_COLOR : textColorFor(style);
  const start =
    accounting && horizontal !== undefined && local !== undefined && insets !== undefined
      ? {
          x: local.x + insets.left * PT_PER_PIXEL,
          y: origin.y,
        }
      : origin;
  const runWidth =
    accounting && local !== undefined && insets !== undefined
      ? Math.max(0, local.width - insets.left * PT_PER_PIXEL - insets.right * PT_PER_PIXEL)
      : width;
  if (underline !== 'none') {
    const offset = size * 0.12;
    const from = add(start, normal, offset);
    const to = add(from, direction, runWidth);
    const p1 = mapPoint(from.x, from.y, geometry);
    const p2 = mapPoint(to.x, to.y, geometry);
    out.push(lineItem(p1, p2, color, stroke));
    if (underline === 'double' || underline === 'doubleAccounting') {
      const secondFrom = add(start, normal, offset + size * 0.12);
      const secondTo = add(secondFrom, direction, runWidth);
      out.push(
        lineItem(
          mapPoint(secondFrom.x, secondFrom.y, geometry),
          mapPoint(secondTo.x, secondTo.y, geometry),
          color,
          stroke,
        ),
      );
    }
  }
  if (style.font.strike) {
    const from = add(origin, normal, -size * 0.3);
    const to = add(from, direction, width);
    out.push(
      lineItem(mapPoint(from.x, from.y, geometry), mapPoint(to.x, to.y, geometry), color, stroke),
    );
  }
}

function add(
  point: { x: number; y: number },
  direction: { x: number; y: number },
  amount: number,
): { x: number; y: number } {
  return { x: point.x + direction.x * amount, y: point.y + direction.y * amount };
}

function paintStacked(
  cell: PrintCell,
  local: PrintBox,
  geometry: PageGeometry,
  measure: PrintTextMeasure,
  out: PrintItem[],
  clip: PrintBox,
  color: string,
): void {
  const style = cell.style;
  const characters = Array.from(cell.displayText);
  if (characters.length === 0) return;
  const width = Math.max(1, local.width - TEXT_PADDING * 2);
  const height = Math.max(1, local.height - TEXT_PADDING * 2);
  let size = Math.max(MIN_FONT_POINTS, Math.max(1, style.font.size));
  let widest = 0;
  for (const character of characters) {
    widest = Math.max(widest, measure(character, printFont(style, size, 1, color)));
  }
  if (widest > width) size = Math.max(MIN_FONT_POINTS, size * (width / widest));
  const lineHeight = Math.min(size * LINE_SPACING, height / characters.length);
  const blockHeight = lineHeight * characters.length;
  const centreX = local.x + local.width / 2;
  const startY = local.y + (local.height - blockHeight) / 2 + lineHeight / 2;
  const font = printFont(style, size, geometry.scale, color);
  for (const [index, character] of characters.entries()) {
    const baseline = startY + index * lineHeight + size * 0.35;
    const point = mapPoint(centreX, baseline, geometry);
    out.push(textItem(character, point.x, point.y, 'middle', font, 0, clip));
  }
}

/** The largest font size at or below the style's that fits `maximum` points. */
function fitFontSize(
  text: string,
  maximum: number,
  style: ProjectedStyle,
  size: number,
  measure: PrintTextMeasure,
): number {
  const fits = (candidate: number): boolean =>
    measure(text, printFont(style, candidate, 1, textColorFor(style))) <= maximum;
  if (fits(size)) return size;
  if (!fits(MIN_FONT_POINTS)) return MIN_FONT_POINTS;
  let low = MIN_FONT_POINTS;
  let high = size;
  for (let step = 0; step < 24; step += 1) {
    const middle = (low + high) / 2;
    if (fits(middle)) low = middle;
    else high = middle;
  }
  return low;
}

function printFont(style: ProjectedStyle, size: number, scale: number, color: string): PrintFont {
  return {
    family: style.font.name,
    size: Math.max(MIN_FONT_POINTS, size * scale),
    bold: style.font.bold,
    italic: style.font.italic,
    color,
  };
}

function textColorFor(style: ProjectedStyle): string {
  return style.numberFormatColor ?? style.font.color ?? DEFAULT_TEXT_COLOR;
}

function textItem(
  text: string,
  x: number,
  y: number,
  anchor: PrintTextAnchor,
  font: PrintFont,
  rotation: number,
  clip: PrintBox,
): PrintTextItem {
  return { kind: 'text', text, x, y, anchor, font, rotation, clip };
}

function mapPoint(x: number, y: number, geometry: PageGeometry): { x: number; y: number } {
  return {
    x: round(geometry.originX + x * geometry.scale),
    y: round(geometry.originY + y * geometry.scale),
  };
}

function lineItem(
  from: { x: number; y: number },
  to: { x: number; y: number },
  color: string,
  width: number,
): PrintItem {
  return { kind: 'line', x1: from.x, y1: from.y, x2: to.x, y2: to.y, color, width, dash: null };
}

// ---------------------------------------------------------------------------
// Borders
// ---------------------------------------------------------------------------

const BORDER_STROKES: Readonly<
  Record<
    string,
    { readonly width: number; readonly dash: readonly number[] | null; readonly double: boolean }
  >
> = {
  hair: { width: 0.5, dash: null, double: false },
  thin: { width: 0.75, dash: null, double: false },
  medium: { width: 1.5, dash: null, double: false },
  thick: { width: 2.25, dash: null, double: false },
  dashed: { width: 0.75, dash: [3, 2], double: false },
  dotted: { width: 0.75, dash: [1, 1], double: false },
  dashDot: { width: 0.75, dash: [3, 1, 1, 1], double: false },
  dashDotDot: { width: 0.75, dash: [3, 1, 1, 1, 1, 1], double: false },
  double: { width: 0.75, dash: null, double: true },
  mediumDashed: { width: 1.5, dash: [4, 2], double: false },
  mediumDashDot: { width: 1.5, dash: [4, 1.5, 1.5, 1.5], double: false },
  mediumDashDotDot: { width: 1.5, dash: [4, 1.5, 1.5, 1.5, 1.5, 1.5], double: false },
  slantDashDot: { width: 1.5, dash: [4, 1.5, 1.5, 1.5], double: false },
};

type BorderEdge = 'left' | 'right' | 'top' | 'bottom';

function emitCellBorder(
  border: BorderStyle,
  local: PrintBox,
  geometry: PageGeometry,
  sheet: PrintSheetSnapshot,
  row: number,
  col: number,
  band: PrintBand,
  out: PrintItem[],
): void {
  emitBorderBox(border, local, geometry, sheet, row, col, row, col, band, out);
}

function emitBorderBox(
  border: BorderStyle,
  local: PrintBox,
  geometry: PageGeometry,
  sheet: PrintSheetSnapshot,
  startRow: number,
  startCol: number,
  endRow: number,
  endCol: number,
  band: PrintBand,
  out: PrintItem[],
): void {
  const page = mapBox(local, geometry);
  const edges: Record<BorderEdge, { readonly edge: BorderEdgeStyle; readonly draw: boolean }> = {
    left: {
      edge: border.left,
      draw: ownsEdge(
        border.left,
        startCol <= band.startCol
          ? null
          : (sheet.cellAt(startRow, startCol - 1)?.style.border.right ?? null),
        'stronger-or-equal',
      ),
    },
    right: {
      edge: border.right,
      draw: ownsEdge(
        border.right,
        endCol >= band.endCol
          ? null
          : (sheet.cellAt(startRow, endCol + 1)?.style.border.left ?? null),
        'stronger',
      ),
    },
    top: {
      edge: border.top,
      draw: ownsEdge(
        border.top,
        startRow <= band.startRow
          ? null
          : (sheet.cellAt(startRow - 1, startCol)?.style.border.bottom ?? null),
        'stronger-or-equal',
      ),
    },
    bottom: {
      edge: border.bottom,
      draw: ownsEdge(
        border.bottom,
        endRow >= band.endRow
          ? null
          : (sheet.cellAt(endRow + 1, startCol)?.style.border.top ?? null),
        'stronger',
      ),
    },
  };
  if (edges.left.draw) emitEdge('left', edges.left.edge, page, geometry, out);
  if (edges.right.draw) emitEdge('right', edges.right.edge, page, geometry, out);
  if (edges.top.draw) emitEdge('top', edges.top.edge, page, geometry, out);
  if (edges.bottom.draw) emitEdge('bottom', edges.bottom.edge, page, geometry, out);
  if (border.diagonalUp || border.diagonalDown) {
    const stroke = BORDER_STROKES[border.diagonal.style] ?? null;
    if (stroke !== null) {
      const color = border.diagonal.color ?? AUTOMATIC_BORDER_COLOR;
      const width = Math.max(0.25, stroke.width * geometry.scale);
      if (border.diagonalUp) {
        out.push({
          kind: 'line',
          x1: page.x,
          y1: page.y + page.height,
          x2: page.x + page.width,
          y2: page.y,
          color,
          width,
          dash: stroke.dash?.map((value) => value * geometry.scale) ?? null,
        });
      }
      if (border.diagonalDown) {
        out.push({
          kind: 'line',
          x1: page.x,
          y1: page.y,
          x2: page.x + page.width,
          y2: page.y + page.height,
          color,
          width,
          dash: stroke.dash?.map((value) => value * geometry.scale) ?? null,
        });
      }
    }
  }
}

function ownsEdge(
  edge: BorderEdgeStyle,
  neighbour: BorderEdgeStyle | null,
  tie: 'stronger' | 'stronger-or-equal',
): boolean {
  const rank = edgeRank(edge);
  if (rank < 0) return false;
  if (neighbour === null) return true;
  const neighbourRank = edgeRank(neighbour);
  return tie === 'stronger' ? rank > neighbourRank : rank >= neighbourRank;
}

function edgeRank(edge: BorderEdgeStyle): number {
  if (edge.style === 'none' || edge.style === '') return -1;
  switch (edge.weight) {
    case 'hair':
      return 0;
    case 'thin':
      return 1;
    case 'medium':
      return 2;
    case 'thick':
      return 3;
    default:
      break;
  }
  switch (edge.style) {
    case 'hair':
      return 0;
    case 'medium':
    case 'mediumDashed':
    case 'mediumDashDot':
    case 'mediumDashDotDot':
    case 'slantDashDot':
    case 'double':
      return 2;
    case 'thick':
      return 3;
    default:
      return 1;
  }
}

function emitEdge(
  edgeName: BorderEdge,
  edge: BorderEdgeStyle,
  page: PrintBox,
  geometry: PageGeometry,
  out: PrintItem[],
): void {
  const stroke = BORDER_STROKES[edge.style] ?? BORDER_STROKES.thin!;
  const color = edge.color ?? AUTOMATIC_BORDER_COLOR;
  const width = Math.max(0.25, stroke.width * geometry.scale);
  const dash = stroke.dash?.map((value) => value * geometry.scale) ?? null;
  const coordinates = (offset: number): { x1: number; y1: number; x2: number; y2: number } => {
    switch (edgeName) {
      case 'left':
        return { x1: page.x + offset, y1: page.y, x2: page.x + offset, y2: page.y + page.height };
      case 'right':
        return {
          x1: page.x + page.width + offset,
          y1: page.y,
          x2: page.x + page.width + offset,
          y2: page.y + page.height,
        };
      case 'top':
        return { x1: page.x, y1: page.y + offset, x2: page.x + page.width, y2: page.y + offset };
      case 'bottom':
        return {
          x1: page.x,
          y1: page.y + page.height + offset,
          x2: page.x + page.width,
          y2: page.y + page.height + offset,
        };
    }
  };
  if (stroke.double) {
    const offset = Math.max(0.4, 0.75 * geometry.scale) / 2;
    for (const direction of [-offset, offset]) {
      out.push({ kind: 'line', ...coordinates(direction), color, width, dash });
    }
    return;
  }
  out.push({ kind: 'line', ...coordinates(0), color, width, dash });
}

// ---------------------------------------------------------------------------
// Merged cells
// ---------------------------------------------------------------------------

/**
 * Paint one merged range on this page (print-pdf §4): the region is clipped
 * per page, its text only on the anchor's page, and its borders only on the
 * sides that are the merge's true outer edges.
 */
function paintMerge(
  input: PageInput,
  merge: MergeSpan,
  columns: ReadonlyMap<number, ColumnPosition>,
  rows: ReadonlyMap<number, RowPosition>,
  geometry: PageGeometry,
  body: PrintBox,
  fills: PrintFillItem[],
  texts: PrintItem[],
  borders: PrintItem[],
): void {
  const band = input.band;
  const lastRow = merge.anchorRow + merge.rowSpan - 1;
  const lastCol = merge.anchorCol + merge.colSpan - 1;
  const firstVisibleCol = columns.get(Math.max(merge.anchorCol, band.startCol));
  const lastVisibleCol = columns.get(Math.min(lastCol, band.endCol));
  const firstVisibleRow = rows.get(Math.max(merge.anchorRow, band.startRow));
  const lastVisibleRow = rows.get(Math.min(lastRow, band.endRow));
  if (
    firstVisibleCol === undefined ||
    lastVisibleCol === undefined ||
    firstVisibleRow === undefined ||
    lastVisibleRow === undefined
  ) {
    return;
  }
  const local: PrintBox = {
    x: firstVisibleCol.offset,
    y: firstVisibleRow.offset,
    width: lastVisibleCol.offset + lastVisibleCol.size - firstVisibleCol.offset,
    height: lastVisibleRow.offset + lastVisibleRow.size - firstVisibleRow.offset,
  };
  const anchor = input.sheet.cellAt(merge.anchorRow, merge.anchorCol);
  const style = anchor?.style ?? null;
  const page = mapBox(local, geometry);
  const clip = intersectBox(page, body);
  if (clip !== null && style !== null) {
    const fill = fillItemFor(style, clip);
    if (fill !== null) fills.push(fill);
  }
  const anchorOnPage =
    merge.anchorRow >= band.startRow &&
    merge.anchorRow <= band.endRow &&
    merge.anchorCol >= band.startCol &&
    merge.anchorCol <= band.endCol;
  if (anchorOnPage && anchor !== null) {
    emitCellText(anchor, local, geometry, input.measure, texts, clip ?? page);
  }
  const resolved = mergedBorder(merge, input.sheet);
  const trueLeft = merge.anchorCol >= band.startCol;
  const trueTop = merge.anchorRow >= band.startRow;
  const trueRight = lastCol <= band.endCol;
  const trueBottom = lastRow <= band.endRow;
  const box: BorderStyle = {
    left: trueLeft ? resolved.left : noneBorderEdge(),
    right: trueRight ? resolved.right : noneBorderEdge(),
    top: trueTop ? resolved.top : noneBorderEdge(),
    bottom: trueBottom ? resolved.bottom : noneBorderEdge(),
    diagonal: noneBorderEdge(),
    diagonalUp: false,
    diagonalDown: false,
  };
  emitBorderBox(
    box,
    local,
    geometry,
    input.sheet,
    merge.anchorRow,
    merge.anchorCol,
    lastRow,
    lastCol,
    band,
    borders,
  );
}

/** Mirror of `src/grid/merged-borders.ts` over the print snapshot. */
function mergedBorder(merge: MergeSpan, sheet: PrintSheetSnapshot): BorderStyle {
  const anchor = sheet.cellAt(merge.anchorRow, merge.anchorCol)?.style.border ?? null;
  const lastRow = merge.anchorRow + merge.rowSpan - 1;
  const lastCol = merge.anchorCol + merge.colSpan - 1;
  const pick = (
    cells: { row: number; col: number }[],
    edge: BorderEdge,
    fallback: BorderEdgeStyle,
  ): BorderEdgeStyle => {
    for (const { row, col } of cells) {
      const candidate = sheet.cellAt(row, col)?.style.border[edge];
      if (candidate !== undefined && candidate.style !== 'none' && candidate.style !== '') {
        return candidate;
      }
    }
    return fallback;
  };
  const leftCells = [];
  const rightCells = [];
  for (let row = merge.anchorRow; row <= lastRow; row += 1) {
    leftCells.push({ row, col: merge.anchorCol });
    rightCells.push({ row, col: lastCol });
  }
  const topCells = [];
  const bottomCells = [];
  for (let col = merge.anchorCol; col <= lastCol; col += 1) {
    topCells.push({ row: merge.anchorRow, col });
    bottomCells.push({ row: lastRow, col });
  }
  return {
    left: pick(leftCells, 'left', anchor?.left ?? noneBorderEdge()),
    right: pick(rightCells, 'right', anchor?.right ?? noneBorderEdge()),
    top: pick(topCells, 'top', anchor?.top ?? noneBorderEdge()),
    bottom: pick(bottomCells, 'bottom', anchor?.bottom ?? noneBorderEdge()),
    diagonal: anchor?.diagonal ?? noneBorderEdge(),
    diagonalUp: anchor?.diagonalUp ?? false,
    diagonalDown: anchor?.diagonalDown ?? false,
  };
}

function noneBorderEdge(): BorderEdgeStyle {
  return { style: 'none', color: null, weight: null };
}

function clamp(value: number, low: number, high: number): number {
  if (!Number.isFinite(value)) return 1;
  return Math.min(high, Math.max(low, value));
}
