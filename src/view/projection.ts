import type { AutoFilter, Cell, Sheet, Workbook } from '@model/types';
import { cellKey, columnToLetters } from '@model/types';
import { computeHiddenLayers } from '@model/autofilter';
import type { CfCellSource } from '@model/conditional-format';
import { BLANK, formatCellValue } from '@model/value';
import { DEFAULT_ALIGNMENT, DEFAULT_BORDER, DEFAULT_FILL, DEFAULT_FONT } from '@model/style';
import { SparseAxisGeometry, columnWidthCharsToPixels, pointsToPixels } from './geometry';
import {
  createNumberFormatter,
  formatForLayout,
  numberFormatColor,
  numberFormatColorForText,
  numberFormatColorForValue,
} from './number-format';
import { buildSheetIndex, type SheetIndex } from './sheet-index';
import { createSheetConditionalFormatting } from './cf-projection';
import type {
  MergeSpan,
  ProjectedAutoFilter,
  ProjectedCell,
  ProjectedConditionalFormatting,
  ProjectedStyle,
  ProjectedTextLayout,
  ProjectionOptions,
  SelectionStatistics,
  SheetProjection,
} from './types';

const DEFAULT_ROW_HEIGHT = 20;
const STYLES_PART_PATH = 'xl/styles.xml';

/**
 * The rows a sheet hides at one document version: the stored flags plus the
 * rows the active AutoFilter derives (sort-autofilter §3.3).
 */
const effectiveHiddenCache = new WeakMap<Sheet, { version: number; rows: Set<number> }>();

function effectiveHiddenRows(
  sheet: Sheet,
  manualHidden: readonly number[],
  version: number,
): Set<number> {
  const cached = effectiveHiddenCache.get(sheet);
  if (cached !== undefined && cached.version === version) return cached.rows;
  const rows = new Set(computeHiddenLayers(sheet, sheet.autoFilter, manualHidden).effectiveHidden);
  effectiveHiddenCache.set(sheet, { version, rows });
  return rows;
}
const DEFAULT_COLUMN_WIDTH = 64;
const MAX_ROWS = 1_048_576;
const MAX_COLUMNS = 16_384;
const FORMAT_CACHE_LIMIT = 4096;
const MIN_ZOOM = 0.1;
const MAX_ZOOM = 4;

/**
 * The sheet index is built once per document version. It is keyed by the
 * sheet's cell map, which survives edits, so a scroll never rebuilds it.
 */
const indexCache = new WeakMap<ReadonlyMap<string, Cell>, { version: number; index: SheetIndex }>();

export function createSheetProjection(
  workbook: Workbook,
  sheetId: string,
  options: ProjectionOptions,
): SheetProjection {
  const sheet = workbook.sheets.find((candidate) => candidate.id === sheetId);
  if (sheet === undefined) throw new Error(`Unknown sheet: ${sheetId}`);
  // The saved zoom is document state; the live zoom override is view state.
  // Both are ratios here, while the workbook stores an Excel percentage.
  const zoom = clampZoom(
    options.zoom === undefined ? sheet.view.zoom / 100 : normalizeZoom(options.zoom),
  );
  const defaultRowHeight =
    options.defaultRowHeight ??
    (sheet.defaultRowHeight === null ? DEFAULT_ROW_HEIGHT : pointsToPixels(sheet.defaultRowHeight));
  const defaultColumnWidth =
    options.defaultColumnWidth ??
    (sheet.defaultColumnWidth === null
      ? DEFAULT_COLUMN_WIDTH
      : columnWidthCharsToPixels(sheet.defaultColumnWidth));
  const rowCount = options.rowCount ?? MAX_ROWS;
  const columnCount = options.columnCount ?? MAX_COLUMNS;

  // Filter-hidden rows are derived from the criteria and never stored
  // (sort-autofilter §3.3), so the projection folds them into the hidden set
  // alongside the stored flags. The evaluation is cached per sheet and document
  // version, like the CF rules: a zoom or other view-only rebuild reuses it.
  const dimensionHiddenRows = [...sheet.rowDimensions.entries()]
    .filter(([row, dimension]) => row >= 0 && row < rowCount && dimension.hidden === true)
    .map(([row]) => row);
  const hiddenRows = effectiveHiddenRows(sheet, dimensionHiddenRows, options.version);
  const hiddenColumns = new Set(
    [...sheet.columnDimensions.entries()]
      .filter(([col, dimension]) => col >= 0 && col < columnCount && dimension.hidden === true)
      .map(([col]) => col),
  );
  const rowOverrides = rowHeightOverrides(sheet.rowDimensions, defaultRowHeight);
  // Filter-hidden rows are derived from the criteria and never stored
  // (sort-autofilter §3.3), so `rowHeightOverrides` cannot see them. Collapse
  // every effective-hidden row to zero height, exactly like a stored hidden
  // row: Excel closes the gap, it does not paint an empty row of normal height.
  for (const row of hiddenRows) rowOverrides.set(row, 0);
  const columnOverrides = columnWidthOverrides(sheet.columnDimensions, defaultColumnWidth);
  // A drag resize is view state until the pointer is released. Overriding one
  // sparse entry keeps the painted geometry live without touching the model.
  const preview = options.dimensionPreview ?? null;
  if (preview !== null && Number.isFinite(preview.size) && preview.size >= 0) {
    // A header drag previews the same size on every selected visible member.
    for (const index of preview.indices) {
      if (preview.axis === 'row') {
        if (index >= 0 && index < rowCount && !hiddenRows.has(index)) {
          rowOverrides.set(index, pointsToPixels(preview.size));
        }
      } else if (index >= 0 && index < columnCount && !hiddenColumns.has(index)) {
        columnOverrides.set(index, columnWidthCharsToPixels(preview.size));
      }
    }
  }
  const rowGeometry = new SparseAxisGeometry(
    rowCount,
    defaultRowHeight * zoom,
    scaleOverrides(rowOverrides, zoom),
  );
  const columnGeometry = new SparseAxisGeometry(
    columnCount,
    defaultColumnWidth * zoom,
    scaleOverrides(columnOverrides, zoom),
  );

  const merges = buildMergeIndex(sheet.merges);
  const formatter = createNumberFormatter();
  const styleCache = new Map<number, ProjectedStyle>();
  const valueStyleCache = new Map<string, ProjectedStyle>();
  const formatCache = new Map<string, string>();
  const colorCache = new Map<string, string | null>();

  const styleFor = (styleId: number): ProjectedStyle => {
    const cached = styleCache.get(styleId);
    if (cached !== undefined) return cached;
    const record = workbook.styles.find((candidate) => candidate.id === styleId);
    const style: ProjectedStyle =
      record === undefined
        ? {
            font: DEFAULT_FONT,
            fill: DEFAULT_FILL,
            border: DEFAULT_BORDER,
            alignment: DEFAULT_ALIGNMENT,
            numberFormat: null,
            numberFormatColor: null,
          }
        : {
            font: record.font,
            fill: record.fill,
            border: record.border,
            alignment: record.alignment,
            numberFormat: record.numFmtCode,
            numberFormatColor: numberFormatColor(record.numFmtCode, workbook.indexedPalette),
          };
    styleCache.set(styleId, style);
    return style;
  };

  /**
   * A number format may colour its sections differently (`[Red]0` negative,
   * `[Blue]0` positive, `[Red]@` text). Resolve the colour for this value's
   * section and reuse the immutable style object when it matches the
   * style-level colour.
   */
  const styleForValue = (cell: Cell, style: ProjectedStyle): ProjectedStyle => {
    if (style.numberFormat === null) return style;
    let colorKey: string;
    let color: string | null;
    if (cell.value.type === 'number') {
      colorKey = `n\u0000${style.numberFormat}\u0000${cell.value.value}`;
      const cached = colorCache.get(colorKey);
      if (cached !== undefined) color = cached;
      else {
        color = numberFormatColorForValue(
          style.numberFormat,
          cell.value.value,
          workbook.indexedPalette,
        );
        if (colorCache.size >= FORMAT_CACHE_LIMIT) colorCache.clear();
        colorCache.set(colorKey, color);
      }
    } else if (cell.value.type === 'text') {
      colorKey = `t\u0000${style.numberFormat}`;
      const cached = colorCache.get(colorKey);
      if (cached !== undefined) color = cached;
      else {
        color = numberFormatColorForText(style.numberFormat, workbook.indexedPalette);
        if (colorCache.size >= FORMAT_CACHE_LIMIT) colorCache.clear();
        colorCache.set(colorKey, color);
      }
    } else {
      return style;
    }
    if (color === style.numberFormatColor) return style;
    const key = `${cell.styleId}\u0000${color ?? ''}`;
    const cached = valueStyleCache.get(key);
    if (cached !== undefined) return cached;
    const resolved: ProjectedStyle = { ...style, numberFormatColor: color };
    if (valueStyleCache.size >= FORMAT_CACHE_LIMIT) valueStyleCache.clear();
    valueStyleCache.set(key, resolved);
    return resolved;
  };

  /** Width-agnostic display text; the paint width seam is `layoutTextFor`. */
  const displayTextFor = (cell: Cell, style: ProjectedStyle): string => {
    const key = `${style.numberFormat ?? ''}\u0000${valueKey(cell)}`;
    const cached = formatCache.get(key);
    if (cached !== undefined) return cached;
    const text = formatter.format(cell.value, style.numberFormat, workbook.date1904);
    if (formatCache.size >= FORMAT_CACHE_LIMIT) formatCache.clear();
    formatCache.set(key, text);
    return text;
  };

  const rawTextFor = (cell: Cell): string => {
    // The formula bar and the editor show the authored formula, with the
    // leading equals sign Excel uses, or the raw value of a plain cell.
    if (cell.formula !== null) return `=${cell.formula}`;
    return formatCellValue(cell.value);
  };

  const layoutTextFor = (
    cell: Cell,
    style: ProjectedStyle,
    widthPx: number,
    measure: (text: string) => number,
  ): ProjectedTextLayout => {
    if (cell.value.type !== 'number') {
      return { text: displayTextFor(cell, style), overflow: false };
    }
    const laidOut = formatForLayout(cell.value, style.numberFormat, {
      widthPx,
      measure,
      date1904: workbook.date1904,
      indexedPalette: workbook.indexedPalette,
      wrapText: style.alignment.wrapText,
      shrinkToFit: style.alignment.shrinkToFit,
    });
    return { text: laidOut.text, overflow: laidOut.overflow };
  };

  const indexFor = (): SheetIndex => {
    const cached = indexCache.get(sheet.cells);
    if (cached !== undefined && cached.version === options.version) return cached.index;
    const index = buildSheetIndex(sheet.cells);
    indexCache.set(sheet.cells, { version: options.version, index });
    return index;
  };

  const project = (row: number, col: number, cell: Cell): ProjectedCell => {
    const style = styleForValue(cell, styleFor(cell.styleId));
    return {
      row,
      col,
      value: cell.value,
      displayText: displayTextFor(cell, style),
      rawText: rawTextFor(cell),
      formula: cell.formula,
      styleId: cell.styleId,
      stale: cell.stale === true || (cell.formula !== null && cell.savedResult === null),
      style,
      merge: merges.get(cellKey(row, col)) ?? null,
      hyperlink: sheet.hyperlinks.get(cellKey(row, col)) ?? null,
      layoutText: (widthPx, measure) => layoutTextFor(cell, style, widthPx, measure),
    };
  };

  /**
   * Conditional formatting for painting (conditional formatting §5, §6). The
   * engine's cell source reads the same values and display text the projection
   * paints, so a text rule matches the text the user sees.
   */
  const cfCells: CfCellSource = {
    valueAt(row, col) {
      return sheet.cells.get(cellKey(row, col))?.value ?? BLANK;
    },
    displayTextAt(row, col) {
      const cell = sheet.cells.get(cellKey(row, col));
      if (cell === undefined) return '';
      return displayTextFor(cell, styleForValue(cell, styleFor(cell.styleId)));
    },
  };
  const conditionalFormatting: ProjectedConditionalFormatting | null =
    createSheetConditionalFormatting({
      sheet,
      version: options.version,
      date1904: workbook.date1904,
      indexedPalette: workbook.indexedPalette,
      sourceBytes: workbook.sourceParts.get(sheet.partPath) ?? null,
      stylesBytes: workbook.sourceParts.get(STYLES_PART_PATH) ?? null,
      cells: cfCells,
      format: (value, code) => formatter.format(value, code, workbook.date1904),
      inputs: options.conditionalFormats,
    });
  const autoFilter = projectAutoFilter(sheet.autoFilter, columnCount);

  return {
    sheetId,
    sheetName: sheet.name,
    version: options.version,
    rowCount,
    columnCount,
    freezeRows: Math.min(sheet.view.freezeRows, rowCount),
    freezeCols: Math.min(sheet.view.freezeCols, columnCount),
    showGridLines: options.showGridLines ?? sheet.view.showGridLines,
    showHeadings: options.showHeadings ?? sheet.view.showHeadings,
    zoom,
    date1904: workbook.date1904,
    autoFilter,
    conditionalFormatting,
    commentCells: new Set(sheet.comments.keys()),
    commentAt(row, col) {
      if (row < 0 || col < 0 || row >= rowCount || col >= columnCount) return null;
      return sheet.comments.get(cellKey(row, col)) ?? null;
    },
    hyperlinkAt(row, col) {
      if (row < 0 || col < 0 || row >= rowCount || col >= columnCount) return null;
      return sheet.hyperlinks.get(cellKey(row, col)) ?? null;
    },
    defaultRowHeightPoints: sheet.defaultRowHeight,
    defaultColumnWidthChars: sheet.defaultColumnWidth,
    cellAt(row, col) {
      if (row < 0 || col < 0 || row >= rowCount || col >= columnCount) return null;
      const cell = sheet.cells.get(cellKey(row, col));
      if (cell === undefined) return null;
      return project(row, col, cell);
    },
    window(startRow, endRow, startCol, endCol) {
      const cells: ProjectedCell[] = [];
      const firstRow = Math.max(0, Math.min(startRow, endRow));
      const lastRow = Math.min(rowCount - 1, Math.max(startRow, endRow));
      const firstCol = Math.max(0, Math.min(startCol, endCol));
      const lastCol = Math.min(columnCount - 1, Math.max(startCol, endCol));
      for (let row = firstRow; row <= lastRow; row += 1) {
        for (let col = firstCol; col <= lastCol; col += 1) {
          const cell = sheet.cells.get(cellKey(row, col));
          if (cell === undefined) continue;
          cells.push(project(row, col, cell));
        }
      }
      return cells;
    },
    rowHeight: (row) => rowGeometry.sizeAt(row),
    columnWidth: (col) => columnGeometry.sizeAt(col),
    isRowHidden: (row) => hiddenRows.has(row),
    isColumnHidden: (col) => hiddenColumns.has(col),
    rowDimension(row) {
      const dimension = sheet.rowDimensions.get(row);
      return dimension === undefined ? null : { ...dimension };
    },
    columnDimension(col) {
      const dimension = sheet.columnDimensions.get(col);
      return dimension === undefined ? null : { ...dimension };
    },
    columnLabel: columnToLetters,
    rowLabel: (row) => String(row + 1),
    rowY: (row) => rowGeometry.offsetOf(row),
    columnX: (col) => columnGeometry.offsetOf(col),
    rowAt: (y) => rowGeometry.indexAt(y),
    columnAt: (x) => columnGeometry.indexAt(x),
    totalHeight: () => rowGeometry.totalSize(),
    totalWidth: () => columnGeometry.totalSize(),
    mergeAt(row, col) {
      if (row < 0 || col < 0) return null;
      return merges.get(cellKey(row, col)) ?? null;
    },
    usedRange() {
      const index = indexFor();
      if (index.cellCount === 0) return null;
      return {
        startRow: index.minRow,
        endRow: index.maxRow,
        startCol: index.minCol,
        endCol: index.maxCol,
      };
    },
    rowDataColumns(row) {
      return indexFor().rows.get(row) ?? [];
    },
    columnDataRows(col) {
      return indexFor().columns.get(col) ?? [];
    },
    rawTextAt(row, col) {
      const cell = sheet.cells.get(cellKey(row, col));
      if (cell === undefined) return '';
      return rawTextFor(cell);
    },
    layoutTextAt(row, col, widthPx, measure) {
      if (row < 0 || col < 0 || row >= rowCount || col >= columnCount) return null;
      const cell = sheet.cells.get(cellKey(row, col));
      if (cell === undefined) return null;
      const style = styleForValue(cell, styleFor(cell.styleId));
      return layoutTextFor(cell, style, widthPx, measure);
    },
    selectionStatistics(range) {
      const startRow = Math.max(0, Math.min(range.startRow, range.endRow));
      const endRow = Math.min(rowCount - 1, Math.max(range.startRow, range.endRow));
      const startCol = Math.max(0, Math.min(range.startCol, range.endCol));
      const endCol = Math.min(columnCount - 1, Math.max(range.startCol, range.endCol));
      if (startRow > endRow || startCol > endCol) return EMPTY_STATISTICS;
      const index = indexFor();
      let count = 0;
      let sum = 0;
      const add = (row: number, col: number): void => {
        const cell = sheet.cells.get(cellKey(row, col));
        if (cell === undefined || cell.value.type !== 'number') return;
        count += 1;
        sum += cell.value.value;
      };
      // Walk the smaller axis and binary-search the sparse coordinate list on
      // the other, so a whole-column or whole-sheet selection touches only
      // cells that exist, not the addressable grid.
      if (endRow - startRow <= endCol - startCol) {
        for (let row = startRow; row <= endRow; row += 1) {
          const columns = index.rows.get(row);
          if (columns === undefined) continue;
          const bounds = sortedBounds(columns, startCol, endCol);
          for (let position = bounds.start; position < bounds.end; position += 1) {
            add(row, columns[position]!);
          }
        }
      } else {
        for (let col = startCol; col <= endCol; col += 1) {
          const rows = index.columns.get(col);
          if (rows === undefined) continue;
          const bounds = sortedBounds(rows, startRow, endRow);
          for (let position = bounds.start; position < bounds.end; position += 1) {
            add(rows[position]!, col);
          }
        }
      }
      return {
        count,
        sum,
        average: count === 0 ? 0 : sum / count,
      };
    },
  };
}

const EMPTY_STATISTICS: SelectionStatistics = Object.freeze({
  count: 0,
  sum: 0,
  average: 0,
});

function normalizeZoom(value: number): number {
  // Callers pass a ratio; a value above 400% is read as an Excel percentage
  // so a direct `zoom: 150` cannot be mistaken for a 15000% scale.
  return value > MAX_ZOOM ? value / 100 : value;
}

function clampZoom(value: number): number {
  if (!Number.isFinite(value)) return 1;
  return Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, value));
}

function scaleOverrides(overrides: ReadonlyMap<number, number>, zoom: number): Map<number, number> {
  if (zoom === 1) return overrides as Map<number, number>;
  const scaled = new Map<number, number>();
  for (const [index, size] of overrides) scaled.set(index, size * zoom);
  return scaled;
}

/** The half-open entries of a sorted index list inside an inclusive range. */
function sortedBounds(
  values: readonly number[],
  low: number,
  high: number,
): { readonly start: number; readonly end: number } {
  let start = 0;
  let end = values.length;
  while (start < end) {
    const middle = (start + end) >> 1;
    if ((values[middle] ?? 0) < low) start = middle + 1;
    else end = middle;
  }
  const from = start;
  end = values.length;
  while (start < end) {
    const middle = (start + end) >> 1;
    if ((values[middle] ?? 0) <= high) start = middle + 1;
    else end = middle;
  }
  return { start: from, end: start };
}

function rowHeightOverrides(
  dimensions: ReadonlyMap<number, { height?: number; hidden?: boolean }>,
  defaultHeight: number,
): Map<number, number> {
  const overrides = new Map<number, number>();
  for (const [row, dimension] of dimensions) {
    // A hidden row collapses to zero height, whatever height it records.
    if (dimension.hidden === true) {
      overrides.set(row, 0);
      continue;
    }
    if (dimension.height === undefined) continue;
    const pixels = pointsToPixels(dimension.height);
    if (pixels > 0 && pixels !== defaultHeight) overrides.set(row, pixels);
  }
  return overrides;
}

function columnWidthOverrides(
  dimensions: ReadonlyMap<number, { width?: number; hidden?: boolean }>,
  defaultWidth: number,
): Map<number, number> {
  const overrides = new Map<number, number>();
  for (const [column, dimension] of dimensions) {
    // A hidden column collapses to zero width, whatever width it records.
    if (dimension.hidden === true) {
      overrides.set(column, 0);
      continue;
    }
    if (dimension.width === undefined) continue;
    const pixels = columnWidthCharsToPixels(dimension.width);
    if (pixels > 0 && pixels !== defaultWidth) overrides.set(column, pixels);
  }
  return overrides;
}

/**
 * Reduce the sheet's AutoFilter to the paint-ready form. Columns outside the
 * addressable sheet are dropped; the criteria map is keyed by offset from the
 * range's first column, so adding that offset gives the header column.
 */
function projectAutoFilter(
  filter: AutoFilter | null,
  columnCount: number,
): ProjectedAutoFilter | null {
  if (filter === null) return null;
  const filteredColumns = new Set<number>();
  const unsupportedColumns = new Set<number>();
  for (const [offset, column] of filter.columns) {
    const col = filter.range.startCol + offset;
    if (col < 0 || col >= columnCount) continue;
    if (column.kind === 'unsupported') unsupportedColumns.add(col);
    else filteredColumns.add(col);
  }
  return {
    range: { ...filter.range },
    filteredColumns,
    unsupportedColumns,
  };
}

function valueKey(cell: Cell): string {
  const value = cell.value;
  switch (value.type) {
    case 'number':
      return `n${value.value}`;
    case 'text':
      return `t${value.value}`;
    case 'boolean':
      return `b${value.value ? 1 : 0}`;
    case 'error':
      return `e${value.value}`;
    case 'blank':
      return 'z';
  }
}

function buildMergeIndex(
  merges: readonly {
    readonly range: { startRow: number; startCol: number; endRow: number; endCol: number };
  }[],
): Map<string, MergeSpan> {
  const index = new Map<string, MergeSpan>();
  for (const merge of merges) {
    const { startRow, startCol, endRow, endCol } = merge.range;
    const span: MergeSpan = {
      anchorRow: startRow,
      anchorCol: startCol,
      rowSpan: endRow - startRow + 1,
      colSpan: endCol - startCol + 1,
    };
    for (let row = startRow; row <= endRow; row += 1) {
      for (let col = startCol; col <= endCol; col += 1) {
        index.set(cellKey(row, col), span);
      }
    }
  }
  return index;
}

export function visibleRange(
  projection: SheetProjection,
  scrollY: number,
  scrollX: number,
  width: number,
  height: number,
): { startRow: number; endRow: number; startCol: number; endCol: number } {
  return {
    startRow: projection.rowAt(scrollY),
    endRow: projection.rowAt(scrollY + height),
    startCol: projection.columnAt(scrollX),
    endCol: projection.columnAt(scrollX + width),
  };
}
