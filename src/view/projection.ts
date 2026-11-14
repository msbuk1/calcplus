import type { Workbook } from '@model/types';
import { cellKey, columnToLetters } from '@model/types';
import { formatCellValue } from '@model/value';
import type { ProjectedCell, ProjectionOptions, SheetProjection } from './types';

const DEFAULT_ROW_HEIGHT = 20;
const DEFAULT_COLUMN_WIDTH = 96;
const MAX_ROWS = 1_048_576;
const MAX_COLUMNS = 16_384;

export function createSheetProjection(
  workbook: Workbook,
  sheetId: string,
  options: ProjectionOptions,
): SheetProjection {
  const sheet = workbook.sheets.find((candidate) => candidate.id === sheetId);
  if (sheet === undefined) throw new Error(`Unknown sheet: ${sheetId}`);
  const defaultRowHeight = options.defaultRowHeight ?? DEFAULT_ROW_HEIGHT;
  const defaultColumnWidth = options.defaultColumnWidth ?? DEFAULT_COLUMN_WIDTH;
  const rowCount = options.rowCount ?? MAX_ROWS;
  const columnCount = options.columnCount ?? MAX_COLUMNS;

  // M1 paints with the default geometry. Imported row heights and column
  // widths stay in the model, untouched, for the formatting milestone.
  const rowHeightAt = (): number => defaultRowHeight;
  const columnWidthAt = (): number => defaultColumnWidth;

  const rowY = (row: number): number => row * defaultRowHeight;

  const columnX = (col: number): number => col * defaultColumnWidth;

  const rowAt = (y: number): number => {
    if (y <= 0) return 0;
    return Math.min(rowCount - 1, Math.floor(y / defaultRowHeight));
  };

  const columnAt = (x: number): number => {
    if (x <= 0) return 0;
    return Math.min(columnCount - 1, Math.floor(x / defaultColumnWidth));
  };

  return {
    sheetId,
    sheetName: sheet.name,
    version: options.version,
    rowCount,
    columnCount,
    cellAt(row, col) {
      if (row < 0 || col < 0 || row >= rowCount || col >= columnCount) return null;
      const cell = sheet.cells.get(cellKey(row, col));
      if (cell === undefined) return null;
      return {
        row,
        col,
        value: cell.value,
        displayText: formatCellValue(cell.value),
        formula: cell.formula,
        styleId: cell.styleId,
        stale: cell.formula !== null && cell.savedResult === null,
      };
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
          cells.push({
            row,
            col,
            value: cell.value,
            displayText: formatCellValue(cell.value),
            formula: cell.formula,
            styleId: cell.styleId,
            stale: cell.formula !== null && cell.savedResult === null,
          });
        }
      }
      return cells;
    },
    rowHeight: rowHeightAt,
    columnWidth: columnWidthAt,
    columnLabel: columnToLetters,
    rowY,
    columnX,
    rowAt,
    columnAt,
  };
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
