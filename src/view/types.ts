import type { CellValue } from '@model/value';

export interface ProjectedCell {
  readonly row: number;
  readonly col: number;
  readonly value: CellValue;
  readonly displayText: string;
  readonly formula: string | null;
  readonly styleId: number;
  readonly stale: boolean;
}

export interface SheetProjection {
  readonly sheetId: string;
  readonly sheetName: string;
  readonly version: number;
  readonly rowCount: number;
  readonly columnCount: number;
  cellAt(row: number, col: number): ProjectedCell | null;
  window(startRow: number, endRow: number, startCol: number, endCol: number): ProjectedCell[];
  rowHeight(row: number): number;
  columnWidth(col: number): number;
  columnLabel(col: number): string;
  rowY(row: number): number;
  columnX(col: number): number;
  rowAt(y: number): number;
  columnAt(x: number): number;
}

export interface ProjectionOptions {
  readonly version: number;
  readonly defaultRowHeight?: number;
  readonly defaultColumnWidth?: number;
  readonly rowCount?: number;
  readonly columnCount?: number;
}
