import type { CellValue } from '@model/value';
import type { AlignmentStyle, FillStyle, FontStyle } from '@model/types';

export interface ProjectedStyle {
  readonly font: FontStyle;
  readonly fill: FillStyle;
  readonly alignment: AlignmentStyle;
  /** The authored number-format code, or `null` for the general format. */
  readonly numberFormat: string | null;
}

/** The span of a merged range, anchored at its top-left cell. */
export interface MergeSpan {
  readonly anchorRow: number;
  readonly anchorCol: number;
  readonly rowSpan: number;
  readonly colSpan: number;
}

export interface ProjectedCell {
  readonly row: number;
  readonly col: number;
  readonly value: CellValue;
  readonly displayText: string;
  /** The raw value or formula, as the formula bar and the editor show it. */
  readonly rawText: string;
  readonly formula: string | null;
  readonly styleId: number;
  readonly stale: boolean;
  readonly style: ProjectedStyle;
  /** The merged range this cell belongs to, or `null`. */
  readonly merge: MergeSpan | null;
}

export interface SheetProjection {
  readonly sheetId: string;
  readonly sheetName: string;
  readonly version: number;
  readonly rowCount: number;
  readonly columnCount: number;
  readonly freezeRows: number;
  readonly freezeCols: number;
  readonly showGridLines: boolean;
  readonly showHeadings: boolean;
  readonly date1904: boolean;
  cellAt(row: number, col: number): ProjectedCell | null;
  window(startRow: number, endRow: number, startCol: number, endCol: number): ProjectedCell[];
  rowHeight(row: number): number;
  columnWidth(col: number): number;
  columnLabel(col: number): string;
  rowLabel(row: number): string;
  rowY(row: number): number;
  columnX(col: number): number;
  rowAt(y: number): number;
  columnAt(x: number): number;
  totalHeight(): number;
  totalWidth(): number;
  /** The merged range covering a cell, or `null`. */
  mergeAt(row: number, col: number): MergeSpan | null;
  /** The used range of the sheet, or `null` when the sheet is empty. */
  usedRange(): { startRow: number; endRow: number; startCol: number; endCol: number } | null;
  /** The sorted columns that hold data in a row. */
  rowDataColumns(row: number): readonly number[];
  /** The sorted rows that hold data in a column. */
  columnDataRows(col: number): readonly number[];
  /** The raw content of a cell, as the formula bar shows it. */
  rawTextAt(row: number, col: number): string;
}

export interface ProjectionOptions {
  readonly version: number;
  readonly defaultRowHeight?: number;
  readonly defaultColumnWidth?: number;
  readonly rowCount?: number;
  readonly columnCount?: number;
}
