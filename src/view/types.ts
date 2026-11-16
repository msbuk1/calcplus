import type { CellValue } from '@model/value';
import type {
  AlignmentStyle,
  BorderStyle,
  ColumnDimension,
  FillStyle,
  FontStyle,
  RowDimension,
} from '@model/types';

export interface ProjectedStyle {
  readonly font: FontStyle;
  readonly fill: FillStyle;
  readonly border: BorderStyle;
  readonly alignment: AlignmentStyle;
  /** The authored number-format code, or `null` for the general format. */
  readonly numberFormat: string | null;
  /** A text colour carried by the number-format code, or `null`. */
  readonly numberFormatColor: string | null;
}

/**
 * A width-aware display string. The grid paints `text`; when `overflow` is
 * `true` it is Excel's `#` overflow marker, not the formatted value.
 */
export interface ProjectedTextLayout {
  readonly text: string;
  readonly overflow: boolean;
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
  /**
   * Number-format display for a content-box width. `displayText` is
   * width-agnostic; the grid uses this when it knows the paint width so `*`
   * fill, `_` skip and Excel's `#` overflow marker behave correctly.
   */
  layoutText(widthPx: number, measure: (text: string) => number): ProjectedTextLayout;
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
  /** `true` when the imported row is hidden; a hidden row has zero height. */
  isRowHidden(row: number): boolean;
  /** `true` when the imported column is hidden; a hidden column has zero width. */
  isColumnHidden(col: number): boolean;
  /** Imported sheet defaults; `null` when the file declared none. */
  readonly defaultRowHeightPoints: number | null;
  readonly defaultColumnWidthChars: number | null;
  /** The imported dimension record for a row, or `null`. */
  rowDimension(row: number): Readonly<RowDimension> | null;
  /** The imported dimension record for a column, or `null`. */
  columnDimension(col: number): Readonly<ColumnDimension> | null;
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
  /**
   * Width-aware number-format display for a cell, or `null` when it is empty.
   * Pass the content-box width and the painter's text measure function.
   */
  layoutTextAt(
    row: number,
    col: number,
    widthPx: number,
    measure: (text: string) => number,
  ): ProjectedTextLayout | null;
}

export interface ProjectionOptions {
  readonly version: number;
  readonly defaultRowHeight?: number;
  readonly defaultColumnWidth?: number;
  readonly rowCount?: number;
  readonly columnCount?: number;
}
