import type { CellValue } from '@model/value';
import type {
  AlignmentStyle,
  BorderStyle,
  ColumnDimension,
  FillStyle,
  FontStyle,
  GridRange,
  RowDimension,
} from '@model/types';
import type {
  CfCalendarDate,
  CfCellEvaluation,
  CfCellSubject,
  CfDxfResolver,
  ConditionalFormatRule,
} from '@model/conditional-format';

/**
 * The conditional-formatting descriptors travel to the renderer through the
 * projection (ADR 0007). They are the engine's paint-ready overlay shapes,
 * pinned by the engine's own tests; the grid consumes them and never
 * evaluates a rule.
 */
export type {
  CfBorderEdgePatch,
  CfBorderPatch,
  CfCellEvaluation,
  CfCellSubject,
  CfColorScaleOverlay,
  CfDataBarOverlay,
  CfDifferentialFormat,
  CfDxfResolver,
  CfFillPatch,
  CfFontPatch,
  CfIconSetId,
  CfIconSetOverlay,
  CfNumberFormatPatch,
  CfOverlay,
  CfThreshold,
  ConditionalFormatRule,
} from '@model/conditional-format';

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

/** Sum, average and numeric count of a selection. */
export interface SelectionStatistics {
  readonly count: number;
  readonly sum: number;
  readonly average: number;
}

/**
 * Conditional formatting resolved for painting (conditional formatting §2,
 * §5 and §6). The projection owns the rule index and its caches, so the grid
 * reads paint-ready overlays and never evaluates a rule or touches the model.
 *
 * The overlays are derived, never stored: evaluating a cell changes nothing,
 * and the cell's own value and style stay exactly as authored (§2).
 */
export interface ProjectedConditionalFormatting {
  /** How many rules the sheet carries; `0` means the projection omits the seam. */
  readonly ruleCount: number;
  /**
   * Every true rule for a cell, in priority order, with the merged overlay and
   * the colour-scale, data-bar and icon-set overlays. `subject` lets the
   * renderer pass the value and display text it already resolved.
   */
  evaluate(row: number, col: number, subject?: CfCellSubject): CfCellEvaluation;
  /** The display text under a CF number format code, for the painted cell. */
  textFor(value: CellValue, code: string): string;
  /**
   * Drop cached results for one cell after a value edit. A projection is
   * rebuilt per document version, so a fresh projection needs no invalidation.
   */
  invalidateCell(row: number, col: number): void;
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
  /**
   * The display scale, where `1` is 100%. Geometry below is already scaled:
   * row heights, column widths and offsets are CSS pixels at this zoom.
   */
  readonly zoom: number;
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
  /** Sum, average and numeric count for the cells inside a selection. */
  selectionStatistics(range: GridRange): SelectionStatistics;
  /**
   * Conditional formatting for painting, or `null` when the sheet carries no
   * rendered rules. Built with the projection, so it is versioned with the
   * document.
   */
  readonly conditionalFormatting: ProjectedConditionalFormatting | null;
}

/** A transient row or column size shown while the user drags a header edge. */
export interface DimensionPreview {
  readonly axis: 'row' | 'column';
  /** Every row or column the drag changes; a multi-header selection moves together. */
  readonly indices: readonly number[];
  /** Points for a row, Excel character units for a column. */
  readonly size: number;
}

export interface ProjectionOptions {
  readonly version: number;
  readonly defaultRowHeight?: number;
  readonly defaultColumnWidth?: number;
  readonly rowCount?: number;
  readonly columnCount?: number;
  /**
   * A live resize that is not yet a document change. The projection paints and
   * hit-tests at this size on every selected visible member while the header
   * drag is in progress.
   */
  readonly dimensionPreview?: DimensionPreview | null;
  /**
   * Display scale, where `1` is 100%. Clamped to 10%-400%. A value above `4`
   * is read as an Excel percentage (`150` means 150%).
   */
  readonly zoom?: number;
  /**
   * Display overrides for the sheet's imported view state. Omitted values use
   * the imported setting. These are view state and never mutate the model.
   */
  readonly showGridLines?: boolean;
  readonly showHeadings?: boolean;
  /**
   * Conditional formatting inputs. Omitted, the projection rebuilds the rules
   * from `sheet.conditionalFormats`; `null` disables CF painting entirely.
   * The differential-format resolver is supplied by the caller because the
   * `dxfs` table lives in the styles part, which the model does not yet carry.
   */
  readonly conditionalFormats?: ConditionalFormattingInputs | null;
}

/** Optional inputs that override or extend the projection's CF rules. */
export interface ConditionalFormattingInputs {
  /** Explicit rules; omitted, the rules are parsed from the sheet's records. */
  readonly rules?: readonly ConditionalFormatRule[];
  /** Resolve a rule's `dxfId` to fill, font, border and number format. */
  readonly dxfResolver?: CfDxfResolver | null;
  /** The local calendar date for the ten relative date presets; defaults to today. */
  readonly today?: CfCalendarDate;
}
