import type { BorderStyle, FontStyle, GridRange, PrintSetupRecord } from '@model/types';
import type { CellValue } from '@model/value';
import type { MergeSpan, ProjectedStyle, ProjectedTextLayout } from '@view/types';

/**
 * The print plan (print-pdf §2): the frozen contract between the pure planner
 * and the three renderers (SVG preview, PDFKit vector PDF, 300 DPI canvas
 * fallback).
 *
 * Rules the renderers may rely on:
 *
 * - **Units are points** (1/72 inch), CSS pixels at 96 dpi times 0.75. The
 *   scale is already baked into every coordinate and font size.
 * - **Origin is the page's top-left**, with y growing downwards. A page's
 *   content, its row/column bands and its items all live in that space, so a
 *   renderer positions the page box at 0,0 and draws the list.
 * - **Items paint in array order.** Gridlines first, then manual and
 *   conditional fills, then data bars, then text (and its underline/strike
 *   lines), then manual and conditional borders, then icons, as Excel paints
 *   them (conditional formatting §5.1).
 * - **A clip is a hard clip.** Text never paints outside its `clip`. Lines at
 *   a clipped page edge may extend past the page box; the renderer clips the
 *   page itself.
 * - **The plan is plain data.** No functions, no DOM, no canvas, so it can be
 *   structured-cloned to a worker and compared with `toEqual`.
 * - **A page carries only the cells of its band.** A merged range that crosses
 *   a page break is one clipped region on each page, and its text appears only
 *   on the page that holds the anchor (print-pdf §4).
 */

/** A rectangle in points, page coordinates. */
export interface PrintBox {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

export type PrintOrientation = 'portrait' | 'landscape';

/** The resolved paper: the box after the orientation is applied. */
export interface PrintPaper {
  /** The paper's recognised name (`A4`, `Letter`, …), or `Custom`. */
  readonly name: string;
  /** The OOXML `pageSetup@paperSize` code, or `null` when the sheet names none. */
  readonly code: number | null;
  readonly orientation: PrintOrientation;
  readonly width: number;
  readonly height: number;
}

/** The page margins in points, from `pageMargins` (stored in inches). */
export interface PrintMargins {
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
  readonly left: number;
}

/** The inclusive row and column range a page carries. */
export interface PrintBand {
  readonly startRow: number;
  readonly endRow: number;
  readonly startCol: number;
  readonly endCol: number;
}

/** One column of a page band, positioned in page coordinates. */
export interface PrintPageColumn {
  readonly col: number;
  readonly x: number;
  readonly width: number;
}

/** One row of a page band, positioned in page coordinates. */
export interface PrintPageRow {
  readonly row: number;
  readonly y: number;
  readonly height: number;
}

/** A fill. `pattern` is the OOXML pattern name; `solid` is a solid fill. */
export interface PrintFillItem {
  readonly kind: 'fill';
  readonly box: PrintBox;
  readonly pattern: string;
  /** The solid colour, or a pattern's foreground colour. `null` paints nothing. */
  readonly color: string | null;
  /** The colour under a pattern fill, or `null`. */
  readonly background: string | null;
}

/** A straight line: a border edge, a gridline, or a text decoration. */
export interface PrintLineItem {
  readonly kind: 'line';
  readonly x1: number;
  readonly y1: number;
  readonly x2: number;
  readonly y2: number;
  readonly color: string;
  /** Stroke width in points. */
  readonly width: number;
  /** Dash pattern in points, or `null` for a solid line. */
  readonly dash: readonly number[] | null;
}

/**
 * A conditional-formatting data bar (conditional formatting §5.1 step 3). A
 * rect variant: fill the box with `color`, or with a horizontal gradient from
 * `gradient.from` (left) to `gradient.to` (right); stroke the box at 1pt with
 * `borderColor` when it is not null; then draw a 1pt vertical line at `axisX`
 * in `axisColor` when it is not null.
 */
export interface PrintDataBarItem {
  readonly kind: 'dataBar';
  readonly box: PrintBox;
  readonly color: string;
  readonly gradient: { readonly from: string; readonly to: string } | null;
  readonly borderColor: string | null;
  /** The zero axis in page points, or `null` when the bar has none. */
  readonly axisX: number | null;
  readonly axisColor: string;
}

/**
 * The icon families a conditional-formatting icon set can paint (conditional
 * formatting §5.1 step 8). The renderers implement one glyph per family, and
 * the geometry is `paintCfIcon` in `src/grid/cf-paint.ts`: that shipped canvas
 * artwork is authoritative, and the formulas below are the ones it draws.
 *
 * The index is clamped to `0 … iconCount - 1` before any family paints, as the
 * grid clamps it.
 *
 * - `arrow`: an up-pointing filled arrow rotated clockwise by
 *   `-45° + 90° * iconIndex / (iconCount - 1) + 90°` — the canvas's
 *   `arrowAngle` plus its quarter turn from the up-facing path — so a
 *   three-icon set points up-right, sideways and down-right. A one-icon set is
 *   rotated a flat `90°`.
 * - `traffic`: a filled disc with a dark outline.
 * - `sign`: index 0 is a red disc with a white cross, the last index a green
 *   disc with a white tick, and every middle index an amber triangle with a
 *   bar and a dot.
 * - `flag`: a filled flag on a pole.
 * - `rating`: `iconCount` vertical bars; the first `iconIndex + 1` are dark
 *   `#404040`, the rest `#E3E3E3`.
 * - `quarter`: a grey disc with a dark pie wedge of
 *   `clamp((iconIndex + 1) / max(1, iconCount - 1))` of a turn, starting at
 *   12 o'clock and turning clockwise. A clamped fraction of 1 is a full dark
 *   disc, so the last wedge of a set is solid.
 * - `dot`: a filled circle.
 */
export type PrintIconFamily = 'arrow' | 'traffic' | 'sign' | 'flag' | 'rating' | 'quarter' | 'dot';

/** One resolved conditional-formatting icon glyph. */
export interface PrintIconItem {
  readonly kind: 'icon';
  readonly box: PrintBox;
  readonly family: PrintIconFamily;
  readonly color: string;
  readonly iconIndex: number;
  readonly iconCount: number;
}

/** The resolved font of a text item. */
export interface PrintFont {
  readonly family: string;
  /** Size in points, already scaled. */
  readonly size: number;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly color: string;
}

/**
 * Horizontal text anchoring. `start` puts `x` at the left edge, `middle`
 * centres the run on `x`, `end` puts the right edge at `x`.
 */
export type PrintTextAnchor = 'start' | 'middle' | 'end';

/**
 * One positioned text run. `y` is the alphabetic baseline, so a renderer never
 * needs a middle-baseline feature. `rotation` is degrees clockwise about
 * `(x, y)` (negative is counter-clockwise); the anchor point travels with the
 * rotation, so angled text stays centred in its cell.
 */
export interface PrintTextItem {
  readonly kind: 'text';
  readonly text: string;
  readonly x: number;
  readonly y: number;
  readonly anchor: PrintTextAnchor;
  readonly font: PrintFont;
  readonly rotation: number;
  readonly clip: PrintBox;
}

/**
 * A reserved image or chart anchor. v1 emits none (print-pdf §5: images and
 * charts are v2), but the vocabulary is here so the layout does not need a new
 * primitive later. `placeholder` is always `true` in v1.
 */
export interface PrintImageItem {
  readonly kind: 'image';
  readonly box: PrintBox;
  readonly anchor: { readonly row: number; readonly col: number };
  readonly placeholder: true;
}

export type PrintItem =
  PrintFillItem | PrintLineItem | PrintTextItem | PrintImageItem | PrintDataBarItem | PrintIconItem;

/** One printed page. */
export interface PrintPage {
  /** 0-based position in the whole plan. */
  readonly index: number;
  /** Which print area (page set) this page belongs to. */
  readonly setIndex: number;
  /** 0-based position inside its page set. */
  readonly pageInSetIndex: number;
  readonly paper: PrintPaper;
  /** The page box: `{ x: 0, y: 0, width, height }` in points. */
  readonly box: PrintBox;
  /** The body inside the margins. */
  readonly body: PrintBox;
  /** The inclusive cell range this page carries. Hidden rows/columns are absent. */
  readonly band: PrintBand;
  /** The band's columns, left to right, with page coordinates. */
  readonly columns: readonly PrintPageColumn[];
  /** The band's rows, top to bottom, with page coordinates. */
  readonly rows: readonly PrintPageRow[];
  /** Paint in order. */
  readonly items: readonly PrintItem[];
  /**
   * `true` when a single row or column is larger than the body and had to be
   * clipped at the page edge (print-pdf §2: clipped, never dropped).
   */
  readonly clipped: boolean;
}

/** One print area's page sequence (print-pdf §4: each area is its own set). */
export interface PrintPageSet {
  readonly index: number;
  readonly range: GridRange;
  /** The plan index of the set's first page. */
  readonly firstPage: number;
  readonly pageCount: number;
}

/** Everything a print renderer needs. */
export interface PrintPlan {
  readonly sheetId: string;
  readonly sheetName: string;
  readonly paper: PrintPaper;
  readonly margins: PrintMargins;
  /** The uniform scale; `1` is 100%. Never below 0.1 or above 4. */
  readonly scale: number;
  readonly fitToPage: boolean;
  /** `printOptions@gridLines` **and** `@gridLinesSet`; never the view state. */
  readonly showGridLines: boolean;
  /** `printOptions@headings`. */
  readonly showHeadings: boolean;
  readonly horizontalCentered: boolean;
  readonly verticalCentered: boolean;
  readonly pageSets: readonly PrintPageSet[];
  readonly pages: readonly PrintPage[];
}

// ---------------------------------------------------------------------------
// Planner input: a plain snapshot of the display records.
// ---------------------------------------------------------------------------

/** Text measurement in sheet points, supplied by the caller (never the DOM here). */
export type PrintTextMeasure = (text: string, font: PrintFont) => number;

/**
 * A cell as print consumes it. The adapter builds one from a `ProjectedCell`,
 * so print reads the same value, style, merge, hyperlink and width-aware
 * number format the grid paints (print-pdf §4).
 */
export interface PrintCell {
  readonly row: number;
  readonly col: number;
  readonly value: CellValue;
  readonly displayText: string;
  readonly style: ProjectedStyle;
  /** `true` when the cell carries a hyperlink; print paints it as a link. */
  readonly hyperlink: boolean;
  readonly merge: MergeSpan | null;
  /**
   * The width-aware number format for a content box in **sheet points** (the
   * same relative width the grid measures in CSS pixels).
   */
  readonly layoutText: (width: number, measure: PrintTextMeasure) => ProjectedTextLayout;
  /**
   * The resolved conditional-formatting paint (conditional formatting §5.1),
   * or `null`/absent when no rule paints the cell. The adapter computes it
   * with the grid's own `cfPaintOps`, so priority and stop-if-true are the
   * grid's; the planner only lays the already-ordered paint out on the page.
   */
  readonly conditional?: PrintConditionalPaint | null;
}

/**
 * One cell's conditional-formatting paint, in the grid's §5.1 order and in
 * **cell-local points** (origin at the cell's, or the merged anchor's,
 * top-left). Produced by `printConditionalFor` from `cfPaintOps`.
 */
export interface PrintConditionalPaint {
  /** The CF fill or colour-scale blend (§5.1 step 2), or `null`. */
  readonly fillColor: string | null;
  /** The data bar (§5.1 step 3), or `null`. */
  readonly dataBar: PrintConditionalDataBar | null;
  /** How the text paints (§5.1 steps 6-7); always present when painted. */
  readonly text: PrintConditionalText;
  /** The CF border edges that paint over the manual border (§5.1 step 5). */
  readonly border: BorderStyle | null;
  /** The icon glyph (§5.1 step 8), or `null`. */
  readonly icon: PrintConditionalIcon | null;
}

export interface PrintConditionalDataBar {
  readonly box: PrintBox;
  readonly color: string;
  readonly gradient: { readonly from: string; readonly to: string } | null;
  readonly borderColor: string | null;
  readonly axisX: number | null;
  readonly axisColor: string;
}

export interface PrintConditionalIcon {
  readonly box: PrintBox;
  readonly family: PrintIconFamily;
  readonly color: string;
  readonly iconIndex: number;
  readonly iconCount: number;
}

export interface PrintConditionalText {
  /** `true` for an icon-only rule or a data bar that hides its value. */
  readonly hidden: boolean;
  /** Points the text shifts right to clear an icon. */
  readonly shift: number;
  /** The manual style with the CF font and number format laid over it. */
  readonly style: ProjectedStyle;
  readonly text: string;
  /** `false` when a CF number format produced the text, so no width layout. */
  readonly layoutText: boolean;
}

/**
 * The pure planner's input. Everything here is plain data plus a few lookups;
 * no method touches the DOM, a canvas or the network.
 */
export interface PrintSheetSnapshot {
  readonly sheetId: string;
  readonly sheetName: string;
  /** The sheet's `_xlnm.Print_Area` ranges; empty falls back to the used range. */
  readonly printAreas: readonly GridRange[];
  /** The used range, or `null` when the sheet is empty. */
  readonly usedRange: GridRange | null;
  /**
   * The print setup record, or `null` when the sheet declares none. The
   * planner reads orientation, paper size, scale, `fitToPage`, margins and the
   * print options from it.
   */
  readonly printSetup: PrintSetupRecord | null;
  /** Rows and columns print at 100% in points, hidden ones at zero. */
  rowHeight(row: number): number;
  columnWidth(col: number): number;
  /** The print heading labels; the sheet owns the letters and numbers. */
  columnLabel(col: number): string;
  rowLabel(row: number): string;
  /** The font headings print in (the workbook's default font). */
  readonly headingFont: FontStyle;
  cellAt(row: number, col: number): PrintCell | null;
  /** The rows and columns of a box, cells only for the ones that exist. */
  cellsIn(startRow: number, endRow: number, startCol: number, endCol: number): readonly PrintCell[];
  /** The merged ranges intersecting the box, including anchors outside it. */
  mergesIn(
    startRow: number,
    endRow: number,
    startCol: number,
    endCol: number,
  ): readonly MergeSpan[];
}
