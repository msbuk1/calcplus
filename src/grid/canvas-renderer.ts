import type { ProjectedCell, ProjectedStyle, SheetProjection } from '@view/types';
import type { GridCellAddress, GridRange, GridRenderState, RenderFrame } from './frame';
import {
  cfFillColor,
  cfPaintOps,
  paintCfDataBar,
  paintCfFill,
  paintCfIcon,
  type CfPaintOp,
  type CfTextPaint,
} from './cf-paint';
import { gridDebugState } from './debug';
import { filterButtonRect } from './filter-affordance';
import { mergedBorderStyle } from './merged-borders';
import {
  cellRect,
  computePaneLayout,
  intersectRects,
  mergeSpillsPane,
  mergedRect,
  panePaintRect,
  type CellRect,
  type PaneLayout,
  type PaneRect,
} from './panes';
import { activeCell, selectionRange } from './selection';
import {
  DEFAULT_TEXT_COLOR,
  MIN_FONT_PIXELS,
  fontFor,
  fontSizePixels,
  fitFontSize,
  horizontalFor,
  horizontalPadding,
  lineHeightFor,
  rotationFor,
  textColorFor,
  textInsets,
  verticalFor,
  wrapTextLines,
} from './text-layout';

const COLORS = Object.freeze({
  background: '#ffffff',
  headerBackground: '#f5f5f5',
  headerText: '#444444',
  headerActiveBackground: '#1a73e8',
  headerActiveText: '#ffffff',
  gridline: '#e1e1e1',
  frozenDivider: '#9aa0a6',
  text: DEFAULT_TEXT_COLOR,
  activeBorder: '#1a73e8',
  selectionFill: 'rgba(26, 115, 232, 0.10)',
  selectionBorder: '#1a73e8',
  fillHandle: '#1a73e8',
  painterFill: 'rgba(237, 125, 49, 0.18)',
  painterBorder: '#c55a11',
  /** The CSS theme accent, read at paint time; this is the fallback. */
  link: '#107c41',
});

/** Excel's default automatic border colour. */
const DEFAULT_BORDER_COLOR = '#000000';
/** The note/comment indicator triangle (file-compat §9 F12). */
export const COMMENT_INDICATOR_COLOR = '#c00000';
const MEASURE_CACHE_LIMIT = 8192;

interface Pane {
  readonly rect: PaneRect;
  readonly startRow: number;
  readonly endRow: number;
  readonly startCol: number;
  readonly endCol: number;
  readonly frozenRows: boolean;
  readonly frozenCols: boolean;
}

interface BorderMetrics {
  readonly width: number;
  readonly dash: readonly number[];
  readonly double: boolean;
}

/** One visible cell, resolved once per pane and painted in the pane's passes. */
interface VisibleCell {
  readonly cell: ProjectedCell;
  readonly rect: CellRect;
  readonly fill: ProjectedStyle['fill'] | null;
  readonly paintText: boolean;
  readonly paintBorder: boolean;
  /** A covered cell of a merge paints the anchor's fill and nothing else. */
  readonly covered: boolean;
  /**
   * The conditional-formatting plan for this cell, in §5.1 order. Empty when
   * no rule paints it; a covered cell never carries a plan.
   */
  readonly cf: readonly CfPaintOp[];
  /** The anchor's CF fill colour, for a covered cell of a merged range. */
  readonly cfFill: string | null;
}

/** A merged range that needs painting outside the pane that holds its anchor. */
interface MergedOverflow {
  readonly cell: ProjectedCell;
  readonly rect: CellRect;
  readonly clip: CellRect;
  /** The anchor's CF plan, painted whole and clipped at the pane edge (§5.2). */
  readonly cf: readonly CfPaintOp[];
}

export class CanvasRenderFrame implements RenderFrame {
  private readonly canvas: HTMLCanvasElement;
  private readonly context: CanvasRenderingContext2D;
  private readonly measureCache = new Map<string, number>();
  private readonly patternCache = new Map<string, CanvasPattern | null>();
  private readonly mergedBorderCache = new WeakMap<
    SheetProjection,
    Map<string, ProjectedStyle['border']>
  >();
  private width = 0;
  private height = 0;
  private dpr = 1;
  private lastFont = '';
  /** The resolved theme accent for a hyperlink, cached after the first link. */
  private linkColor: string | null = null;

  constructor(canvas: HTMLCanvasElement) {
    const context = canvas.getContext('2d');
    if (context === null) throw new Error('Canvas 2D is unavailable');
    this.canvas = canvas;
    this.context = context;
  }

  resize(width: number, height: number, dpr: number): void {
    this.width = width;
    this.height = height;
    this.dpr = dpr;
    this.canvas.width = Math.max(1, Math.round(width * dpr));
    this.canvas.height = Math.max(1, Math.round(height * dpr));
    this.canvas.style.width = `${width}px`;
    this.canvas.style.height = `${height}px`;
  }

  /** Measure text with the canvas metrics autofit needs, using the same cache. */
  measureText(text: string, font: string): number {
    return this.measure(text, font);
  }

  render(state: GridRenderState): void {
    const { context, width, height, dpr } = this;
    const { projection, scrollX, scrollY } = state;
    const layout = computePaneLayout(projection, width, height);
    context.setTransform(dpr, 0, 0, dpr, 0, 0);
    context.clearRect(0, 0, width, height);
    context.fillStyle = COLORS.background;
    context.fillRect(0, 0, width, height);
    this.lastFont = '';
    // Cell text sits on an alphabetic baseline. The header pass sets a middle
    // baseline for its labels, so restore the cell baseline before every frame.
    context.textBaseline = 'alphabetic';

    const panes = buildPanes(projection, layout, scrollX, scrollY);
    const overflow: MergedOverflow[] = [];
    for (const pane of panes) {
      this.paintPane(pane, projection, layout, scrollX, scrollY, overflow);
    }
    this.paintMergedOverflow(overflow, projection);
    this.paintPointing(state, layout, panes, scrollX, scrollY);
    this.paintPainter(state, layout, panes, scrollX, scrollY);
    this.paintSelection(state, layout, panes, scrollX, scrollY);
    this.paintHeaders(state, layout, panes, scrollX, scrollY);
  }

  /**
   * While a formula is being pointed at the sheet, the reference range is
   * tinted even though keyboard focus stays in the in-cell editor.
   */
  private paintPointing(
    state: GridRenderState,
    layout: PaneLayout,
    panes: readonly Pane[],
    scrollX: number,
    scrollY: number,
  ): void {
    const range = state.pointing ?? null;
    if (range === null) return;
    this.paintRangeHighlight(
      state,
      range,
      layout,
      panes,
      scrollX,
      scrollY,
      COLORS.selectionFill,
      COLORS.selectionBorder,
      false,
    );
  }

  /** The Format Painter target follows the pointer before a single apply. */
  private paintPainter(
    state: GridRenderState,
    layout: PaneLayout,
    panes: readonly Pane[],
    scrollX: number,
    scrollY: number,
  ): void {
    const range = state.painter ?? null;
    if (range === null) return;
    this.paintRangeHighlight(
      state,
      range,
      layout,
      panes,
      scrollX,
      scrollY,
      COLORS.painterFill,
      COLORS.painterBorder,
      true,
    );
  }

  private paintRangeHighlight(
    state: GridRenderState,
    range: NonNullable<GridRenderState['pointing']>,
    layout: PaneLayout,
    panes: readonly Pane[],
    scrollX: number,
    scrollY: number,
    fill: string,
    stroke: string,
    dashed: boolean,
  ): void {
    const { context } = this;
    context.save();
    context.fillStyle = fill;
    for (const pane of panes) {
      const startRow = Math.max(range.startRow, pane.startRow);
      const endRow = Math.min(range.endRow, pane.endRow);
      const startCol = Math.max(range.startCol, pane.startCol);
      const endCol = Math.min(range.endCol, pane.endCol);
      if (startRow > endRow || startCol > endCol) continue;
      const topLeft = cellRect(state.projection, layout, scrollX, scrollY, startRow, startCol);
      const bottomRight = cellRect(state.projection, layout, scrollX, scrollY, endRow, endCol);
      context.fillRect(
        topLeft.x,
        topLeft.y,
        bottomRight.x + bottomRight.width - topLeft.x,
        bottomRight.y + bottomRight.height - topLeft.y,
      );
    }
    context.strokeStyle = stroke;
    context.lineWidth = 2;
    context.setLineDash(dashed ? [6, 3] : []);
    for (const pane of panes) {
      const startRow = Math.max(range.startRow, pane.startRow);
      const endRow = Math.min(range.endRow, pane.endRow);
      const startCol = Math.max(range.startCol, pane.startCol);
      const endCol = Math.min(range.endCol, pane.endCol);
      if (startRow > endRow || startCol > endCol) continue;
      const topLeft = cellRect(state.projection, layout, scrollX, scrollY, startRow, startCol);
      const bottomRight = cellRect(state.projection, layout, scrollX, scrollY, endRow, endCol);
      context.strokeRect(
        topLeft.x + 1,
        topLeft.y + 1,
        Math.max(0, bottomRight.x + bottomRight.width - topLeft.x - 2),
        Math.max(0, bottomRight.y + bottomRight.height - topLeft.y - 2),
      );
    }
    context.restore();
  }

  private paintPane(
    pane: Pane,
    projection: SheetProjection,
    layout: PaneLayout,
    scrollX: number,
    scrollY: number,
    overflow: MergedOverflow[],
  ): void {
    if (pane.rect.width <= 0 || pane.rect.height <= 0) return;
    const { context } = this;
    context.save();
    context.beginPath();
    context.rect(pane.rect.x, pane.rect.y, pane.rect.width, pane.rect.height);
    context.clip();

    const rowOf = (row: number): number =>
      pane.frozenRows
        ? layout.headerHeight + projection.rowY(row)
        : layout.headerHeight + layout.frozenHeight + projection.rowY(row) - scrollY;
    const colOf = (col: number): number =>
      pane.frozenCols
        ? layout.headerWidth + projection.columnX(col)
        : layout.headerWidth + layout.frozenWidth + projection.columnX(col) - scrollX;

    const cells = collectVisibleCells(pane, projection, layout, scrollX, scrollY, rowOf, colOf);

    // A merged range that reaches past this pane is painted once, after every
    // pane (see `paintMergedOverflow`). Here the pane clip would cut it at the
    // freeze divider, and a pane painted later fills the covered cells over its
    // text. The covered cells still paint the anchor's fill in their own pane.
    const owned: VisibleCell[] = [];
    for (const visible of cells) {
      const merge = visible.cell.merge;
      if (!visible.covered && merge !== null && mergeSpillsPane(merge, pane)) {
        overflow.push({
          cell: visible.cell,
          rect: visible.rect,
          clip: intersectRects(panePaintRect(layout, pane), visible.rect),
          cf: visible.cf,
        });
        continue;
      }
      owned.push(visible);
    }

    // Gridlines sit behind fills, as Excel paints them. A fill covers the
    // gridline at its own edge; the shared edge between two filled cells
    // disappears, matching Excel.
    if (projection.showGridLines) {
      this.paintGridlines(pane, projection, rowOf, colOf);
    }
    // §5.1: the manual fill, then the CF fill or colour-scale fill. A CF fill
    // covers the gridlines, so it paints after the gridline pass. A covered
    // cell of a merge paints the anchor's CF fill in its own rectangle.
    for (const visible of owned) {
      if (visible.fill !== null) this.paintFill(visible.rect, visible.fill);
    }
    for (const visible of owned) {
      if (visible.covered) {
        if (visible.cfFill === null) continue;
        paintCfFill(this.context, { kind: 'fill', rect: visible.rect, color: visible.cfFill });
        continue;
      }
      const fill = opOfKind(visible.cf, 'fill');
      if (fill !== undefined) paintCfFill(this.context, fill);
    }
    // §5.1: the data bar sits under the text and over the fill.
    for (const visible of owned) {
      const bar = opOfKind(visible.cf, 'dataBar');
      if (bar !== undefined) paintCfDataBar(this.context, bar);
    }
    // §5.1: the cell text, shifted right when an icon is shown, in the CF font
    // and CF number format when the rule sets one.
    for (const visible of owned) {
      if (!visible.paintText) continue;
      const text = opOfKind(visible.cf, 'text');
      if (text !== undefined) {
        if (text.hidden || text.text === '') continue;
        this.paintText(visible.cell, visible.rect, projection.zoom, text);
        continue;
      }
      if (visible.cell.displayText !== '') {
        this.paintText(visible.cell, visible.rect, projection.zoom);
      }
    }
    // §5.1: the manual border, then the CF border over it.
    for (const visible of owned) {
      if (!visible.paintBorder) continue;
      const merge = visible.cell.merge;
      this.paintBorderRect(
        merge === null ? visible.cell.style.border : this.mergedBorderFor(projection, merge),
        visible.rect,
      );
    }
    for (const visible of owned) {
      const border = opOfKind(visible.cf, 'border');
      if (border !== undefined) this.paintBorderRect(border.border, visible.rect);
    }
    // §5.1: the icon last, over the text baseline.
    for (const visible of owned) {
      const icon = opOfKind(visible.cf, 'icon');
      if (icon !== undefined) paintCfIcon(this.context, icon);
    }

    // The AutoFilter dropdown button sits inside each header cell of the filter
    // range (sort-autofilter §3.8); a filtered column shows the funnel form.
    this.paintFilterHeaders(pane, projection, layout, scrollX, scrollY);

    // File-compat §9 F12: the note/comment indicator is a small red triangle in
    // the cell's top-right corner, painted last so a fill or CF never hides it.
    this.paintCommentIndicators(pane, projection, rowOf, colOf);

    context.restore();
  }

  /**
   * Paint the merged ranges that span more than one pane. The range is a single
   * cell in the model, so it is painted whole and clipped at its own edge
   * instead of at the divider of the pane that holds its anchor. The clip keeps
   * it inside the grid body and inside the owning pane on any axis that pane
   * scrolls, where the neighbouring strips are already on the canvas.
   */
  private paintMergedOverflow(
    ranges: readonly MergedOverflow[],
    projection: SheetProjection,
  ): void {
    if (ranges.length === 0) return;
    const { context } = this;
    for (const range of ranges) {
      const { cell, rect, clip, cf } = range;
      if (clip.width <= 0 || clip.height <= 0) continue;
      context.save();
      context.beginPath();
      context.rect(clip.x, clip.y, clip.width, clip.height);
      context.clip();
      // §5.1 order, evaluated once at the anchor and painted across the whole
      // merged rectangle: manual fill, CF fill, data bar, manual border, CF
      // border, text, icon. §5.3: the clip above keeps every stroke inside the
      // pane's own edge on the axis it scrolls.
      if (cell.style.fill !== null) this.paintFill(rect, cell.style.fill);
      const cfFill = opOfKind(cf, 'fill');
      if (cfFill !== undefined) paintCfFill(context, cfFill);
      const bar = opOfKind(cf, 'dataBar');
      if (bar !== undefined) paintCfDataBar(context, bar);
      const merge = cell.merge;
      if (merge !== null) {
        this.paintBorderRect(this.mergedBorderFor(projection, merge), rect);
      }
      const cfBorder = opOfKind(cf, 'border');
      if (cfBorder !== undefined) this.paintBorderRect(cfBorder.border, rect);
      const cfText = opOfKind(cf, 'text');
      if (cfText !== undefined) {
        if (!cfText.hidden && cfText.text !== '') {
          this.paintText(cell, rect, projection.zoom, cfText);
        }
      } else if (cell.displayText !== '') {
        this.paintText(cell, rect, projection.zoom);
      }
      const icon = opOfKind(cf, 'icon');
      if (icon !== undefined) paintCfIcon(context, icon);
      context.restore();
    }
  }

  /**
   * Paint the corner marker for every cell that carries a note or threaded
   * comment. The projection supplies the key set (`commentCells`), so the
   * renderer never reads the model or the comment text. A comment may be
   * anchored to an empty cell, so this pass walks the set rather than the
   * visible cells.
   */
  private paintCommentIndicators(
    pane: Pane,
    projection: SheetProjection,
    rowOf: (row: number) => number,
    colOf: (col: number) => number,
  ): void {
    const cells = projection.commentCells;
    if (cells === undefined || cells.size === 0) return;
    const { context } = this;
    const size = Math.max(3, Math.min(8, Math.round(5 * projection.zoom)));
    context.fillStyle = COMMENT_INDICATOR_COLOR;
    for (const key of cells) {
      const separator = key.indexOf(':');
      if (separator < 0) continue;
      const row = Number(key.slice(0, separator));
      const col = Number(key.slice(separator + 1));
      if (!Number.isInteger(row) || !Number.isInteger(col)) continue;
      if (row < pane.startRow || row > pane.endRow) continue;
      if (col < pane.startCol || col > pane.endCol) continue;
      if (projection.isRowHidden(row) || projection.isColumnHidden(col)) continue;
      const x = colOf(col) + projection.columnWidth(col);
      const y = rowOf(row);
      context.beginPath();
      context.moveTo(x - size, y);
      context.lineTo(x, y);
      context.lineTo(x, y + size);
      context.closePath();
      context.fill();
    }
  }

  private paintFill(rect: CellRect, fill: ProjectedStyle['fill']): void {
    const { x, y, width, height } = rect;
    if (fill.pattern === 'none') return;
    if (fill.pattern === 'solid' || fill.pattern === '') {
      if (fill.color === null) return;
      this.context.fillStyle = fill.color;
      this.context.fillRect(x, y, width, height);
      return;
    }
    if (fill.color === null) {
      // A pattern fill with no foreground colour still paints its background.
      // The dialog allows that combination; omitting it would make the cell
      // look unfilled even though the style carries a colour.
      const background = fill.backgroundColor;
      if (background === null || background === undefined) return;
      this.context.fillStyle = background;
      this.context.fillRect(x, y, width, height);
      return;
    }
    const pattern = this.patternFor(fill.pattern, fill.color);
    if (pattern === null) {
      this.context.fillStyle = fill.color;
      this.context.fillRect(x, y, width, height);
      return;
    }
    // Excel's pattern fill paints over its background colour; the model keeps
    // the imported background when the source declares one.
    this.context.fillStyle = fill.backgroundColor ?? '#ffffff';
    this.context.fillRect(x, y, width, height);
    this.context.fillStyle = pattern;
    this.context.fillRect(x, y, width, height);
  }

  private applyPattern(context: CanvasRenderingContext2D, pattern: string, color: string): void {
    context.fillStyle = color;
    context.strokeStyle = color;
    context.lineWidth = 1;
    context.globalAlpha = 1;
    switch (pattern) {
      case 'gray125':
        context.fillRect(1, 1, 1, 1);
        context.fillRect(5, 5, 1, 1);
        break;
      case 'gray0625':
        context.fillRect(1, 1, 1, 1);
        break;
      case 'lightGray':
        context.globalAlpha = 0.25;
        context.fillRect(0, 0, 8, 8);
        break;
      case 'mediumGray':
        context.globalAlpha = 0.5;
        context.fillRect(0, 0, 8, 8);
        break;
      case 'darkGray':
        context.globalAlpha = 0.75;
        context.fillRect(0, 0, 8, 8);
        break;
      case 'lightHorizontal':
        context.beginPath();
        context.moveTo(0, 6.5);
        context.lineTo(8, 6.5);
        context.stroke();
        break;
      case 'darkHorizontal':
        context.beginPath();
        context.moveTo(0, 2.5);
        context.lineTo(8, 2.5);
        context.moveTo(0, 6.5);
        context.lineTo(8, 6.5);
        context.stroke();
        break;
      case 'lightVertical':
        context.beginPath();
        context.moveTo(6.5, 0);
        context.lineTo(6.5, 8);
        context.stroke();
        break;
      case 'darkVertical':
        context.beginPath();
        context.moveTo(2.5, 0);
        context.lineTo(2.5, 8);
        context.moveTo(6.5, 0);
        context.lineTo(6.5, 8);
        context.stroke();
        break;
      case 'lightDown':
        context.beginPath();
        context.moveTo(0.5, 7.5);
        context.lineTo(7.5, 0.5);
        context.stroke();
        break;
      case 'darkDown':
        context.beginPath();
        context.moveTo(-2.5, 7.5);
        context.lineTo(7.5, -2.5);
        context.moveTo(2.5, 11.5);
        context.lineTo(11.5, 2.5);
        context.stroke();
        break;
      case 'lightUp':
        context.beginPath();
        context.moveTo(0.5, 0.5);
        context.lineTo(7.5, 7.5);
        context.stroke();
        break;
      case 'darkUp':
        context.beginPath();
        context.moveTo(-2.5, 0.5);
        context.lineTo(7.5, 10.5);
        context.moveTo(2.5, -3.5);
        context.lineTo(11.5, 6.5);
        context.stroke();
        break;
      case 'lightGrid':
        context.beginPath();
        context.moveTo(0, 3.5);
        context.lineTo(8, 3.5);
        context.moveTo(3.5, 0);
        context.lineTo(3.5, 8);
        context.stroke();
        break;
      case 'darkGrid':
        context.beginPath();
        context.moveTo(0, 1.5);
        context.lineTo(8, 1.5);
        context.moveTo(0, 5.5);
        context.lineTo(8, 5.5);
        context.moveTo(1.5, 0);
        context.lineTo(1.5, 8);
        context.moveTo(5.5, 0);
        context.lineTo(5.5, 8);
        context.stroke();
        break;
      case 'lightTrellis':
        context.beginPath();
        context.moveTo(0, 4);
        context.lineTo(4, 0);
        context.moveTo(4, 8);
        context.lineTo(8, 4);
        context.stroke();
        break;
      case 'darkTrellis':
        context.beginPath();
        context.moveTo(0, 2);
        context.lineTo(2, 0);
        context.moveTo(4, 8);
        context.lineTo(8, 4);
        context.moveTo(0, 6);
        context.lineTo(2, 8);
        context.moveTo(6, 0);
        context.lineTo(8, 2);
        context.stroke();
        break;
      default:
        // Unknown pattern names keep the colour visible as a light hatch.
        context.globalAlpha = 0.5;
        context.fillRect(0, 0, 8, 8);
        break;
    }
  }

  private patternFor(pattern: string, color: string): CanvasPattern | null {
    const key = `${pattern}\u0000${color}`;
    const cached = this.patternCache.get(key);
    if (cached !== undefined) return cached;
    if (typeof document === 'undefined') return null;
    const tile = document.createElement('canvas');
    tile.width = 8;
    tile.height = 8;
    const context = tile.getContext('2d');
    if (context === null) return null;
    this.applyPattern(context, pattern, color);
    const created = this.context.createPattern(tile, 'repeat');
    if (this.patternCache.size >= 128) this.patternCache.clear();
    this.patternCache.set(key, created);
    return created;
  }

  private paintGridlines(
    pane: Pane,
    projection: SheetProjection,
    rowOf: (row: number) => number,
    colOf: (col: number) => number,
  ): void {
    const { context } = this;
    context.strokeStyle = COLORS.gridline;
    context.lineWidth = 1;
    context.setLineDash([]);
    context.beginPath();
    // Each line is drawn per visible row or column and skipped inside a
    // merged range, so a merge has no gridline through it.
    for (let row = pane.startRow; row <= pane.endRow; row += 1) {
      if (projection.isRowHidden(row)) continue;
      const height = projection.rowHeight(row);
      if (height <= 0) continue;
      const top = Math.round(rowOf(row)) + 0.5;
      for (let col = Math.max(0, pane.startCol); col <= pane.endCol + 1; col += 1) {
        if (col >= projection.columnCount) break;
        if (col > 0 && projection.isColumnHidden(col)) continue;
        const left = Math.round(colOf(col)) + 0.5;
        if (sameMerge(projection, row, col - 1, row, col)) continue;
        context.moveTo(left, top);
        context.lineTo(left, top + Math.max(0, height));
      }
    }
    for (let col = pane.startCol; col <= pane.endCol; col += 1) {
      if (projection.isColumnHidden(col)) continue;
      const width = projection.columnWidth(col);
      if (width <= 0) continue;
      const left = Math.round(colOf(col)) + 0.5;
      for (let row = Math.max(0, pane.startRow); row <= pane.endRow + 1; row += 1) {
        if (row >= projection.rowCount) break;
        if (row > 0 && projection.isRowHidden(row)) continue;
        const top = Math.round(rowOf(row)) + 0.5;
        if (sameMerge(projection, row - 1, col, row, col)) continue;
        context.moveTo(left, top);
        context.lineTo(left + Math.max(0, width), top);
      }
    }
    context.stroke();
    context.setLineDash([]);
  }

  private paintText(cell: ProjectedCell, rect: CellRect, zoom: number, cf?: CfTextPaint): void {
    // A CF dxf supplies the font and number format; the manual style supplies
    // everything the dxf leaves unset (§3.1). An icon shifts the content box
    // right, so the text never paints under the glyph (§5.1 step 6).
    const style = cf?.style ?? cell.style;
    const shift = cf?.shift ?? 0;
    // File-compat §9 F13: a linked cell paints in the theme accent with an
    // underline, whether or not the imported style already carries them.
    const link = cell.hyperlink !== null;
    const rotation = rotationFor(style);
    if (rotation.kind === 'stacked') {
      this.paintStackedText(cf?.text ?? cell.displayText, style, rect, zoom, link);
      return;
    }
    if (rotation.kind === 'angle') {
      this.paintRotatedText(cell, rect, rotation.radians, zoom, style, cf);
      return;
    }

    const { context } = this;
    const horizontal = horizontalFor(style, cell.value);
    const insets = textInsets(style, horizontal, zoom);
    // The content box starts after the icon; alignment is computed inside it.
    const contentX = rect.x + shift;
    const contentWidth = Math.max(1, rect.width - shift);
    const maxWidth = Math.max(1, contentWidth - insets.left - insets.right);
    const shrink = style.alignment.shrinkToFit && !style.alignment.wrapText;
    const baseSize = fontSizePixels(style, zoom);
    const baseFont = fontFor(style, baseSize);
    // The projection owns the width layout (`*` fill, `_` skip and the `#`
    // overflow marker); the renderer supplies the paint width and metrics. A
    // CF number format already produced its text, so it paints as-is.
    const laidOut =
      cf === undefined || cf.layoutText
        ? cell.layoutText(maxWidth, (value) => this.measure(value, baseFont))
        : { text: cf.text, overflow: false };
    const size = shrink
      ? fitFontSize(
          style,
          laidOut.text,
          maxWidth,
          (value, font) => this.measure(value, font),
          undefined,
          zoom,
        )
      : baseSize;
    const font = fontFor(style, size);
    // Shrinking changes the metrics, so lay the text out again at the painted
    // size to keep the fill and skip widths exact.
    const text =
      size === baseSize || (cf !== undefined && !cf.layoutText)
        ? laidOut.text
        : cell.layoutText(maxWidth, (value) => this.measure(value, font)).text;
    if (text === '') return;
    const wrap = style.alignment.wrapText;
    let lines: string[];
    if (wrap) {
      lines = wrapTextLines(text, maxWidth, (value) => this.measure(value, font));
      const capacity = Math.max(1, Math.floor(rect.height / lineHeightFor(size)));
      if (lines.length > capacity) lines = lines.slice(0, capacity);
    } else {
      lines = [text];
    }

    const lineHeight = lineHeightFor(size);
    const blockHeight = Math.max(lineHeight, lines.length * lineHeight);
    const vertical = verticalFor(style);
    let top: number;
    if (vertical === 'top') top = rect.y + 2;
    else if (vertical === 'center') top = rect.y + (rect.height - blockHeight) / 2;
    else top = rect.y + rect.height - blockHeight - 1.5;
    if (vertical === 'bottom' || vertical === 'center') {
      // A block taller than the cell still starts at the top edge.
      if (top + blockHeight > rect.y + rect.height) top = rect.y + 1;
    }

    const x =
      horizontal === 'center'
        ? contentX + contentWidth / 2
        : horizontal === 'right'
          ? rect.x + rect.width - insets.right
          : contentX + insets.left;
    this.ensureFont(font);
    context.save();
    context.beginPath();
    context.rect(rect.x, rect.y, rect.width, rect.height);
    context.clip();
    context.fillStyle = link ? this.linkColorValue() : textColorFor(style);
    context.textAlign =
      horizontal === 'center' ? 'center' : horizontal === 'right' ? 'right' : 'left';
    context.textBaseline = 'alphabetic';
    context.setLineDash([]);
    for (let index = 0; index < lines.length; index += 1) {
      const line = lines[index] ?? '';
      if (line === '') continue;
      const baseline = top + index * lineHeight + size;
      context.fillText(line, x, baseline);
      this.paintDecorations(
        style,
        line,
        x,
        baseline,
        horizontal,
        size,
        font,
        contentX + insets.left,
        Math.max(0, contentWidth - insets.left - insets.right),
        link,
      );
    }
    context.restore();
  }

  private paintRotatedText(
    cell: ProjectedCell,
    rect: CellRect,
    radians: number,
    zoom: number,
    style: ProjectedStyle = cell.style,
    cf?: CfTextPaint,
  ): void {
    const { context } = this;
    const link = cell.hyperlink !== null;
    const padding = horizontalPadding(zoom);
    // Rotation gives the cell's height to the text run.
    const available = Math.max(1, rect.height - padding * 2);
    const shrink = style.alignment.shrinkToFit;
    const baseSize = fontSizePixels(style, zoom);
    const baseFont = fontFor(style, baseSize);
    const laidOut =
      cf === undefined || cf.layoutText
        ? cell.layoutText(available, (value) => this.measure(value, baseFont))
        : { text: cf.text, overflow: false };
    const size = shrink
      ? fitFontSize(
          style,
          laidOut.text,
          available,
          (value, font) => this.measure(value, font),
          undefined,
          zoom,
        )
      : baseSize;
    const font = fontFor(style, size);
    const paint =
      size === baseSize || (cf !== undefined && !cf.layoutText)
        ? laidOut.text
        : cell.layoutText(available, (value) => this.measure(value, font)).text;
    if (paint === '') return;
    const width = this.measure(paint, font);
    this.ensureFont(font);
    context.save();
    context.beginPath();
    context.rect(rect.x, rect.y, rect.width, rect.height);
    context.clip();
    context.translate(rect.x + rect.width / 2, rect.y + rect.height / 2);
    context.rotate(radians);
    context.fillStyle = link ? this.linkColorValue() : textColorFor(style);
    context.textAlign = 'center';
    context.textBaseline = 'middle';
    context.setLineDash([]);
    context.fillText(paint, 0, 0);
    const underline = underlineKind(style);
    if (underline !== 'none' || style.font.strike || link) {
      context.strokeStyle = link ? this.linkColorValue() : textColorFor(style);
      context.lineWidth = Math.max(1, size / 14);
      const left = -width / 2;
      if (underline !== 'none' || link) {
        context.beginPath();
        context.moveTo(left, size * 0.55);
        context.lineTo(left + width, size * 0.55);
        context.stroke();
        if (underline === 'double') {
          context.beginPath();
          context.moveTo(left, size * 0.62);
          context.lineTo(left + width, size * 0.62);
          context.stroke();
        }
      }
      if (style.font.strike) {
        context.beginPath();
        context.moveTo(left, -size * 0.18);
        context.lineTo(left + width, -size * 0.18);
        context.stroke();
      }
    }
    context.restore();
  }

  private paintStackedText(
    text: string,
    style: ProjectedStyle,
    rect: CellRect,
    zoom: number,
    link = false,
  ): void {
    const { context } = this;
    const characters = Array.from(text);
    if (characters.length === 0) return;
    const padding = horizontalPadding(zoom);
    const width = Math.max(1, rect.width - padding * 2);
    const height = Math.max(1, rect.height - padding * 2);
    let size = fontSizePixels(style, zoom);
    const oneCharacter = (value: string, font: string): number => this.measure(value, font);
    let widest = 0;
    for (const character of characters) {
      widest = Math.max(widest, oneCharacter(character, fontFor(style, size)));
    }
    if (widest > width) size = Math.max(MIN_FONT_PIXELS * zoom, size * (width / widest));
    const lineHeight = Math.min(lineHeightFor(size), height / characters.length);
    const font = fontFor(style, size);
    this.ensureFont(font);
    context.save();
    context.beginPath();
    context.rect(rect.x, rect.y, rect.width, rect.height);
    context.clip();
    context.fillStyle = link ? this.linkColorValue() : textColorFor(style);
    context.textAlign = 'center';
    context.textBaseline = 'middle';
    context.setLineDash([]);
    const blockHeight = lineHeight * characters.length;
    const startY = rect.y + (rect.height - blockHeight) / 2 + lineHeight / 2;
    for (let index = 0; index < characters.length; index += 1) {
      context.fillText(
        characters[index] ?? '',
        rect.x + rect.width / 2,
        startY + index * lineHeight,
      );
    }
    context.restore();
  }

  private paintDecorations(
    style: ProjectedStyle,
    line: string,
    x: number,
    baseline: number,
    horizontal: 'left' | 'center' | 'right',
    size: number,
    font: string,
    cellLeft: number,
    cellWidth: number,
    link = false,
  ): void {
    const underline = underlineKind(style);
    if (underline === 'none' && !style.font.strike && !link) return;
    const width = this.measure(line, font);
    const left = horizontal === 'center' ? x - width / 2 : horizontal === 'right' ? x - width : x;
    const accounting = isAccountingUnderline(style);
    const lineLeft = accounting ? cellLeft : left;
    const lineWidth = accounting ? cellWidth : width;
    const { context } = this;
    const color = link ? this.linkColorValue() : textColorFor(style);
    context.strokeStyle = color;
    context.lineWidth = Math.max(1, size / 14);
    context.setLineDash([]);
    if (underline !== 'none' || link) {
      const y = baseline + Math.max(1, size * 0.08);
      context.beginPath();
      context.moveTo(lineLeft, y);
      context.lineTo(lineLeft + lineWidth, y);
      context.stroke();
      if (underline === 'double') {
        context.beginPath();
        context.moveTo(lineLeft, y + Math.max(1.5, size * 0.1));
        context.lineTo(lineLeft + lineWidth, y + Math.max(1.5, size * 0.1));
        context.stroke();
      }
    }
    if (style.font.strike) {
      context.beginPath();
      context.moveTo(left, baseline - size * 0.3);
      context.lineTo(left + width, baseline - size * 0.3);
      context.stroke();
    }
  }

  private ensureFont(font: string): void {
    if (font === this.lastFont) return;
    this.context.font = font;
    this.lastFont = font;
  }

  /**
   * The hyperlink colour is the application's theme accent, so the grid never
   * introduces a colour the chrome does not use. It is resolved lazily (the
   * stylesheet may load after the canvas) and cached for the frame's lifetime.
   */
  private linkColorValue(): string {
    if (this.linkColor !== null) return this.linkColor;
    let resolved = '';
    if (typeof getComputedStyle === 'function') {
      resolved = getComputedStyle(this.canvas).getPropertyValue('--accent').trim();
    }
    this.linkColor = resolved === '' ? COLORS.link : resolved;
    return this.linkColor;
  }

  private measure(text: string, font: string): number {
    const key = `${font}\u0000${text}`;
    const cached = this.measureCache.get(key);
    if (cached !== undefined) return cached;
    this.ensureFont(font);
    const width = this.context.measureText(text).width;
    if (this.measureCache.size >= MEASURE_CACHE_LIMIT) this.measureCache.clear();
    this.measureCache.set(key, width);
    return width;
  }

  private mergedBorderFor(
    projection: SheetProjection,
    merge: NonNullable<ProjectedCell['merge']>,
  ): ProjectedStyle['border'] {
    let cache = this.mergedBorderCache.get(projection);
    if (cache === undefined) {
      cache = new Map();
      this.mergedBorderCache.set(projection, cache);
    }
    const key = `${merge.anchorRow}:${merge.anchorCol}:${merge.rowSpan}:${merge.colSpan}`;
    const cached = cache.get(key);
    if (cached !== undefined) return cached;
    const border = mergedBorderStyle(projection, merge);
    cache.set(key, border);
    return border;
  }

  private paintBorderRect(border: ProjectedStyle['border'], rect: CellRect): void {
    if (
      border.left.style === 'none' &&
      border.right.style === 'none' &&
      border.top.style === 'none' &&
      border.bottom.style === 'none' &&
      (border.diagonal.style === 'none' || (!border.diagonalUp && !border.diagonalDown))
    ) {
      return;
    }
    const right = rect.x + rect.width;
    const bottom = rect.y + rect.height;
    this.paintBorderLine(
      rect.x,
      rect.y,
      rect.x,
      bottom,
      border.left.style,
      border.left.color,
      border.left.weight,
    );
    this.paintBorderLine(
      right,
      rect.y,
      right,
      bottom,
      border.right.style,
      border.right.color,
      border.right.weight,
    );
    this.paintBorderLine(
      rect.x,
      rect.y,
      right,
      rect.y,
      border.top.style,
      border.top.color,
      border.top.weight,
    );
    this.paintBorderLine(
      rect.x,
      bottom,
      right,
      bottom,
      border.bottom.style,
      border.bottom.color,
      border.bottom.weight,
    );
    if (border.diagonal.style !== 'none') {
      const inset = 1;
      if (border.diagonalUp) {
        this.paintBorderLine(
          rect.x + inset,
          bottom - inset,
          right - inset,
          rect.y + inset,
          border.diagonal.style,
          border.diagonal.color,
          border.diagonal.weight,
        );
      }
      if (border.diagonalDown) {
        this.paintBorderLine(
          rect.x + inset,
          rect.y + inset,
          right - inset,
          bottom - inset,
          border.diagonal.style,
          border.diagonal.color,
          border.diagonal.weight,
        );
      }
    }
  }

  private paintBorderLine(
    x1: number,
    y1: number,
    x2: number,
    y2: number,
    style: string,
    color: string | null,
    weight: ProjectedStyle['border']['left']['weight'],
  ): void {
    if (style === 'none' || style === '') return;
    const metrics = borderMetrics(style, weight);
    if (metrics.width <= 0) return;
    const { context } = this;
    context.save();
    context.strokeStyle = color ?? DEFAULT_BORDER_COLOR;
    context.lineWidth = metrics.width;
    context.setLineDash(metrics.dash);
    if (metrics.double) {
      const horizontal = y1 === y2;
      const vertical = x1 === x2;
      if (horizontal) {
        const y = snap(y1, 1);
        context.beginPath();
        context.moveTo(x1, y - 1);
        context.lineTo(x2, y - 1);
        context.moveTo(x1, y + 1);
        context.lineTo(x2, y + 1);
        context.stroke();
      } else if (vertical) {
        const x = snap(x1, 1);
        context.beginPath();
        context.moveTo(x - 1, y1);
        context.lineTo(x - 1, y2);
        context.moveTo(x + 1, y1);
        context.lineTo(x + 1, y2);
        context.stroke();
      } else {
        // A diagonal has no parallel offset that reads as double; one crisp
        // stroke keeps the imported edge visible.
        context.beginPath();
        context.moveTo(x1, y1);
        context.lineTo(x2, y2);
        context.stroke();
      }
    } else {
      const x = snap(x1, metrics.width);
      const y = snap(y1, metrics.width);
      const endX = snap(x2, metrics.width);
      const endY = snap(y2, metrics.width);
      context.beginPath();
      context.moveTo(x, y);
      context.lineTo(endX, endY);
      context.stroke();
    }
    context.setLineDash([]);
    context.restore();
  }

  private paintSelection(
    state: GridRenderState,
    layout: PaneLayout,
    panes: readonly Pane[],
    scrollX: number,
    scrollY: number,
  ): void {
    const { context } = this;
    const selected = selectionRange(state.selection);
    // A merged cell is one cell (ticket 38). A collapsed selection that lands
    // inside a merged range paints the whole merged rectangle; otherwise the
    // selected range is already the geometry to paint.
    const range = selectionPaintRange(state.projection, selected);
    // A collapsed selection is one cell: the active-cell box is the selection,
    // and it keeps the single-cell look. A range leaves the active cell
    // unfilled — the grid background shows through it — with no outline of its
    // own; the range keeps its border and its fill handle (ticket 31 A).
    const collapsed = range.startRow === range.endRow && range.startCol === range.endCol;
    const activeRect = this.activeCellRect(
      state.projection,
      layout,
      scrollX,
      scrollY,
      activeCell(state.selection),
    );
    const hole = collapsed ? null : activeRect;
    context.fillStyle = COLORS.selectionFill;
    let handleCorner: { x: number; y: number } | null = null;
    for (const pane of panes) {
      const startRow = Math.max(range.startRow, pane.startRow);
      const endRow = Math.min(range.endRow, pane.endRow);
      const startCol = Math.max(range.startCol, pane.startCol);
      const endCol = Math.min(range.endCol, pane.endCol);
      if (startRow > endRow || startCol > endCol) continue;
      const topLeft = cellRect(state.projection, layout, scrollX, scrollY, startRow, startCol);
      const bottomRight = cellRect(state.projection, layout, scrollX, scrollY, endRow, endCol);
      const x = topLeft.x;
      const y = topLeft.y;
      const width = bottomRight.x + bottomRight.width - x;
      const height = bottomRight.y + bottomRight.height - y;
      fillRectExcept(context, x, y, width, height, hole);
      if (
        handleCorner === null ||
        bottomRight.x + bottomRight.width >= handleCorner.x ||
        bottomRight.y + bottomRight.height >= handleCorner.y
      ) {
        handleCorner = {
          x: bottomRight.x + bottomRight.width,
          y: bottomRight.y + bottomRight.height,
        };
      }
    }

    context.strokeStyle = COLORS.selectionBorder;
    context.lineWidth = 2;
    for (const pane of panes) {
      const startRow = Math.max(range.startRow, pane.startRow);
      const endRow = Math.min(range.endRow, pane.endRow);
      const startCol = Math.max(range.startCol, pane.startCol);
      const endCol = Math.min(range.endCol, pane.endCol);
      if (startRow > endRow || startCol > endCol) continue;
      const topLeft = cellRect(state.projection, layout, scrollX, scrollY, startRow, startCol);
      const bottomRight = cellRect(state.projection, layout, scrollX, scrollY, endRow, endCol);
      const x = topLeft.x;
      const y = topLeft.y;
      const width = bottomRight.x + bottomRight.width - x;
      const height = bottomRight.y + bottomRight.height - y;
      context.strokeRect(x + 1, y + 1, Math.max(0, width - 2), Math.max(0, height - 2));
    }

    if (activeRect !== null) {
      // The active-cell box belongs to a collapsed selection only.
      if (collapsed) {
        context.strokeStyle = COLORS.activeBorder;
        context.lineWidth = 2;
        context.strokeRect(
          activeRect.x + 1,
          activeRect.y + 1,
          Math.max(0, activeRect.width - 2),
          Math.max(0, activeRect.height - 2),
        );
      }
      // The fill handle sits at the bottom-right corner of the selection.
      const handleX = activeRect.x + activeRect.width - 3;
      const handleY = activeRect.y + activeRect.height - 3;
      const corner = handleCorner;
      const handleLeft = corner === null ? handleX : corner.x - 3;
      const handleTop = corner === null ? handleY : corner.y - 3;
      context.fillStyle = COLORS.fillHandle;
      context.fillRect(handleLeft, handleTop, 6, 6);
      context.strokeStyle = '#ffffff';
      context.lineWidth = 1;
      context.strokeRect(handleLeft + 0.5, handleTop + 0.5, 5, 5);
      gridDebugState.fillHandle = {
        x: handleLeft,
        y: handleTop,
        width: 6,
        height: 6,
      };
    }
  }

  private activeCellRect(
    projection: SheetProjection,
    layout: PaneLayout,
    scrollX: number,
    scrollY: number,
    cell: GridCellAddress,
  ): CellRect | null {
    const merge = projection.mergeAt(cell.row, cell.col);
    if (merge === null) {
      return cellRect(projection, layout, scrollX, scrollY, cell.row, cell.col);
    }
    return mergedRect(
      projection,
      layout,
      scrollX,
      scrollY,
      merge.anchorRow,
      merge.anchorCol,
      merge.rowSpan,
      merge.colSpan,
    );
  }

  private paintHeaders(
    state: GridRenderState,
    layout: PaneLayout,
    panes: readonly Pane[],
    scrollX: number,
    scrollY: number,
  ): void {
    const { context, width, height } = this;
    const { projection, selection } = state;
    // Every header of the selected extent shades, as Excel's do; the active
    // cell is shown by the unfilled cell inside the range (ticket 31 A).
    const range = selectionRange(selection);
    const headerRight = width - layout.scrollbarSize;
    const headerBottom = height - layout.scrollbarSize;

    context.fillStyle = COLORS.headerBackground;
    context.fillRect(0, 0, headerRight, layout.headerHeight);
    context.fillRect(0, 0, layout.headerWidth, headerBottom);

    if (projection.showHeadings) {
      // The header strips follow the panes: a pane that shows frozen rows
      // carries the frozen row headers, a pane that scrolls carries the
      // scrolling ones. A pane with no frozen rows or columns is empty.
      const frozenRowPane = panes.find((pane) => pane.frozenRows);
      const scrollingRowPane = panes.find((pane) => !pane.frozenRows);
      const frozenColPane = panes.find((pane) => pane.frozenCols);
      const scrollingColPane = panes.find((pane) => !pane.frozenCols);

      context.font = `${Math.max(6, 12 * projection.zoom)}px system-ui, sans-serif`;
      context.textBaseline = 'middle';
      context.textAlign = 'center';
      this.lastFont = '';

      if (frozenColPane !== undefined) {
        for (let col = frozenColPane.startCol; col <= frozenColPane.endCol; col += 1) {
          if (projection.isColumnHidden(col)) continue;
          if (projection.columnWidth(col) <= 0) continue;
          this.paintColumnHeader(
            projection,
            layout.headerWidth + projection.columnX(col),
            projection.columnWidth(col),
            layout.headerHeight,
            col,
            col >= range.startCol && col <= range.endCol,
          );
        }
      }
      if (scrollingColPane !== undefined) {
        // A partially scrolled column must not paint its header over the
        // frozen columns; clip the scrolling header strip at the divider.
        context.save();
        context.beginPath();
        context.rect(
          layout.headerWidth + layout.frozenWidth,
          0,
          Math.max(0, headerRight - layout.headerWidth - layout.frozenWidth),
          layout.headerHeight,
        );
        context.clip();
        for (let col = scrollingColPane.startCol; col <= scrollingColPane.endCol; col += 1) {
          if (projection.isColumnHidden(col)) continue;
          if (projection.columnWidth(col) <= 0) continue;
          const x = layout.headerWidth + layout.frozenWidth + projection.columnX(col) - scrollX;
          this.paintColumnHeader(
            projection,
            x,
            projection.columnWidth(col),
            layout.headerHeight,
            col,
            col >= range.startCol && col <= range.endCol,
          );
        }
        context.restore();
      }

      context.textAlign = 'right';
      if (frozenRowPane !== undefined) {
        for (let row = frozenRowPane.startRow; row <= frozenRowPane.endRow; row += 1) {
          if (projection.isRowHidden(row)) continue;
          if (projection.rowHeight(row) <= 0) continue;
          const y = layout.headerHeight + projection.rowY(row);
          this.paintRowHeader(
            projection,
            y,
            projection.rowHeight(row),
            layout.headerWidth,
            row,
            row >= range.startRow && row <= range.endRow,
          );
        }
      }
      if (scrollingRowPane !== undefined) {
        // A partially scrolled row must not paint its number over the frozen
        // rows above the divider; clip the scrolling row header strip at the
        // divider, exactly as the scrolling column strip is clipped at its own.
        context.save();
        context.beginPath();
        context.rect(
          0,
          layout.headerHeight + layout.frozenHeight,
          layout.headerWidth,
          Math.max(0, headerBottom - layout.headerHeight - layout.frozenHeight),
        );
        context.clip();
        for (let row = scrollingRowPane.startRow; row <= scrollingRowPane.endRow; row += 1) {
          if (projection.isRowHidden(row)) continue;
          if (projection.rowHeight(row) <= 0) continue;
          const y = layout.headerHeight + layout.frozenHeight + projection.rowY(row) - scrollY;
          this.paintRowHeader(
            projection,
            y,
            projection.rowHeight(row),
            layout.headerWidth,
            row,
            row >= range.startRow && row <= range.endRow,
          );
        }
        context.restore();
      }
    }

    // The frozen dividers separate the four panes.
    context.strokeStyle = COLORS.frozenDivider;
    context.lineWidth = 1;
    context.beginPath();
    if (layout.frozenWidth > 0) {
      const x = Math.round(layout.headerWidth + layout.frozenWidth) + 0.5;
      context.moveTo(x, 0);
      context.lineTo(x, headerBottom);
    }
    if (layout.frozenHeight > 0) {
      const y = Math.round(layout.headerHeight + layout.frozenHeight) + 0.5;
      context.moveTo(0, y);
      context.lineTo(headerRight, y);
    }
    context.stroke();

    // The header separators.
    context.strokeStyle = COLORS.gridline;
    context.beginPath();
    context.moveTo(0, layout.headerHeight + 0.5);
    context.lineTo(headerRight, layout.headerHeight + 0.5);
    context.moveTo(layout.headerWidth + 0.5, 0);
    context.lineTo(layout.headerWidth + 0.5, headerBottom);
    context.stroke();
  }

  /**
   * The AutoFilter affordance inside the filter range's header cells
   * (sort-autofilter §3.8). The DOM overlay in `CanvasGrid` gives the same
   * rectangle its accessible name and keyboard target.
   */
  private paintFilterHeaders(
    pane: Pane,
    projection: SheetProjection,
    layout: PaneLayout,
    scrollX: number,
    scrollY: number,
  ): void {
    const filter = projection.autoFilter;
    if (filter === null) return;
    const headerRow = filter.range.startRow;
    if (headerRow < pane.startRow || headerRow > pane.endRow) return;
    if (projection.isRowHidden(headerRow)) return;
    const firstCol = Math.max(pane.startCol, filter.range.startCol);
    const lastCol = Math.min(pane.endCol, filter.range.endCol);
    for (let col = firstCol; col <= lastCol; col += 1) {
      if (projection.isColumnHidden(col) || projection.columnWidth(col) <= 0) continue;
      const cell = cellRect(projection, layout, scrollX, scrollY, headerRow, col);
      if (cell.width < 10 || cell.height < 8) continue;
      this.paintFilterButton(
        filterButtonRect(cell, projection.zoom),
        filter.filteredColumns.has(col) || filter.unsupportedColumns.has(col),
        projection.zoom,
      );
    }
  }

  private paintFilterButton(
    rect: {
      readonly x: number;
      readonly y: number;
      readonly width: number;
      readonly height: number;
    },
    filtered: boolean,
    zoom: number,
  ): void {
    const { context } = this;
    context.save();
    context.fillStyle = 'rgba(255, 255, 255, 0.92)';
    context.fillRect(rect.x, rect.y, rect.width, rect.height);
    context.strokeStyle = '#b7b7b7';
    context.lineWidth = 1;
    context.strokeRect(
      Math.round(rect.x) + 0.5,
      Math.round(rect.y) + 0.5,
      Math.max(0, Math.round(rect.width) - 1),
      Math.max(0, Math.round(rect.height) - 1),
    );
    const centerX = rect.x + rect.width / 2;
    const centerY = rect.y + rect.height / 2;
    const halfWidth = Math.max(2, rect.width * 0.26);
    const halfHeight = Math.max(2, rect.height * 0.24);
    context.fillStyle = COLORS.headerText;
    context.strokeStyle = COLORS.headerText;
    context.lineWidth = Math.max(1, 1.2 * zoom);
    if (filtered) {
      // A funnel: wide mouth, two sloping sides and a short stem.
      context.beginPath();
      context.moveTo(centerX - halfWidth, centerY - halfHeight);
      context.lineTo(centerX + halfWidth, centerY - halfHeight);
      context.lineTo(centerX + halfWidth * 0.24, centerY + halfHeight * 0.05);
      context.lineTo(centerX + halfWidth * 0.24, centerY + halfHeight);
      context.lineTo(centerX - halfWidth * 0.24, centerY + halfHeight * 0.65);
      context.lineTo(centerX - halfWidth * 0.24, centerY + halfHeight * 0.05);
      context.closePath();
      context.fill();
    } else {
      // A chevron, Excel's unfiltered dropdown mark.
      context.beginPath();
      context.moveTo(centerX - halfWidth, centerY - halfHeight * 0.45);
      context.lineTo(centerX, centerY + halfHeight * 0.55);
      context.lineTo(centerX + halfWidth, centerY - halfHeight * 0.45);
      context.stroke();
    }
    context.restore();
  }

  private paintColumnHeader(
    projection: SheetProjection,
    x: number,
    width: number,
    headerHeight: number,
    col: number,
    selected: boolean,
  ): void {
    const { context } = this;
    if (selected) {
      context.fillStyle = COLORS.headerActiveBackground;
      context.fillRect(x, 0, width, headerHeight);
      context.fillStyle = COLORS.headerActiveText;
    } else {
      context.fillStyle = COLORS.headerText;
    }
    context.textAlign = 'center';
    context.fillText(projection.columnLabel(col), x + width / 2, headerHeight / 2);
  }

  private paintRowHeader(
    projection: SheetProjection,
    y: number,
    height: number,
    headerWidth: number,
    row: number,
    selected: boolean,
  ): void {
    const { context } = this;
    if (selected) {
      context.fillStyle = COLORS.headerActiveBackground;
      context.fillRect(0, y, headerWidth, height);
      context.fillStyle = COLORS.headerActiveText;
    } else {
      context.fillStyle = COLORS.headerText;
    }
    context.textAlign = 'right';
    context.fillText(projection.rowLabel(row), headerWidth - 8 * projection.zoom, y + height / 2);
  }

  destroy(): void {
    // The canvas is owned by React; releasing the context is enough.
    this.context.setTransform(1, 0, 0, 1, 0, 0);
    this.context.clearRect(0, 0, this.canvas.width, this.canvas.height);
    this.measureCache.clear();
    this.patternCache.clear();
  }
}

/**
 * The rectangle the selection paint covers. A collapsed selection that lands
 * inside a merged range is one merged cell, so it covers the whole range
 * (ticket 38); a selection that already spans cells keeps its rectangle,
 * which contains every merged cell it covers.
 */
function selectionPaintRange(projection: SheetProjection, range: GridRange): GridRange {
  if (range.startRow !== range.endRow || range.startCol !== range.endCol) return range;
  const merge = projection.mergeAt(range.startRow, range.startCol);
  if (merge === null) return range;
  return {
    startRow: merge.anchorRow,
    startCol: merge.anchorCol,
    endRow: merge.anchorRow + merge.rowSpan - 1,
    endCol: merge.anchorCol + merge.colSpan - 1,
  };
}

/**
 * Fill a rectangle except for an optional hole inside it. The selection tint
 * covers the whole selected range apart from the active cell, which keeps the
 * grid background (ticket 31 A).
 */
function fillRectExcept(
  context: CanvasRenderingContext2D,
  x: number,
  y: number,
  width: number,
  height: number,
  hole: CellRect | null,
): void {
  if (hole === null) {
    context.fillRect(x, y, width, height);
    return;
  }
  const left = Math.max(x, hole.x);
  const top = Math.max(y, hole.y);
  const right = Math.min(x + width, hole.x + hole.width);
  const bottom = Math.min(y + height, hole.y + hole.height);
  if (right <= left || bottom <= top) {
    context.fillRect(x, y, width, height);
    return;
  }
  // Four strips around the hole; a zero-width strip paints nothing.
  context.fillRect(x, y, width, top - y);
  context.fillRect(x, bottom, width, y + height - bottom);
  context.fillRect(x, top, left - x, bottom - top);
  context.fillRect(right, top, x + width - right, bottom - top);
}

/**
 * Resolve the visible cells once per pane. The passes then paint fills, text
 * and borders from the same records, so a frame costs one projection lookup
 * per visible cell, not one per paint pass.
 */
function collectVisibleCells(
  pane: Pane,
  projection: SheetProjection,
  layout: PaneLayout,
  scrollX: number,
  scrollY: number,
  rowOf: (row: number) => number,
  colOf: (col: number) => number,
): VisibleCell[] {
  const cells: VisibleCell[] = [];
  for (let row = pane.startRow; row <= pane.endRow; row += 1) {
    if (projection.isRowHidden(row)) continue;
    const height = projection.rowHeight(row);
    if (height <= 0) continue;
    const y = rowOf(row);
    for (let col = pane.startCol; col <= pane.endCol; col += 1) {
      if (projection.isColumnHidden(col)) continue;
      const width = projection.columnWidth(col);
      if (width <= 0) continue;
      const cell = projection.cellAt(row, col);
      if (cell === null) continue;
      const merge = cell.merge;
      if (merge !== null && (merge.anchorRow !== row || merge.anchorCol !== col)) {
        // Covered cells still paint the anchor's fill, so a merged range has
        // no hole; they never paint text, an interior border or a CF overlay.
        // The anchor's CF fill is resolved once and painted per pane (§5.3).
        const anchor = projection.cellAt(merge.anchorRow, merge.anchorCol);
        cells.push({
          cell,
          rect: { x: colOf(col), y, width, height },
          fill: anchor?.style.fill ?? null,
          paintText: false,
          paintBorder: false,
          covered: true,
          cf: [],
          cfFill:
            anchor === null || anchor === undefined
              ? null
              : cfFillAt(projection, anchor, merge.anchorRow, merge.anchorCol),
        });
        continue;
      }
      const rect: CellRect =
        merge === null
          ? { x: colOf(col), y, width, height }
          : mergedRect(
              projection,
              layout,
              scrollX,
              scrollY,
              merge.anchorRow,
              merge.anchorCol,
              merge.rowSpan,
              merge.colSpan,
            );
      cells.push({
        cell,
        rect,
        fill: cell.style.fill,
        paintText: true,
        paintBorder: true,
        covered: false,
        // §5.2: a merged range evaluates at its anchor and its plan paints the
        // whole merged rectangle — fills, bars, borders and the icon alike.
        cf: cfPlanAt(projection, cell, row, col, rect),
        cfFill: null,
      });
    }
  }
  return cells;
}

/** Evaluate one cell's rules and plan its conditional-formatting paint (§5.1). */
function cfPlanAt(
  projection: SheetProjection,
  cell: ProjectedCell,
  row: number,
  col: number,
  rect: CellRect,
): readonly CfPaintOp[] {
  const conditional = projection.conditionalFormatting;
  if (conditional === null) return [];
  const evaluation = conditional.evaluate(row, col, {
    value: cell.value,
    displayText: cell.displayText,
  });
  if (!evaluation.painted) return [];
  // §5.1 step 7: a CF number format drives the text the cell paints, while
  // the formula bar keeps showing the raw value.
  const code = evaluation.overlay.numberFormat?.code;
  return cfPaintOps({
    evaluation,
    rect,
    zoom: projection.zoom,
    style: cell.style,
    displayText: cell.displayText,
    cfDisplayText: code === undefined ? null : conditional.textFor(cell.value, code),
  });
}

/** The CF fill colour for a cell, without planning the rest of its paint. */
function cfFillAt(
  projection: SheetProjection,
  cell: ProjectedCell,
  row: number,
  col: number,
): string | null {
  const conditional = projection.conditionalFormatting;
  if (conditional === null) return null;
  const evaluation = conditional.evaluate(row, col, {
    value: cell.value,
    displayText: cell.displayText,
  });
  return cfFillColor(evaluation.overlay, evaluation.colorScale);
}

/** The first op of one kind: each kind appears at most once per cell plan. */
function opOfKind<K extends CfPaintOp['kind']>(
  ops: readonly CfPaintOp[],
  kind: K,
): Extract<CfPaintOp, { kind: K }> | undefined {
  for (const op of ops) {
    if (op.kind === kind) return op as Extract<CfPaintOp, { kind: K }>;
  }
  return undefined;
}

function buildPanes(
  projection: SheetProjection,
  layout: PaneLayout,
  scrollX: number,
  scrollY: number,
): Pane[] {
  const frozenRows = projection.freezeRows;
  const frozenCols = projection.freezeCols;
  const firstScrollRow = Math.max(frozenRows, projection.rowAt(scrollY));
  const lastScrollRow = projection.rowAt(scrollY + layout.scrollRegionHeight);
  const firstScrollCol = Math.max(frozenCols, projection.columnAt(scrollX));
  const lastScrollCol = projection.columnAt(scrollX + layout.scrollRegionWidth);

  const panes: Pane[] = [
    {
      rect: {
        x: layout.headerWidth,
        y: layout.headerHeight,
        width: layout.frozenWidth,
        height: layout.frozenHeight,
      },
      // A pane with no frozen rows or no frozen columns is empty: an
      // end index below the start index paints nothing.
      startRow: 0,
      endRow: frozenRows - 1,
      startCol: 0,
      endCol: frozenCols - 1,
      frozenRows: true,
      frozenCols: true,
    },
    {
      rect: {
        x: layout.headerWidth + layout.frozenWidth,
        y: layout.headerHeight,
        width: layout.scrollRegionWidth,
        height: layout.frozenHeight,
      },
      startRow: 0,
      endRow: frozenRows - 1,
      startCol: firstScrollCol,
      endCol: Math.max(firstScrollCol, lastScrollCol),
      frozenRows: true,
      frozenCols: false,
    },
    {
      rect: {
        x: layout.headerWidth,
        y: layout.headerHeight + layout.frozenHeight,
        width: layout.frozenWidth,
        height: layout.scrollRegionHeight,
      },
      startRow: firstScrollRow,
      endRow: Math.max(firstScrollRow, lastScrollRow),
      startCol: 0,
      endCol: frozenCols - 1,
      frozenRows: false,
      frozenCols: true,
    },
    {
      rect: {
        x: layout.headerWidth + layout.frozenWidth,
        y: layout.headerHeight + layout.frozenHeight,
        width: layout.scrollRegionWidth,
        height: layout.scrollRegionHeight,
      },
      startRow: firstScrollRow,
      endRow: Math.max(firstScrollRow, lastScrollRow),
      startCol: firstScrollCol,
      endCol: Math.max(firstScrollCol, lastScrollCol),
      frozenRows: false,
      frozenCols: false,
    },
  ];
  return panes;
}

function sameMerge(
  projection: SheetProjection,
  rowA: number,
  colA: number,
  rowB: number,
  colB: number,
): boolean {
  if (rowA < 0 || colA < 0 || rowB < 0 || colB < 0) return false;
  const left = projection.mergeAt(rowA, colA);
  if (left === null) return false;
  const right = projection.mergeAt(rowB, colB);
  return (
    right !== null &&
    left.anchorRow === right.anchorRow &&
    left.anchorCol === right.anchorCol &&
    left.rowSpan === right.rowSpan &&
    left.colSpan === right.colSpan
  );
}

function underlineKind(style: ProjectedStyle): 'none' | 'single' | 'double' {
  const variant = style.font.underlineStyle;
  if (variant === 'double' || variant === 'doubleAccounting') return 'double';
  if (variant === 'single' || variant === 'singleAccounting') return 'single';
  return style.font.underline ? 'single' : 'none';
}

function isAccountingUnderline(style: ProjectedStyle): boolean {
  return (
    style.font.underlineStyle === 'singleAccounting' ||
    style.font.underlineStyle === 'doubleAccounting'
  );
}

function borderMetrics(
  style: string,
  weight: ProjectedStyle['border']['left']['weight'],
): BorderMetrics {
  const width = weight === 'medium' || weight === 'thick' ? (weight === 'thick' ? 3 : 2) : 1;
  switch (style) {
    case 'double':
      return { width: 1, dash: [], double: true };
    case 'dotted':
      return { width: 1, dash: [1, 1], double: false };
    case 'dashed':
      return { width, dash: [4, 2], double: false };
    case 'dashDot':
      return { width, dash: [4, 2, 1, 2], double: false };
    case 'dashDotDot':
      return { width, dash: [4, 2, 1, 2, 1, 2], double: false };
    case 'mediumDashed':
      return { width: 2, dash: [5, 3], double: false };
    case 'mediumDashDot':
      return { width: 2, dash: [5, 3, 1, 3], double: false };
    case 'mediumDashDotDot':
      return { width: 2, dash: [5, 3, 1, 3, 1, 3], double: false };
    case 'slantDashDot':
      return { width: 2, dash: [5, 3, 1, 3], double: false };
    case 'hair':
    case 'thin':
    default:
      return { width, dash: [], double: false };
  }
}

function snap(value: number, lineWidth: number): number {
  return lineWidth % 2 === 1 ? Math.round(value) + 0.5 : Math.round(value);
}

/** The cell under a canvas point, or `null` when the point is in a header. */
export function cellAtPoint(
  projection: SheetProjection,
  layout: PaneLayout,
  pointX: number,
  pointY: number,
  scrollX: number,
  scrollY: number,
): GridCellAddress | null {
  if (pointX < layout.headerWidth || pointY < layout.headerHeight) return null;
  const col =
    pointX < layout.headerWidth + layout.frozenWidth
      ? projection.columnAt(pointX - layout.headerWidth)
      : projection.columnAt(pointX - layout.headerWidth - layout.frozenWidth + scrollX);
  const row =
    pointY < layout.headerHeight + layout.frozenHeight
      ? projection.rowAt(pointY - layout.headerHeight)
      : projection.rowAt(pointY - layout.headerHeight - layout.frozenHeight + scrollY);
  return { row, col };
}
