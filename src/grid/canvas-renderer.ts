import type { CellValue } from '@model/value';
import type { ProjectedStyle, SheetProjection } from '@view/types';
import {
  GRID_METRICS,
  type GridCellAddress,
  type GridRenderState,
  type RenderFrame,
} from './frame';
import { gridDebugState } from './debug';
import {
  cellRect,
  computePaneLayout,
  mergedRect,
  type CellRect,
  type PaneLayout,
  type PaneRect,
} from './panes';
import { selectionRange } from './selection';

const COLORS = Object.freeze({
  background: '#ffffff',
  headerBackground: '#f5f5f5',
  headerText: '#444444',
  headerActiveBackground: '#1a73e8',
  headerActiveText: '#ffffff',
  gridline: '#e1e1e1',
  frozenDivider: '#9aa0a6',
  text: '#202020',
  activeBorder: '#1a73e8',
  selectionFill: 'rgba(26, 115, 232, 0.10)',
  selectionBorder: '#1a73e8',
  fillHandle: '#1a73e8',
});

const TEXT_PADDING = 4;

interface Pane {
  readonly rect: PaneRect;
  readonly startRow: number;
  readonly endRow: number;
  readonly startCol: number;
  readonly endCol: number;
  readonly frozenRows: boolean;
  readonly frozenCols: boolean;
}

export class CanvasRenderFrame implements RenderFrame {
  private readonly canvas: HTMLCanvasElement;
  private readonly context: CanvasRenderingContext2D;
  private width = 0;
  private height = 0;
  private dpr = 1;
  private lastFont = '';

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
    for (const pane of panes) {
      this.paintPane(pane, projection, layout, scrollX, scrollY);
    }
    this.paintPointing(state, layout, panes, scrollX, scrollY);
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
    const { context } = this;
    context.fillStyle = COLORS.selectionFill;
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
      context.strokeRect(
        topLeft.x + 1,
        topLeft.y + 1,
        Math.max(0, bottomRight.x + bottomRight.width - topLeft.x - 2),
        Math.max(0, bottomRight.y + bottomRight.height - topLeft.y - 2),
      );
    }
  }

  private paintPane(
    pane: Pane,
    projection: SheetProjection,
    layout: PaneLayout,
    scrollX: number,
    scrollY: number,
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

    // Fills first, then the gridlines, then the text.
    for (let row = pane.startRow; row <= pane.endRow; row += 1) {
      const y = rowOf(row);
      const height = projection.rowHeight(row);
      if (y + height < pane.rect.y || y > pane.rect.y + pane.rect.height) continue;
      for (let col = pane.startCol; col <= pane.endCol; col += 1) {
        const cell = projection.cellAt(row, col);
        if (cell === null) continue;
        const merge = cell.merge;
        if (merge !== null && (merge.anchorRow !== row || merge.anchorCol !== col)) {
          const anchor = projection.cellAt(merge.anchorRow, merge.anchorCol);
          const fill = anchor?.style.fill.color ?? null;
          if (fill !== null) {
            context.fillStyle = fill;
            context.fillRect(colOf(col), y, projection.columnWidth(col), height);
          }
          continue;
        }
        const fill = cell.style.fill.color;
        if (fill === null) continue;
        context.fillStyle = fill;
        context.fillRect(colOf(col), y, projection.columnWidth(col), height);
      }
    }

    if (projection.showGridLines) {
      this.paintGridlines(pane, rowOf, colOf);
    }

    for (let row = pane.startRow; row <= pane.endRow; row += 1) {
      const y = rowOf(row);
      const height = projection.rowHeight(row);
      if (y + height < pane.rect.y || y > pane.rect.y + pane.rect.height) continue;
      for (let col = pane.startCol; col <= pane.endCol; col += 1) {
        const cell = projection.cellAt(row, col);
        if (cell === null || cell.displayText === '') continue;
        const merge = cell.merge;
        if (merge !== null && (merge.anchorRow !== row || merge.anchorCol !== col)) continue;
        const rect: CellRect =
          merge === null
            ? { x: colOf(col), y, width: projection.columnWidth(col), height }
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
        this.paintText(cell.displayText, cell.value, cell.style, rect);
      }
    }
    context.restore();
  }

  private paintGridlines(
    pane: Pane,
    rowOf: (row: number) => number,
    colOf: (col: number) => number,
  ): void {
    const { context } = this;
    context.strokeStyle = COLORS.gridline;
    context.lineWidth = 1;
    context.beginPath();
    for (let row = pane.startRow; row <= pane.endRow + 1; row += 1) {
      const y = Math.round(rowOf(row)) + 0.5;
      if (y < pane.rect.y - 1 || y > pane.rect.y + pane.rect.height + 1) continue;
      context.moveTo(pane.rect.x, y);
      context.lineTo(pane.rect.x + pane.rect.width, y);
    }
    for (let col = pane.startCol; col <= pane.endCol + 1; col += 1) {
      const x = Math.round(colOf(col)) + 0.5;
      if (x < pane.rect.x - 1 || x > pane.rect.x + pane.rect.width + 1) continue;
      context.moveTo(x, pane.rect.y);
      context.lineTo(x, pane.rect.y + pane.rect.height);
    }
    context.stroke();
  }

  private paintText(text: string, value: CellValue, style: ProjectedStyle, rect: CellRect): void {
    const { context } = this;
    const font = fontFor(style);
    if (font !== this.lastFont) {
      context.font = font;
      this.lastFont = font;
    }
    const horizontal = horizontalFor(style, value);
    const padding = TEXT_PADDING + style.alignment.indent * 8;
    const available = rect.width - padding * 2;
    let x = rect.x + padding;
    if (horizontal === 'center') x = rect.x + rect.width / 2;
    else if (horizontal === 'right') x = rect.x + rect.width - padding;
    context.textAlign =
      horizontal === 'center' ? 'center' : horizontal === 'right' ? 'right' : 'left';
    const baseline = verticalBaseline(style, rect);
    if (text.length * 7 > available) {
      context.save();
      context.beginPath();
      context.rect(rect.x, rect.y, rect.width, rect.height);
      context.clip();
      context.fillStyle = COLORS.text;
      context.fillText(text, x, baseline);
      context.restore();
      return;
    }
    context.fillStyle = COLORS.text;
    context.fillText(text, x, baseline);
  }

  private paintSelection(
    state: GridRenderState,
    layout: PaneLayout,
    panes: readonly Pane[],
    scrollX: number,
    scrollY: number,
  ): void {
    const { context } = this;
    const range = selectionRange(state.selection);
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
      context.fillRect(x, y, width, height);
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

    const active = state.selection.focus;
    const activeRect = this.activeCellRect(state.projection, layout, scrollX, scrollY, active);
    if (activeRect !== null) {
      context.strokeStyle = COLORS.activeBorder;
      context.lineWidth = 2;
      context.strokeRect(
        activeRect.x + 1,
        activeRect.y + 1,
        Math.max(0, activeRect.width - 2),
        Math.max(0, activeRect.height - 2),
      );
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
    const active = selection.focus;
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

      context.font = '12px system-ui, sans-serif';
      context.textBaseline = 'middle';
      context.textAlign = 'center';
      this.lastFont = '';

      if (frozenColPane !== undefined) {
        for (let col = frozenColPane.startCol; col <= frozenColPane.endCol; col += 1) {
          this.paintColumnHeader(
            projection,
            layout.headerWidth + projection.columnX(col),
            projection.columnWidth(col),
            col,
            active,
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
          const x = layout.headerWidth + layout.frozenWidth + projection.columnX(col) - scrollX;
          this.paintColumnHeader(projection, x, projection.columnWidth(col), col, active);
        }
        context.restore();
      }

      context.textAlign = 'right';
      if (frozenRowPane !== undefined) {
        for (let row = frozenRowPane.startRow; row <= frozenRowPane.endRow; row += 1) {
          const y = layout.headerHeight + projection.rowY(row);
          this.paintRowHeader(projection, y, projection.rowHeight(row), row, active);
        }
      }
      if (scrollingRowPane !== undefined) {
        for (let row = scrollingRowPane.startRow; row <= scrollingRowPane.endRow; row += 1) {
          const y = layout.headerHeight + layout.frozenHeight + projection.rowY(row) - scrollY;
          this.paintRowHeader(projection, y, projection.rowHeight(row), row, active);
        }
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

  private paintColumnHeader(
    projection: SheetProjection,
    x: number,
    width: number,
    col: number,
    active: GridCellAddress,
  ): void {
    const { context } = this;
    if (active.col === col) {
      context.fillStyle = COLORS.headerActiveBackground;
      context.fillRect(x, 0, width, GRID_METRICS.headerHeight);
      context.fillStyle = COLORS.headerActiveText;
    } else {
      context.fillStyle = COLORS.headerText;
    }
    context.textAlign = 'center';
    context.fillText(projection.columnLabel(col), x + width / 2, GRID_METRICS.headerHeight / 2);
  }

  private paintRowHeader(
    projection: SheetProjection,
    y: number,
    height: number,
    row: number,
    active: GridCellAddress,
  ): void {
    const { context } = this;
    if (active.row === row) {
      context.fillStyle = COLORS.headerActiveBackground;
      context.fillRect(0, y, GRID_METRICS.headerWidth, height);
      context.fillStyle = COLORS.headerActiveText;
    } else {
      context.fillStyle = COLORS.headerText;
    }
    context.textAlign = 'right';
    context.fillText(projection.rowLabel(row), GRID_METRICS.headerWidth - 8, y + height / 2);
  }

  destroy(): void {
    // The canvas is owned by React; releasing the context is enough.
    this.context.setTransform(1, 0, 0, 1, 0, 0);
    this.context.clearRect(0, 0, this.canvas.width, this.canvas.height);
  }
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

function fontFor(style: ProjectedStyle): string {
  const size = Math.max(1, Math.round((style.font.size * 96) / 72));
  const italic = style.font.italic ? 'italic ' : '';
  const bold = style.font.bold ? 'bold ' : '';
  const family =
    style.font.name === '' ? 'sans-serif' : `"${style.font.name}", system-ui, sans-serif`;
  return `${italic}${bold}${size}px ${family}`;
}

function horizontalFor(style: ProjectedStyle, value: CellValue): 'left' | 'center' | 'right' {
  const horizontal = style.alignment.horizontal;
  if (horizontal === 'left' || horizontal === 'center' || horizontal === 'right') return horizontal;
  if (horizontal === 'centerContinuous') return 'center';
  switch (value.type) {
    case 'number':
      return 'right';
    case 'boolean':
      return 'center';
    case 'error':
      return 'right';
    default:
      return 'left';
  }
}

function verticalBaseline(style: ProjectedStyle, rect: CellRect): number {
  const vertical = style.alignment.vertical;
  if (vertical === 'top') return rect.y + Math.min(rect.height - 2, 12);
  if (vertical === 'center') return rect.y + rect.height / 2;
  return rect.y + rect.height - Math.min(rect.height - 2, 6);
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
