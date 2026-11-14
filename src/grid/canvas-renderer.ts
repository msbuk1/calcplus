import type { SheetProjection } from '@view/types';
import {
  GRID_METRICS,
  type GridCellAddress,
  type GridRenderState,
  type RenderFrame,
} from './frame';

const COLORS = Object.freeze({
  background: '#ffffff',
  headerBackground: '#f5f5f5',
  headerText: '#444444',
  gridline: '#e1e1e1',
  text: '#202020',
  activeBorder: '#1a73e8',
  activeFill: 'rgba(26, 115, 232, 0.06)',
});

export class CanvasRenderFrame implements RenderFrame {
  private readonly canvas: HTMLCanvasElement;
  private readonly context: CanvasRenderingContext2D;
  private width = 0;
  private height = 0;
  private dpr = 1;

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
    context.setTransform(dpr, 0, 0, dpr, 0, 0);
    context.clearRect(0, 0, width, height);
    context.fillStyle = COLORS.background;
    context.fillRect(0, 0, width, height);

    const { projection, scrollX, scrollY, active } = state;
    const firstRow = projection.rowAt(scrollY);
    const lastRow = projection.rowAt(scrollY + height - GRID_METRICS.headerHeight);
    const firstCol = projection.columnAt(scrollX);
    const lastCol = projection.columnAt(scrollX + width - GRID_METRICS.headerWidth);

    context.save();
    context.beginPath();
    context.rect(
      GRID_METRICS.headerWidth,
      GRID_METRICS.headerHeight,
      width - GRID_METRICS.headerWidth,
      height - GRID_METRICS.headerHeight,
    );
    context.clip();

    context.strokeStyle = COLORS.gridline;
    context.lineWidth = 1;

    for (let row = firstRow; row <= lastRow + 1; row += 1) {
      const y = GRID_METRICS.headerHeight + projection.rowY(row) - scrollY + 0.5;
      context.beginPath();
      context.moveTo(GRID_METRICS.headerWidth, y);
      context.lineTo(width, y);
      context.stroke();
    }
    for (let col = firstCol; col <= lastCol + 1; col += 1) {
      const x = GRID_METRICS.headerWidth + projection.columnX(col) - scrollX + 0.5;
      context.beginPath();
      context.moveTo(x, GRID_METRICS.headerHeight);
      context.lineTo(x, height);
      context.stroke();
    }

    const activeX = GRID_METRICS.headerWidth + projection.columnX(active.col) - scrollX;
    const activeY = GRID_METRICS.headerHeight + projection.rowY(active.row) - scrollY;
    context.fillStyle = COLORS.activeFill;
    context.fillRect(
      activeX,
      activeY,
      projection.columnWidth(active.col),
      projection.rowHeight(active.row),
    );

    context.font = '13px system-ui, sans-serif';
    context.textBaseline = 'middle';
    context.textAlign = 'left';
    for (let row = firstRow; row <= lastRow; row += 1) {
      const y = GRID_METRICS.headerHeight + projection.rowY(row) - scrollY;
      const rowHeight = projection.rowHeight(row);
      for (let col = firstCol; col <= lastCol; col += 1) {
        const cell = projection.cellAt(row, col);
        if (cell === null || cell.displayText === '') continue;
        const x = GRID_METRICS.headerWidth + projection.columnX(col) - scrollX;
        const width = projection.columnWidth(col);
        context.save();
        context.beginPath();
        context.rect(x + 3, y + 1, width - 6, rowHeight - 2);
        context.clip();
        context.fillStyle = COLORS.text;
        context.fillText(cell.displayText, x + 5, y + rowHeight / 2);
        context.restore();
      }
    }

    context.strokeStyle = COLORS.activeBorder;
    context.lineWidth = 2;
    context.strokeRect(
      activeX + 0.5,
      activeY + 0.5,
      projection.columnWidth(active.col) - 1,
      projection.rowHeight(active.row) - 1,
    );
    context.restore();

    this.drawHeaders(state, firstRow, lastRow, firstCol, lastCol);
  }

  private drawHeaders(
    state: GridRenderState,
    firstRow: number,
    lastRow: number,
    firstCol: number,
    lastCol: number,
  ): void {
    const { context, width, height } = this;
    const { projection, scrollX, scrollY, active } = state;
    context.fillStyle = COLORS.headerBackground;
    context.fillRect(0, 0, width, GRID_METRICS.headerHeight);
    context.fillRect(0, 0, GRID_METRICS.headerWidth, height);
    context.strokeStyle = COLORS.gridline;
    context.lineWidth = 1;
    context.beginPath();
    context.moveTo(0, GRID_METRICS.headerHeight + 0.5);
    context.lineTo(width, GRID_METRICS.headerHeight + 0.5);
    context.moveTo(GRID_METRICS.headerWidth + 0.5, 0);
    context.lineTo(GRID_METRICS.headerWidth + 0.5, height);
    context.stroke();

    context.fillStyle = COLORS.headerText;
    context.font = '12px system-ui, sans-serif';
    context.textBaseline = 'middle';
    context.textAlign = 'center';
    for (let col = firstCol; col <= lastCol; col += 1) {
      const x = GRID_METRICS.headerWidth + projection.columnX(col) - scrollX;
      const width = projection.columnWidth(col);
      if (col === active.col) {
        context.fillStyle = COLORS.activeBorder;
        context.fillRect(x, 0, width, GRID_METRICS.headerHeight);
        context.fillStyle = '#ffffff';
      } else {
        context.fillStyle = COLORS.headerText;
      }
      context.fillText(projection.columnLabel(col), x + width / 2, GRID_METRICS.headerHeight / 2);
    }

    context.textAlign = 'right';
    for (let row = firstRow; row <= lastRow; row += 1) {
      const y = GRID_METRICS.headerHeight + projection.rowY(row) - scrollY;
      const rowHeight = projection.rowHeight(row);
      if (row === active.row) {
        context.fillStyle = COLORS.activeBorder;
        context.fillRect(0, y, GRID_METRICS.headerWidth, rowHeight);
        context.fillStyle = '#ffffff';
      } else {
        context.fillStyle = COLORS.headerText;
      }
      context.fillText(String(row + 1), GRID_METRICS.headerWidth - 8, y + rowHeight / 2);
    }
  }

  destroy(): void {
    // The canvas is owned by React; releasing the context is enough.
    this.context.setTransform(1, 0, 0, 1, 0, 0);
    this.context.clearRect(0, 0, this.canvas.width, this.canvas.height);
  }
}

export function cellAtPoint(
  projection: SheetProjection,
  pointX: number,
  pointY: number,
  scrollX: number,
  scrollY: number,
): GridCellAddress | null {
  if (pointX < GRID_METRICS.headerWidth || pointY < GRID_METRICS.headerHeight) return null;
  const col = projection.columnAt(pointX - GRID_METRICS.headerWidth + scrollX);
  const row = projection.rowAt(pointY - GRID_METRICS.headerHeight + scrollY);
  return { row, col };
}
