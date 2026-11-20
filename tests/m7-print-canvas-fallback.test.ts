import { describe, expect, it } from 'vitest';
import {
  createCanvasRasterizer,
  fallbackPixels,
  paintPage,
  rasterizePage,
  type FallbackCanvas,
} from '../src/print/canvas-fallback';
import type { PrintItem } from '../src/print/types';
import { ONE_PIXEL_PNG, box, makePage, textItem } from './helpers/print-pdf';

/**
 * M7 lane 3 — the 300 DPI canvas fallback (print-pdf §2, §7).
 *
 * The fake canvas records the 2D calls and returns a real 1x1 PNG, so the
 * tests assert the display list is painted and that the raster is sized for
 * 300 DPI without needing a browser.
 */

interface Gradient {
  readonly addColorStop: (offset: number, color: string) => void;
}

class RecordingContext {
  readonly ops: string[] = [];
  readonly gradients: string[] = [];
  fillStyle: unknown = '';
  strokeStyle: unknown = '';
  lineWidth = 1;
  globalAlpha = 1;
  font = '';
  textAlign = 'left';
  textBaseline = 'alphabetic';
  readonly text: string[] = [];
  dashes: number[][] = [];

  save(): void {
    this.ops.push('save');
  }
  restore(): void {
    this.ops.push('restore');
  }
  scale(x: number, y: number): void {
    this.ops.push(`scale:${x}:${y}`);
  }
  fillRect(x: number, y: number, w: number, h: number): void {
    this.ops.push(`fillRect:${this.fillStyle as string}:${x},${y},${w},${h}`);
  }
  strokeRect(x: number, y: number, w: number, h: number): void {
    this.ops.push(`strokeRect:${x},${y},${w},${h}`);
  }
  beginPath(): void {
    this.ops.push('beginPath');
  }
  closePath(): void {
    this.ops.push('closePath');
  }
  rect(x: number, y: number, w: number, h: number): void {
    this.ops.push(`rect:${x},${y},${w},${h}`);
  }
  clip(): void {
    this.ops.push('clip');
  }
  moveTo(x: number, y: number): void {
    this.ops.push(`moveTo:${x},${y}`);
  }
  lineTo(x: number, y: number): void {
    this.ops.push(`lineTo:${x},${y}`);
  }
  arc(x: number, y: number, r: number): void {
    this.ops.push(`arc:${x},${y},${r}`);
  }
  stroke(): void {
    this.ops.push('stroke');
  }
  fill(): void {
    this.ops.push('fill');
  }
  setLineDash(dash: number[]): void {
    this.dashes.push(dash);
  }
  translate(x: number, y: number): void {
    this.ops.push(`translate:${x},${y}`);
  }
  rotate(angle: number): void {
    this.ops.push(`rotate:${angle}`);
  }
  fillText(text: string, x: number, y: number): void {
    this.text.push(text);
    this.ops.push(`fillText:${text}:${x},${y}`);
  }
  createLinearGradient(x1: number, y1: number, x2: number, y2: number): Gradient {
    this.gradients.push(`${x1},${y1},${x2},${y2}`);
    return { addColorStop: () => undefined };
  }
}

function fakeCanvas(
  width: number,
  height: number,
): {
  canvas: FallbackCanvas;
  context: RecordingContext;
  size: { width: number; height: number };
} {
  const context = new RecordingContext();
  const size = { width, height };
  return {
    size,
    context,
    canvas: {
      width,
      height,
      getContext: () => context as unknown as CanvasRenderingContext2D,
      toPng: async () => ONE_PIXEL_PNG,
    },
  };
}

const items: readonly PrintItem[] = [
  { kind: 'fill', box: box(50, 60, 200, 40), pattern: 'solid', color: '#FFEEAA', background: null },
  {
    kind: 'fill',
    box: box(50, 110, 200, 40),
    pattern: 'darkGrid',
    color: '#444444',
    background: '#FFFFFF',
  },
  { kind: 'line', x1: 50, y1: 160, x2: 250, y2: 160, color: '#000000', width: 1, dash: [3, 2] },
  { kind: 'line', x1: 50, y1: 170, x2: 250, y2: 170, color: '#808080', width: 0.5, dash: null },
  {
    kind: 'dataBar',
    box: box(60, 180, 120, 14),
    color: '#638EC6',
    gradient: { from: '#DDEBF7', to: '#638EC6' },
    borderColor: '#4A7EBB',
    axisX: 120,
    axisColor: '#000000',
  },
  {
    kind: 'icon',
    box: box(60, 200, 12, 12),
    family: 'quarter',
    color: '#63BE7B',
    iconIndex: 1,
    iconCount: 3,
  },
  textItem({ text: 'Raster me', x: 60, y: 260, rotation: 0 }),
  textItem({ text: 'Rotated', x: 120, y: 300, anchor: 'middle', rotation: 90 }),
];

describe('fallbackPixels', () => {
  it('sizes A4 at 300 DPI to 2480 x 3508 pixels', () => {
    expect(fallbackPixels(595.28, 841.89, 300)).toEqual({ width: 2480, height: 3508 });
  });
});

describe('paintPage', () => {
  it('paints the whole display list in plan order on the canvas', () => {
    const page = makePage(0, items);
    const { canvas, context } = fakeCanvas(2480, 3508);
    const ctx = canvas.getContext('2d');
    expect(ctx).not.toBeNull();
    paintPage(ctx!, page, 300);

    expect(context.ops[0]).toBe('save');
    expect(context.ops).toContain('scale:4.166666666666667:4.166666666666667');
    expect(context.ops.some((op) => op.startsWith('fillRect:#FFEEAA'))).toBe(true);
    // The pattern tile paints its grid lines inside the clipped cell.
    expect(context.dashes).toContainEqual([]);
    expect(context.dashes).toContainEqual([3, 2]);
    expect(context.gradients).toEqual(['60,180,180,180']);
    expect(context.ops.some((op) => op.startsWith('arc:'))).toBe(true);
    expect(context.text).toEqual(['Raster me', 'Rotated']);
    expect(context.ops).toContain('rotate:1.5707963267948966');
    expect(context.ops.at(-1)).toBe('restore');
  });
});

describe('rasterizePage', () => {
  it('returns a 300 DPI PNG for the page and echoes the reason', async () => {
    const page = makePage(2, items);
    const created = fakeCanvas(0, 0);
    const reason = { code: 'missing-glyph' as const, detail: 'U+6C49' };
    const raster = await rasterizePage(page, reason, {
      createCanvas: (width, height) => {
        created.size.width = width;
        created.size.height = height;
        return created.canvas;
      },
      ensureFonts: async () => undefined,
    });

    expect(created.size).toEqual({ width: 2480, height: 3508 });
    expect(raster.pageIndex).toBe(2);
    expect(raster.dpi).toBe(300);
    expect(raster.width).toBe(2480);
    expect(raster.height).toBe(3508);
    expect(raster.png).toBe(ONE_PIXEL_PNG);
    expect(raster.reason).toEqual(reason);
  });

  it('loads the bundled faces before the first page is drawn', async () => {
    let fontLoads = 0;
    const rasterize = createCanvasRasterizer({
      createCanvas: (width, height) => fakeCanvas(width, height).canvas,
      ensureFonts: async () => {
        fontLoads += 1;
      },
    });
    await rasterize(makePage(0, [textItem()]), { code: 'unknown-family', detail: 'Wingdings' });
    expect(fontLoads).toBe(1);
  });

  it('throws when no canvas is available instead of producing a blank page', async () => {
    await expect(
      rasterizePage(
        makePage(0, [textItem()]),
        { code: 'missing-glyph', detail: 'U+6C49' },
        {
          createCanvas: () => null,
          ensureFonts: async () => undefined,
        },
      ),
    ).rejects.toThrow(/canvas/);
  });
});
