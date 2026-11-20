import type { PrintDataBarItem, PrintPage, PrintTextItem } from './types';
import { iconShapes, patternTileOps, PATTERN_TILE, type IconShape } from './paint-shapes';
import { canvasFamilyFor, PDF_FACE_ASSETS } from './pdf-fonts';

/**
 * The 300 DPI canvas fallback (print-pdf §2, §7; research/26 §4).
 *
 * It paints the **same** `PrintPlan` page the PDFKit renderer would paint, so a
 * fallback page is rasterised content, never different content. The only
 * reason a page reaches this module is a font or script the vector path cannot
 * embed (see `preflightPrintPlan` in `pdf-renderer.ts`).
 *
 * Why 300 DPI: measured at 300 DPI a 20-page report is about 16 MB against
 * 57 KB vector, and PNG keeps line art crisp — JPEG barely saves bytes and
 * rings around text, so the fallback emits PNG (research/26 §5).
 */

export const FALLBACK_DPI = 300;

/** Why a page left the vector path. Documented in `pdf-renderer.ts`. */
export interface FallbackReason {
  readonly code: 'unknown-family' | 'missing-glyph' | 'rtl-script' | 'font-load';
  /** Human-readable detail: the family, the code point or the error. */
  readonly detail: string;
}

/** A rasterised page, ready to be placed as a full-page image. */
export interface RasterPage {
  readonly pageIndex: number;
  readonly dpi: number;
  /** Pixel width and height at `dpi`. */
  readonly width: number;
  readonly height: number;
  readonly png: Uint8Array;
  readonly reason: FallbackReason;
}

/** The small surface both the browser and tests supply. */
export interface FallbackCanvas {
  readonly width: number;
  readonly height: number;
  getContext(type: '2d'): CanvasRenderingContext2D | null;
  toPng(): Promise<Uint8Array>;
}

export type CanvasFactory = (width: number, height: number) => FallbackCanvas | null;

export interface CanvasFallbackOptions {
  readonly dpi?: number;
  readonly createCanvas?: CanvasFactory;
  /** Loads the bundled faces into the document before the first page is drawn. */
  readonly ensureFonts?: () => Promise<void>;
}

/** The pixel size of a page at a DPI, rounded to whole pixels. */
export function fallbackPixels(
  widthPoints: number,
  heightPoints: number,
  dpi: number = FALLBACK_DPI,
): { readonly width: number; readonly height: number } {
  return {
    width: Math.max(1, Math.round((widthPoints * dpi) / 72)),
    height: Math.max(1, Math.round((heightPoints * dpi) / 72)),
  };
}

/**
 * The browser canvas factory. Tests pass their own; without a DOM the default
 * throws rather than silently producing a blank page.
 */
function defaultCanvasFactory(width: number, height: number): FallbackCanvas | null {
  if (typeof document === 'undefined') return null;
  const canvas = document.createElement('canvas');
  canvas.width = width;
  canvas.height = height;
  return {
    width,
    height,
    getContext: (type) => canvas.getContext(type),
    toPng: () =>
      new Promise((resolve, reject) => {
        canvas.toBlob((blob) => {
          if (blob === null) {
            reject(new Error('The canvas produced no PNG data.'));
            return;
          }
          void blob.arrayBuffer().then(
            (buffer) => resolve(new Uint8Array(buffer)),
            (error: unknown) => reject(error),
          );
        }, 'image/png');
      }),
  };
}

let fontsPromise: Promise<void> | null = null;

/**
 * Register the bundled OFL faces so the fallback's text uses the same type
 * design as the vector pages. A face that fails to load is skipped: the
 * browser's own fallback chain still paints readable text, which is the
 * fallback's whole job.
 */
export function ensureFallbackFonts(): Promise<void> {
  if (fontsPromise !== null) return fontsPromise;
  fontsPromise = (async () => {
    if (typeof document === 'undefined' || typeof FontFace === 'undefined') return;
    await Promise.all(
      PDF_FACE_ASSETS.map(async (asset) => {
        try {
          const face = new FontFace(asset.cssFamily, `url(${asset.url})`, {
            weight: String(asset.fontWeight),
            style: asset.fontStyle,
            display: 'block',
          });
          const loaded = await face.load();
          document.fonts.add(loaded);
        } catch {
          // A missing face is not fatal: the CSS stack falls through.
        }
      }),
    );
  })();
  return fontsPromise;
}

/** A rasteriser bound to one export, with the fonts warm before page one. */
export function createCanvasRasterizer(
  options: CanvasFallbackOptions = {},
): (page: PrintPage, reason: FallbackReason) => Promise<RasterPage> {
  return async (page, reason) => {
    await (options.ensureFonts ?? ensureFallbackFonts)();
    return paintPageToRaster(page, reason, options);
  };
}

/** Rasterise one plan page to a PNG at the fallback DPI. */
export async function rasterizePage(
  page: PrintPage,
  reason: FallbackReason,
  options: CanvasFallbackOptions = {},
): Promise<RasterPage> {
  await (options.ensureFonts ?? ensureFallbackFonts)();
  return paintPageToRaster(page, reason, options);
}

async function paintPageToRaster(
  page: PrintPage,
  reason: FallbackReason,
  options: CanvasFallbackOptions,
): Promise<RasterPage> {
  const dpi = options.dpi ?? FALLBACK_DPI;
  const { width, height } = fallbackPixels(page.box.width, page.box.height, dpi);
  const factory = options.createCanvas ?? defaultCanvasFactory;
  const canvas = factory(width, height);
  if (canvas === null) {
    throw new Error('The 300 DPI fallback needs a canvas; none is available in this environment.');
  }
  const context = canvas.getContext('2d');
  if (context === null) {
    throw new Error('The 300 DPI fallback could not get a 2D context.');
  }
  paintPage(context, page, dpi);
  const png = await canvas.toPng();
  return { pageIndex: page.index, dpi, width, height, png, reason };
}

/** Paint the display list in plan order, scaled so one unit is one point. */
export function paintPage(context: CanvasRenderingContext2D, page: PrintPage, dpi: number): void {
  const scale = dpi / 72;
  context.save();
  context.scale(scale, scale);
  context.fillStyle = '#ffffff';
  context.fillRect(0, 0, page.box.width, page.box.height);
  context.beginPath();
  context.rect(0, 0, page.box.width, page.box.height);
  context.clip();
  for (const item of page.items) {
    switch (item.kind) {
      case 'fill':
        paintFill(context, item.box, item.pattern, item.color, item.background);
        break;
      case 'line':
        context.save();
        context.strokeStyle = item.color;
        context.lineWidth = Math.max(0.1, item.width);
        context.setLineDash(item.dash === null ? [] : [...item.dash]);
        context.beginPath();
        context.moveTo(item.x1, item.y1);
        context.lineTo(item.x2, item.y2);
        context.stroke();
        context.restore();
        break;
      case 'text':
        paintText(context, item);
        break;
      case 'dataBar':
        paintDataBar(context, item);
        break;
      case 'icon':
        for (const shape of iconShapes(item)) paintIconShape(context, shape);
        break;
      case 'image':
        // v1 reserves the anchor and paints nothing (print-pdf §5).
        break;
    }
  }
  context.restore();
}

function paintFill(
  context: CanvasRenderingContext2D,
  box: { readonly x: number; readonly y: number; readonly width: number; readonly height: number },
  pattern: string,
  color: string | null,
  background: string | null,
): void {
  if (pattern === 'none') return;
  if ((pattern === 'solid' || pattern === '') && color === null) return;
  if (pattern === 'solid' || pattern === '') {
    context.fillStyle = color as string;
    context.fillRect(box.x, box.y, box.width, box.height);
    return;
  }
  if (color === null) {
    if (background === null) return;
    context.fillStyle = background;
    context.fillRect(box.x, box.y, box.width, box.height);
    return;
  }
  context.save();
  context.beginPath();
  context.rect(box.x, box.y, box.width, box.height);
  context.clip();
  context.fillStyle = background ?? '#ffffff';
  context.fillRect(box.x, box.y, box.width, box.height);
  context.fillStyle = color;
  context.strokeStyle = color;
  context.lineWidth = 1;
  for (let y = box.y; y < box.y + box.height; y += PATTERN_TILE) {
    for (let x = box.x; x < box.x + box.width; x += PATTERN_TILE) {
      for (const op of patternTileOps(pattern)) {
        context.globalAlpha = op.opacity;
        if (op.kind === 'fill') {
          context.fillRect(x + op.x, y + op.y, op.width, op.height);
        } else {
          context.beginPath();
          context.moveTo(x + op.x1, y + op.y1);
          context.lineTo(x + op.x2, y + op.y2);
          context.stroke();
        }
      }
    }
  }
  context.globalAlpha = 1;
  context.restore();
}

function paintText(context: CanvasRenderingContext2D, item: PrintTextItem): void {
  if (item.text === '') return;
  context.save();
  context.beginPath();
  context.rect(item.clip.x, item.clip.y, item.clip.width, item.clip.height);
  context.clip();
  context.translate(item.x, item.y);
  if (item.rotation !== 0) context.rotate((item.rotation * Math.PI) / 180);
  const style = item.font.italic ? 'italic ' : '';
  const weight = item.font.bold ? '700' : '400';
  context.font = `${style}${weight} ${item.font.size}px ${canvasFamilyFor(item.font.family)}`;
  context.fillStyle = item.font.color;
  context.textBaseline = 'alphabetic';
  context.textAlign =
    item.anchor === 'start' ? 'left' : item.anchor === 'middle' ? 'center' : 'right';
  context.fillText(item.text, 0, 0);
  context.restore();
}

function paintDataBar(context: CanvasRenderingContext2D, item: PrintDataBarItem): void {
  const { box } = item;
  if (item.gradient === null) {
    context.fillStyle = item.color;
  } else {
    const gradient = context.createLinearGradient(box.x, box.y, box.x + box.width, box.y);
    gradient.addColorStop(0, item.gradient.from);
    gradient.addColorStop(1, item.gradient.to);
    context.fillStyle = gradient;
  }
  context.fillRect(box.x, box.y, box.width, box.height);
  if (item.borderColor !== null) {
    context.save();
    context.strokeStyle = item.borderColor;
    context.lineWidth = 1;
    context.strokeRect(box.x, box.y, box.width, box.height);
    context.restore();
  }
  if (item.axisX !== null) {
    context.save();
    context.strokeStyle = item.axisColor;
    context.lineWidth = 1;
    context.beginPath();
    context.moveTo(item.axisX, box.y);
    context.lineTo(item.axisX, box.y + box.height);
    context.stroke();
    context.restore();
  }
}

function paintIconShape(context: CanvasRenderingContext2D, shape: IconShape): void {
  context.save();
  context.beginPath();
  for (const segment of shape.segments) {
    switch (segment.t) {
      case 'move':
        context.moveTo(segment.x, segment.y);
        break;
      case 'line':
        context.lineTo(segment.x, segment.y);
        break;
      case 'close':
        context.closePath();
        break;
      case 'circle':
        context.moveTo(segment.x + segment.r, segment.y);
        context.arc(segment.x, segment.y, segment.r, 0, Math.PI * 2);
        context.closePath();
        break;
    }
  }
  if (shape.paint === 'fill') {
    context.fillStyle = shape.color;
    context.fill();
  } else {
    context.strokeStyle = shape.color;
    context.lineWidth = Math.max(0.1, shape.lineWidth);
    context.stroke();
  }
  context.restore();
}
