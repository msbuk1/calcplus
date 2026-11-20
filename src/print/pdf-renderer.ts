import * as fontkit from 'fontkit';
import type { FontkitFont } from 'fontkit';
import type { FallbackReason, RasterPage } from './canvas-fallback';
import {
  faceAssetFor,
  resolvePdfFamily,
  subsetName,
  supplementAssetFor,
  type PdfFaceAsset,
} from './pdf-fonts';
import { iconShapes, patternTileOps, PATTERN_TILE, type IconShape } from './paint-shapes';
import type { PrintFont, PrintItem, PrintPage, PrintPlan, PrintTextItem } from './types';

/**
 * The PDFKit vector renderer (print-pdf §2, §4; research/26).
 *
 * One plan page becomes exactly one PDF page with the plan's page box in
 * points. The display list paints in the plan's array order, so the PDF matches
 * the SVG preview.
 *
 * Fonts are embedded subsetted and every subset is renamed `CALP+<family>`
 * through the `embedProgram` seam the spike measured, because Selawik and
 * Liberation carry Reserved Font Names (licensing §3.6). PDFKit's browser build
 * has no standard-font data, so the document is constructed with `font: null`
 * and an embedded face is selected before any text is drawn.
 *
 * Pages whose text cannot embed safely are rasterised at 300 DPI from the same
 * page by `canvas-fallback.ts` and placed as a full-page PNG. The decision is
 * made before rendering (`preflightPrintPlan`) because a missing glyph is
 * invisible at render time: fontkit maps it to glyph 0 and PDFKit draws an
 * empty `.notdef` without raising (research/26 §4.1).
 */

/** Why a page left the vector path, in the words the export report shows. */
export interface PageDecision {
  readonly pageIndex: number;
  readonly fallback: boolean;
  readonly reason: FallbackReason | null;
  /** Runs that used the Carlito supplement because the primary face lacked a glyph. */
  readonly substitutedRuns: number;
}

export interface PrintPreflight {
  readonly decisions: readonly PageDecision[];
  readonly fallbackPages: readonly number[];
  /** Assets the plan needs, keyed by asset id, with their parsed faces. */
  readonly faces: ReadonlyMap<string, LoadedFace>;
}

/** One font file the plan needs. `font` is null when the file failed to parse. */
export interface LoadedFace {
  readonly asset: PdfFaceAsset;
  readonly bytes: Uint8Array;
  readonly font: FontkitFont | null;
  readonly error: string | null;
}

export type FontBytesLoader = (url: string) => Promise<Uint8Array>;

export interface PdfRenderOptions {
  /** Loads a bundled `.ttf`; the default fetches the asset URL. */
  readonly loadFont?: FontBytesLoader;
  /**
   * Rasterises one page for the fallback. Required as soon as a page needs it:
   * the renderer throws rather than emit a page it cannot draw correctly.
   */
  readonly rasterize?: (page: PrintPage, reason: FallbackReason) => Promise<RasterPage>;
}

export interface PdfRenderResult {
  readonly bytes: Uint8Array;
  readonly pageCount: number;
  readonly fallbackPages: readonly number[];
  readonly decisions: readonly PageDecision[];
  /** The `CALP+<family>` names written into the PDF. */
  readonly embeddedFonts: readonly string[];
}

/** The default loader: fetch the bundled asset (browser) by URL. */
export async function fetchFontBytes(url: string): Promise<Uint8Array> {
  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(`Could not load the print font (${response.status}): ${url}`);
  }
  return new Uint8Array(await response.arrayBuffer());
}

/**
 * Decide, page by page, whether the vector path can draw the text correctly.
 * Detection follows the spike's table (research/26 §7.2):
 *
 * | Check | Trigger | Result |
 * |---|---|---|
 * | family | no metric clone in the bundle | page -> raster |
 * | glyph coverage | a code point the resolved face (and its Carlito supplement) lacks | page -> raster |
 * | script mix | any right-to-left script character | page -> raster (v1 is conservative; PDFKit does no cross-run bidi) |
 * | font load | the file failed to parse or embed | page -> raster |
 *
 * Coverage is per page, so the fallback stays limited to the pages that need
 * it (print-pdf §6).
 */
export async function preflightPrintPlan(
  plan: PrintPlan,
  loadFont: FontBytesLoader = fetchFontBytes,
): Promise<PrintPreflight> {
  const faces = new Map<string, LoadedFace>();
  const load = async (asset: PdfFaceAsset): Promise<LoadedFace> => {
    const existing = faces.get(asset.id);
    if (existing !== undefined) return existing;
    let loaded: LoadedFace;
    try {
      const bytes = await loadFont(asset.url);
      loaded = { asset, bytes, font: safeCreate(bytes), error: null };
    } catch (error) {
      loaded = {
        asset,
        bytes: new Uint8Array(0),
        font: null,
        error: error instanceof Error ? error.message : String(error),
      };
    }
    faces.set(asset.id, loaded);
    return loaded;
  };

  // Pass 1: every face the plan names. An unknown family is decided without a
  // font and goes straight to the page's reason.
  for (const page of plan.pages) {
    for (const item of page.items) {
      if (item.kind !== 'text') continue;
      const asset = primaryAsset(item.font);
      if (asset !== undefined) await load(asset);
    }
  }

  // Pass 2: the Carlito supplement for faces that miss a glyph.
  const decisions: PageDecision[] = [];
  for (const page of plan.pages) {
    let reason: FallbackReason | null = null;
    let substitutedRuns = 0;
    for (const item of page.items) {
      if (item.kind !== 'text' || item.text === '') continue;
      if (reason !== null) break;
      const family = resolvePdfFamily(item.font.family);
      if (family === null) {
        reason = { code: 'unknown-family', detail: item.font.family };
        break;
      }
      if (RTL_SCRIPT.test(item.text)) {
        reason = {
          code: 'rtl-script',
          detail: 'right-to-left text needs Unicode bidi, which the vector path does not apply',
        };
        break;
      }
      const primary = primaryAsset(item.font);
      if (primary === undefined) {
        reason = { code: 'unknown-family', detail: item.font.family };
        break;
      }
      const primaryFace = await load(primary);
      if (primaryFace.font === null) {
        reason = {
          code: 'font-load',
          detail: `${primaryFace.asset.id}: ${primaryFace.error ?? 'could not be parsed'}`,
        };
        break;
      }
      const missing = missingCodePoints(primaryFace.font, item.text);
      if (missing.length === 0) continue;
      const supplement = supplementAssetFor(item.font.bold, item.font.italic);
      if (supplement === undefined) {
        reason = { code: 'missing-glyph', detail: `U+${hex(missing[0])} needs the fallback` };
        break;
      }
      const supplementFace = await load(supplement);
      if (
        supplementFace.font === null ||
        missingCodePoints(supplementFace.font, item.text).length > 0
      ) {
        reason = {
          code: 'missing-glyph',
          detail: `U+${hex(missing[0])} is not in ${primary.subsetFamily} or Carlito`,
        };
        break;
      }
      substitutedRuns += 1;
    }
    decisions.push({
      pageIndex: page.index,
      fallback: reason !== null,
      reason,
      substitutedRuns,
    });
  }

  return {
    decisions,
    fallbackPages: decisions.filter((decision) => decision.fallback).map((d) => d.pageIndex),
    faces,
  };
}

/**
 * Render the plan to a vector PDF. Pages that the preflight routes to the
 * fallback are rasterised at 300 DPI and inserted as a full-page PNG; every
 * other page paints the display list.
 */
export async function renderPdf(
  plan: PrintPlan,
  options: PdfRenderOptions = {},
): Promise<PdfRenderResult> {
  const preflight = await preflightPrintPlan(plan, options.loadFont ?? fetchFontBytes);
  if (preflight.fallbackPages.length > 0 && options.rasterize === undefined) {
    throw new Error(
      `Page ${preflight.fallbackPages[0] + 1} needs the 300 DPI fallback, but no rasteriser was supplied.`,
    );
  }

  const PDFDocument = await loadPdfKit();
  const first = plan.pages[0];
  const doc = new PDFDocument({
    size: first === undefined ? [595.28, 841.89] : [first.box.width, first.box.height],
    margin: 0,
    font: null,
    autoFirstPage: false,
    info: { Title: plan.sheetName, Creator: 'calcplus' },
  } as unknown as PDFKit.PDFDocumentOptions) as PDFKit.PDFDocument;

  const chunks: Uint8Array[] = [];
  doc.on('data', (chunk: Uint8Array) => chunks.push(chunk));
  const finished = new Promise<void>((resolve, reject) => {
    doc.on('end', resolve);
    doc.on('error', reject);
  });

  // Register only the faces a vector page actually paints, and rename every
  // subset before any text. PDFKit embeds every selected face at `doc.end()`,
  // so registering unused faces would add bytes nothing references.
  const usedFaces = new Set<string>();
  for (const page of plan.pages) {
    if (preflight.decisions[page.index]?.fallback === true) continue;
    for (const item of page.items) {
      if (item.kind !== 'text' || item.text === '') continue;
      const face = faceForRun(item.font, item.text, preflight.faces);
      if (face !== null) usedFaces.add(face.asset.id);
    }
  }
  for (const id of usedFaces) {
    const face = preflight.faces.get(id);
    if (face !== undefined && face.font !== null) registerFace(doc, face);
  }

  for (const page of plan.pages) {
    const decision = preflight.decisions[page.index];
    doc.addPage({ size: [page.box.width, page.box.height], margin: 0 });
    if (decision?.fallback === true) {
      const raster = await options.rasterize!(page, decision.reason!);
      doc.image(raster.png as unknown as Buffer, 0, 0, {
        width: page.box.width,
        height: page.box.height,
      });
      continue;
    }
    paintPage(doc, page, preflight.faces);
  }

  doc.end();
  await finished;

  return {
    bytes: concatChunks(chunks),
    pageCount: plan.pages.length,
    fallbackPages: preflight.fallbackPages,
    decisions: preflight.decisions,
    embeddedFonts: [...usedFaces].sort().map((id) => {
      const face = preflight.faces.get(id);
      return face === undefined ? id : subsetName(face.asset);
    }),
  };
}

/** 300 DPI in the renderer's vocabulary, re-exported for the export flow. */
export { FALLBACK_DPI } from './canvas-fallback';

// ---------------------------------------------------------------------------
// Preflight helpers
// ---------------------------------------------------------------------------

function primaryAsset(font: PrintFont): PdfFaceAsset | undefined {
  const family = resolvePdfFamily(font.family);
  if (family === null) return undefined;
  return faceAssetFor(family, font.bold, font.italic);
}

function safeCreate(bytes: Uint8Array): FontkitFont | null {
  try {
    return fontkit.create(bytes);
  } catch {
    return null;
  }
}

/** Right-to-left scripts (Arabic, Hebrew, Syriac, Thaana, NKo, Samaritan...). */
const RTL_SCRIPT =
  /[\u0590-\u05FF\u0600-\u06FF\u0700-\u074F\u0750-\u077F\u0780-\u07BF\u07C0-\u07FF\u0800-\u083F\u0840-\u085F\u0860-\u086F\u0870-\u089F\u08A0-\u08FF\uFB1D-\uFB4F\uFB50-\uFDFF\uFE70-\uFEFF\u{10800}-\u{10FFF}]/u;

function missingCodePoints(font: FontkitFont, text: string): number[] {
  const missing: number[] = [];
  for (const character of text) {
    const codePoint = character.codePointAt(0);
    if (codePoint === undefined || codePoint <= 0x20) continue;
    if (!font.hasGlyphForCodePoint(codePoint)) missing.push(codePoint);
  }
  return missing;
}

function hex(codePoint: number): string {
  return codePoint.toString(16).toUpperCase().padStart(4, '0');
}

// ---------------------------------------------------------------------------
// PDFKit loading and the subset-rename seam
// ---------------------------------------------------------------------------

type PdfKitConstructor = new (options?: PDFKit.PDFDocumentOptions) => PDFKit.PDFDocument;

async function loadPdfKit(): Promise<PdfKitConstructor> {
  const module = (await import('pdfkit')) as unknown as {
    default?: PdfKitConstructor;
    PDFDocument?: PdfKitConstructor;
  };
  const constructor = module.default ?? module.PDFDocument;
  if (constructor === undefined) throw new Error('PDFKit did not provide a PDFDocument export.');
  return constructor;
}

/**
 * Rename the embedded subset to `CALP+<family>` in all three font descriptors
 * (research/26 §3.3). PDFKit writes the descriptor before `embedProgram`
 * returns, so the descriptor's `FontName` is rewritten while the `ref` ends;
 * the returned name replaces both `BaseFont` entries.
 */
function registerFace(doc: PDFKit.PDFDocument, face: LoadedFace): void {
  doc.registerFont(face.asset.id, face.bytes);
  doc.font(face.asset.id);
  const font = internalFont(doc);
  if (font === null || font.patched === true) return;
  font.patched = true;
  const renamed = subsetName(face.asset);
  const original = font.embedProgram.bind(font);
  font.embedProgram = (subset: unknown, complete: boolean) => {
    const originalRef = doc.ref.bind(doc);
    doc.ref = (data?: Record<string, unknown>) => {
      const reference = originalRef(data ?? {});
      if (data?.Type === 'FontDescriptor') {
        const end = reference.end.bind(reference);
        reference.end = (chunk?: Uint8Array) => {
          (reference.data as Record<string, unknown>).FontName = renamed;
          end(chunk);
        };
      }
      return reference;
    };
    try {
      const result = original(subset, complete) as {
        readonly descriptor: PDFKit.PDFKitReference;
        readonly name: string;
        readonly isCFF: boolean;
      };
      if (!complete) {
        (result.descriptor.data as Record<string, unknown>).FontName = renamed;
        return { ...result, name: renamed };
      }
      return result;
    } finally {
      doc.ref = originalRef;
    }
  };
}

interface EmbeddableFont {
  embedProgram(subset: unknown, complete: boolean): unknown;
  patched?: boolean;
}

function internalFont(doc: PDFKit.PDFDocument): EmbeddableFont | null {
  const internal = doc as unknown as { _font: EmbeddableFont | null };
  return internal._font ?? null;
}

// ---------------------------------------------------------------------------
// Painting
// ---------------------------------------------------------------------------

function paintPage(
  doc: PDFKit.PDFDocument,
  page: PrintPage,
  faces: ReadonlyMap<string, LoadedFace>,
): void {
  doc.save();
  doc.rect(0, 0, page.box.width, page.box.height).clip();
  for (const item of page.items) {
    switch (item.kind) {
      case 'fill':
        paintFill(doc, item);
        break;
      case 'line':
        paintLine(doc, item);
        break;
      case 'text':
        paintText(doc, item, faces);
        break;
      case 'dataBar':
        paintDataBar(doc, item);
        break;
      case 'icon':
        for (const shape of iconShapes(item)) paintIconShape(doc, shape);
        break;
      case 'image':
        // v1 reserves the anchor and paints nothing (print-pdf §5).
        break;
    }
  }
  doc.restore();
}

type FillItem = Extract<PrintItem, { kind: 'fill' }>;
type LineItem = Extract<PrintItem, { kind: 'line' }>;
type DataBarItem = Extract<PrintItem, { kind: 'dataBar' }>;

function paintFill(doc: PDFKit.PDFDocument, item: FillItem): void {
  if (item.pattern === 'none') return;
  if (item.pattern === 'solid' || item.pattern === '') {
    if (item.color !== null)
      doc.rect(item.box.x, item.box.y, item.box.width, item.box.height).fill(item.color);
    return;
  }
  if (item.color === null) {
    if (item.background !== null) {
      doc.rect(item.box.x, item.box.y, item.box.width, item.box.height).fill(item.background);
    }
    return;
  }
  doc.save();
  doc.rect(item.box.x, item.box.y, item.box.width, item.box.height).clip();
  doc
    .rect(item.box.x, item.box.y, item.box.width, item.box.height)
    .fill(item.background ?? '#ffffff');
  const ops = patternTileOps(item.pattern);
  for (let y = item.box.y; y < item.box.y + item.box.height; y += PATTERN_TILE) {
    for (let x = item.box.x; x < item.box.x + item.box.width; x += PATTERN_TILE) {
      for (const op of ops) {
        if (op.kind === 'fill') {
          doc.fillOpacity(op.opacity);
          doc.rect(x + op.x, y + op.y, op.width, op.height).fill(item.color);
          doc.fillOpacity(1);
        } else {
          doc
            .strokeOpacity(op.opacity)
            .moveTo(x + op.x1, y + op.y1)
            .lineTo(x + op.x2, y + op.y2)
            .lineWidth(1)
            .stroke(item.color);
          doc.strokeOpacity(1);
        }
      }
    }
  }
  doc.restore();
}

function paintLine(doc: PDFKit.PDFDocument, item: LineItem): void {
  doc.save();
  if (item.dash !== null && item.dash.length > 0 && item.dash.every((length) => length > 0)) {
    doc.dash([...item.dash], {});
  }
  doc
    .moveTo(item.x1, item.y1)
    .lineTo(item.x2, item.y2)
    .lineWidth(Math.max(0.1, item.width))
    .stroke(item.color);
  doc.restore();
}

function paintText(
  doc: PDFKit.PDFDocument,
  item: PrintTextItem,
  faces: ReadonlyMap<string, LoadedFace>,
): void {
  if (item.text === '') return;
  const face = faceForRun(item.font, item.text, faces);
  if (face === null) return;
  doc.save();
  doc.rect(item.clip.x, item.clip.y, item.clip.width, item.clip.height).clip();
  doc.font(face.asset.id).fontSize(Math.max(0.1, item.font.size)).fillColor(item.font.color);
  const width = doc.widthOfString(item.text);
  let x = item.x;
  if (item.anchor === 'middle') x -= width / 2;
  else if (item.anchor === 'end') x -= width;
  if (item.rotation !== 0) doc.rotate(item.rotation, { origin: [item.x, item.y] });
  doc.text(item.text, x, item.y, {
    lineBreak: false,
    baseline: 'alphabetic',
    oblique: item.font.italic && face.asset.italic !== true ? true : undefined,
  });
  doc.restore();
}

/**
 * The face a run paints with: the family's face when it covers the text, else
 * the Carlito supplement (metric substitution, listed by the compatibility
 * report), else null. The preflight guarantees this returns a covering face on
 * a vector page.
 */
function faceForRun(
  font: PrintFont,
  text: string,
  faces: ReadonlyMap<string, LoadedFace>,
): LoadedFace | null {
  const primary = primaryAsset(font);
  const primaryFace = primary === undefined ? undefined : faces.get(primary.id);
  if (primaryFace?.font != null && missingCodePoints(primaryFace.font, text).length === 0) {
    return primaryFace;
  }
  const supplement = supplementAssetFor(font.bold, font.italic);
  const supplementFace = supplement === undefined ? undefined : faces.get(supplement.id);
  if (supplementFace?.font != null && missingCodePoints(supplementFace.font, text).length === 0) {
    return supplementFace;
  }
  // The preflight guarantees a covering face on a vector page; this keeps a
  // stray run (a space-only cluster) drawable with the primary face.
  return primaryFace?.font != null ? primaryFace : null;
}

function paintDataBar(doc: PDFKit.PDFDocument, item: DataBarItem): void {
  const { box } = item;
  if (item.gradient === null) {
    doc.rect(box.x, box.y, box.width, box.height).fill(item.color);
  } else {
    const gradient = doc.linearGradient(box.x, box.y, box.x + box.width, box.y);
    gradient.stop(0, item.gradient.from).stop(1, item.gradient.to);
    doc.rect(box.x, box.y, box.width, box.height).fill(gradient);
  }
  if (item.borderColor !== null) {
    doc
      .rect(box.x, box.y, box.width, box.height)
      .lineWidth(1)
      .strokeColor(item.borderColor)
      .stroke();
  }
  if (item.axisX !== null) {
    doc
      .moveTo(item.axisX, box.y)
      .lineTo(item.axisX, box.y + box.height)
      .lineWidth(1)
      .strokeColor(item.axisColor)
      .stroke();
  }
}

function paintIconShape(doc: PDFKit.PDFDocument, shape: IconShape): void {
  doc.save();
  for (const segment of shape.segments) {
    switch (segment.t) {
      case 'move':
        doc.moveTo(segment.x, segment.y);
        break;
      case 'line':
        doc.lineTo(segment.x, segment.y);
        break;
      case 'close':
        doc.closePath();
        break;
      case 'circle':
        doc.circle(segment.x, segment.y, segment.r);
        break;
    }
  }
  if (shape.paint === 'fill') {
    doc.fill(shape.color);
  } else {
    doc.lineWidth(Math.max(0.1, shape.lineWidth)).stroke(shape.color);
  }
  doc.restore();
}

function concatChunks(chunks: readonly Uint8Array[]): Uint8Array {
  const length = chunks.reduce((total, chunk) => total + chunk.length, 0);
  const bytes = new Uint8Array(length);
  let offset = 0;
  for (const chunk of chunks) {
    bytes.set(chunk, offset);
    offset += chunk.length;
  }
  return bytes;
}
