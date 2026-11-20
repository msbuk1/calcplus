import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { preflightPrintPlan, renderPdf } from '../src/print/pdf-renderer';
import type { PrintItem, PrintPage } from '../src/print/types';
import {
  LETTER_LANDSCAPE,
  ONE_PIXEL_PNG,
  box,
  makePage,
  makePlan,
  textItem,
} from './helpers/print-pdf';

/**
 * M7 lane 3 — the PDFKit vector renderer (print-pdf §2, §4; research/26).
 *
 * Poppler is available on the maintainer's machine but not on every machine,
 * so the assertions read the produced bytes: the header/trailer, the page
 * dictionaries, the font descriptors and the image XObject. Research/26's
 * probe made the same checks with `pdfinfo`/`pdffonts`.
 */

const loadFont = async (url: string): Promise<Uint8Array> => {
  const path = url.startsWith('file:')
    ? fileURLToPath(url)
    : new URL(url.replace(/^\/+/, ''), 'file://' + process.cwd() + '/').pathname;
  return new Uint8Array(await readFile(path));
};

function latin1(bytes: Uint8Array): string {
  return Buffer.from(bytes).toString('latin1');
}

function countMatches(text: string, pattern: RegExp): number {
  return (text.match(pattern) ?? []).length;
}

/** PDF object dictionaries are uncompressed; content streams are not. */
function pageCountOf(pdf: string): number {
  return countMatches(pdf, /\/Type\s*\/Page(?![s])/g);
}

function tablePage(
  index: number,
  options: { readonly rows?: number; readonly columns?: number } = {},
): PrintPage {
  const rows = options.rows ?? 32;
  const columns = options.columns ?? 6;
  const items: PrintItem[] = [];
  for (let row = 0; row < rows; row += 1) {
    const y = 54 + row * 20;
    items.push({
      kind: 'fill',
      box: box(50, y, 495, 20),
      pattern: 'solid',
      color: row === 0 ? '#D9E2F3' : row % 2 === 0 ? '#F2F2F2' : null,
      background: null,
    });
    items.push({
      kind: 'line',
      x1: 50,
      y1: y + 20,
      x2: 545,
      y2: y + 20,
      color: '#BFBFBF',
      width: 0.5,
      dash: null,
    });
    for (let column = 0; column < columns; column += 1) {
      const x = 58 + column * 80;
      items.push(
        textItem({
          text: `R${row}C${column}`,
          x,
          y: y + 13,
          clip: box(50, y, 495, 20),
          font: {
            family: 'Calibri',
            size: 9,
            bold: row === 0,
            italic: false,
            color: '#202020',
          },
        }),
      );
    }
  }
  return makePage(index, items);
}

describe('renderPdf', () => {
  it('writes one PDF page per plan page with the plan page boxes', async () => {
    const plan = makePlan([
      makePage(0, [textItem({ text: 'First' })]),
      makePage(1, [textItem({ text: 'Second', x: 60, y: 140 })]),
      makePage(2, [textItem({ text: 'Landscape' })], { paper: LETTER_LANDSCAPE }),
    ]);
    const result = await renderPdf(plan, { loadFont });

    expect(result.pageCount).toBe(3);
    expect(latin1(result.bytes.subarray(0, 8))).toBe('%PDF-1.3');
    expect(latin1(result.bytes.subarray(-6))).toBe('%%EOF\n');

    const pdf = latin1(result.bytes);
    expect(pageCountOf(pdf)).toBe(3);
    expect(pdf).toContain('/MediaBox [0 0 595.28 841.89]');
    expect(pdf).toContain('/MediaBox [0 0 792 612]');
  });

  it('embeds subsetted fonts and renames every descriptor CALP+<family>', async () => {
    const plan = makePlan([
      makePage(0, [
        textItem({ text: 'Calibri text' }),
        textItem({
          text: 'Arial text',
          y: 140,
          font: { family: 'Arial', size: 11, bold: false, italic: false, color: '#202020' },
        }),
        textItem({
          text: 'Segoe text',
          y: 180,
          font: { family: 'Segoe UI', size: 11, bold: false, italic: false, color: '#202020' },
        }),
        textItem({
          text: 'Bold Calibri',
          y: 220,
          font: { family: 'Calibri', size: 11, bold: true, italic: false, color: '#202020' },
        }),
      ]),
    ]);
    const result = await renderPdf(plan, { loadFont });
    const pdf = latin1(result.bytes);

    // Embedded, not referenced: a FontFile2 stream and no standard font.
    expect(pdf).toContain('/FontFile2');
    expect(pdf).not.toContain('/BaseFont /Helvetica');
    expect(pdf).toContain('/ToUnicode');

    // Three names per subset: FontName plus both BaseFont entries.
    for (const name of [
      'CALP+Carlito',
      'CALP+Carlito-Bold',
      'CALP+LiberationSans',
      'CALP+Selawik',
    ]) {
      expect(pdf).toContain(`/FontName /${name}`);
      expect(
        countMatches(pdf, new RegExp(`/BaseFont /${name.replace('+', '\\+')}(?![\\w-])`, 'g')),
      ).toBe(2);
    }
    // Every BaseFont name in the file is unique per weight and starts CALP+.
    const baseFonts = [...pdf.matchAll(/\/BaseFont \/([^\s/]+)/g)].map((match) => match[1]);
    expect(new Set(baseFonts).size).toBe(baseFonts.length / 2);
    for (const name of baseFonts) expect(name).toMatch(/^CALP\+/);
    // The default PDFKit tag (six letters + PostScript name) must not ship.
    expect(pdf).not.toMatch(/[A-Z]{6}\+(Selawik|Carlito|Liberation)/);
    expect(result.embeddedFonts).toEqual(
      expect.arrayContaining(['CALP+Carlito', 'CALP+LiberationSans', 'CALP+Selawik']),
    );
  });

  it('paints fills, pattern fills, lines, data bars and every icon family', async () => {
    const items: PrintItem[] = [
      {
        kind: 'fill',
        box: box(50, 60, 200, 40),
        pattern: 'solid',
        color: '#FFEEAA',
        background: null,
      },
      {
        kind: 'fill',
        box: box(50, 110, 200, 40),
        pattern: 'gray125',
        color: '#444444',
        background: '#FFFFFF',
      },
      { kind: 'line', x1: 50, y1: 160, x2: 250, y2: 160, color: '#000000', width: 1, dash: [3, 2] },
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
        kind: 'dataBar',
        box: box(60, 200, 120, 14),
        color: '#638EC6',
        gradient: null,
        borderColor: null,
        axisX: null,
        axisColor: '#000000',
      },
    ];
    const families = ['arrow', 'traffic', 'sign', 'flag', 'rating', 'quarter', 'dot'] as const;
    families.forEach((family, index) => {
      items.push({
        kind: 'icon',
        box: box(60 + index * 20, 240, 12, 12),
        family,
        color: '#63BE7B',
        iconIndex: index % 3,
        iconCount: family === 'rating' ? 5 : 3,
      });
    });
    items.push({
      kind: 'image',
      box: box(300, 60, 100, 80),
      anchor: { row: 0, col: 0 },
      placeholder: true,
    });

    const plan = makePlan([makePage(0, items)]);
    const result = await renderPdf(plan, { loadFont });
    const pdf = latin1(result.bytes);

    expect(pageCountOf(pdf)).toBe(1);
    expect(pdf).toContain('/ShadingType 2');
    expect(result.bytes.length).toBeGreaterThan(0);
  });

  it('keeps a 20-page report well under 1 MB of vector output', async () => {
    const pages = Array.from({ length: 20 }, (_, index) => tablePage(index));
    const plan = makePlan(pages);
    const result = await renderPdf(plan, { loadFont });

    expect(result.pageCount).toBe(20);
    expect(result.fallbackPages).toEqual([]);
    expect(result.bytes.length).toBeLessThan(1_000_000);
    // The spike's 20-page probe measured 56.9 KB; a vector report stays tiny.
    expect(result.bytes.length).toBeGreaterThan(10_000);
  }, 60_000);
});

describe('preflightPrintPlan', () => {
  it('routes a page with CJK text to the 300 DPI fallback', async () => {
    const plan = makePlan([
      makePage(0, [textItem({ text: 'Latin only' })]),
      makePage(1, [textItem({ text: '汉字测试' })]),
    ]);
    const preflight = await preflightPrintPlan(plan, loadFont);

    expect(preflight.fallbackPages).toEqual([1]);
    expect(preflight.decisions[0].fallback).toBe(false);
    expect(preflight.decisions[1].reason?.code).toBe('missing-glyph');
    expect(preflight.decisions[1].reason?.detail).toContain('U+6C49');
  });

  it('routes mixed right-to-left text to the fallback (script rule, not coverage)', async () => {
    const plan = makePlan([makePage(0, [textItem({ text: 'Total مرحبا 123' })])]);
    const preflight = await preflightPrintPlan(plan, loadFont);

    expect(preflight.fallbackPages).toEqual([0]);
    expect(preflight.decisions[0].reason?.code).toBe('rtl-script');
  });

  it('routes an unknown font family to the fallback instead of substituting silently', async () => {
    const plan = makePlan([
      makePage(0, [
        textItem({
          font: { family: 'Comic Sans MS', size: 11, bold: false, italic: false, color: '#202020' },
        }),
      ]),
    ]);
    const preflight = await preflightPrintPlan(plan, loadFont);

    expect(preflight.fallbackPages).toEqual([0]);
    expect(preflight.decisions[0].reason).toEqual({
      code: 'unknown-family',
      detail: 'Comic Sans MS',
    });
  });

  it('keeps Greek and Cyrillic vector by using the Carlito supplement', async () => {
    const plan = makePlan([
      makePage(0, [
        textItem({
          text: 'λ Ж',
          font: { family: 'Segoe UI', size: 11, bold: false, italic: false, color: '#202020' },
        }),
      ]),
    ]);
    const preflight = await preflightPrintPlan(plan, loadFont);

    expect(preflight.fallbackPages).toEqual([]);
    expect(preflight.decisions[0].substitutedRuns).toBe(1);
    expect(preflight.faces.has('carlito-regular')).toBe(true);
  });
});

describe('renderPdf fallback pages', () => {
  it('places the rasterised PNG as a full-page XObject', async () => {
    const plan = makePlan([
      makePage(0, [textItem({ text: 'Latin only' })]),
      makePage(1, [textItem({ text: '汉字测试' })]),
    ]);
    const rasterised: number[] = [];
    const result = await renderPdf(plan, {
      loadFont,
      rasterize: async (page) => {
        rasterised.push(page.index);
        return {
          pageIndex: page.index,
          dpi: 300,
          width: 2480,
          height: 3508,
          png: ONE_PIXEL_PNG,
          reason: { code: 'missing-glyph', detail: 'U+6C49' },
        };
      },
    });
    const pdf = latin1(result.bytes);

    expect(rasterised).toEqual([1]);
    expect(result.fallbackPages).toEqual([1]);
    expect(pdf).toContain('/Subtype /Image');
    expect(pdf).toContain('/Width 1');
    expect(pageCountOf(pdf)).toBe(2);
  });

  it('refuses to render a fallback page with no rasteriser', async () => {
    const plan = makePlan([makePage(0, [textItem({ text: '汉字测试' })])]);
    await expect(renderPdf(plan, { loadFont })).rejects.toThrow(/300 DPI fallback/);
  });
});
