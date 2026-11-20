import { describe, expect, it } from 'vitest';
import { paintCfIcon } from '@grid/cf-paint';
import type { CfIconSetId } from '@model/conditional-format';
import { createSheetProjection } from '@view/projection';
import { openWorkbook } from '@ooxml/open';
import { createPrintPlan } from '../src/print/planner';
import { createPrintSnapshot } from '../src/print/adapter';
import { fixtureBytes } from './helpers/roundtrip';
import {
  CHROMIUM_A4_PAGE,
  printPageBox,
  printPageLabel,
  renderPrintPlan,
  renderPrintPage,
} from '../src/print/svg-renderer';
import type {
  PrintBox,
  PrintIconFamily,
  PrintItem,
  PrintPage,
  PrintPaper,
  PrintPlan,
} from '../src/print/types';

/**
 * M7 lane 2 — the SVG preview renderer (print-pdf §2).
 *
 * The plan is plain data, so these tests hand it page objects directly: each
 * item kind is asserted on its own markup, and the page box follows the
 * spike's measured Chromium A4 box.
 */

const A4_PORTRAIT: PrintPaper = {
  name: 'A4',
  code: 9,
  orientation: 'portrait',
  width: 595.28,
  height: 841.89,
};

const A4_LANDSCAPE: PrintPaper = {
  name: 'A4',
  code: 9,
  orientation: 'landscape',
  width: 841.89,
  height: 595.28,
};

const LETTER: PrintPaper = {
  name: 'Letter',
  code: 1,
  orientation: 'portrait',
  width: 612,
  height: 792,
};

const BODY: PrintBox = { x: 50.4, y: 54, width: 494.48, height: 733.89 };

function pageWith(items: readonly PrintItem[], paper: PrintPaper = A4_PORTRAIT): PrintPage {
  return {
    index: 0,
    setIndex: 0,
    pageInSetIndex: 0,
    paper,
    box: { x: 0, y: 0, width: paper.width, height: paper.height },
    body: BODY,
    band: { startRow: 0, endRow: 0, startCol: 0, endCol: 0 },
    columns: [],
    rows: [],
    items,
    clipped: false,
  };
}

function planWith(pages: readonly PrintPage[], paper: PrintPaper = A4_PORTRAIT): PrintPlan {
  return {
    sheetId: 'sheet-1',
    sheetName: 'Sheet1',
    paper,
    margins: { top: 54, right: 50.4, bottom: 54, left: 50.4 },
    scale: 1,
    fitToPage: false,
    showGridLines: false,
    showHeadings: false,
    horizontalCentered: false,
    verticalCentered: false,
    pageSets: [
      {
        index: 0,
        range: { startRow: 0, startCol: 0, endRow: 0, endCol: 0 },
        firstPage: 0,
        pageCount: pages.length,
      },
    ],
    pages,
  };
}

function svgOf(items: readonly PrintItem[], paper: PrintPaper = A4_PORTRAIT): string {
  return renderPrintPage(pageWith(items, paper), printPageLabel(1, 1, 'Sheet1'));
}

describe('the page root', () => {
  it('is one SVG in page points with an accessible label', () => {
    const svg = svgOf([]);
    expect(svg).toContain('viewBox="0 0 595.28 841.89"');
    expect(svg).toContain('width="595.28pt"');
    expect(svg).toContain('height="841.89pt"');
    expect(svg).toContain('role="img"');
    expect(svg).toContain('aria-label="Page 1 of 1, Sheet1"');
    expect(svg).toContain('data-print-page="1"');
    expect(svg.endsWith('</svg>')).toBe(true);
  });

  it('escapes a hostile sheet name in the label attribute', () => {
    const svg = renderPrintPage(pageWith([]), 'Page 1 of 1, A "B" <C> & D');
    expect(svg).toContain('aria-label="Page 1 of 1, A &quot;B&quot; &lt;C&gt; &amp; D"');
  });

  it('renders one element per plan page with unique definition ids', () => {
    const plan = planWith([
      pageWith([{ kind: 'fill', box: BODY, pattern: 'solid', color: '#FFFF00', background: null }]),
      {
        ...pageWith([
          { kind: 'fill', box: BODY, pattern: 'solid', color: '#00FF00', background: null },
        ]),
        index: 1,
      },
      {
        ...pageWith([
          { kind: 'fill', box: BODY, pattern: 'solid', color: '#0000FF', background: null },
        ]),
        index: 2,
      },
    ]);
    const rendered = renderPrintPlan(plan);
    expect(rendered.pageCount).toBe(3);
    expect(rendered.pages.map((page) => page.label)).toEqual([
      'Page 1 of 3, Sheet1',
      'Page 2 of 3, Sheet1',
      'Page 3 of 3, Sheet1',
    ]);
    const ids = rendered.pages.flatMap((page) =>
      [...page.svg.matchAll(/ id="([^"]+)"/g)].map((match) => match[1]),
    );
    expect(new Set(ids).size).toBe(ids.length);
  });
});

describe('fills', () => {
  it('paints a solid fill as an SVG rect, never a CSS background', () => {
    const svg = svgOf([
      {
        kind: 'fill',
        box: { x: 10, y: 20, width: 30, height: 40 },
        pattern: 'solid',
        color: '#FFFF00',
        background: null,
      },
    ]);
    expect(svg).toContain(
      '<rect x="10" y="20" width="30" height="40" fill="#FFFF00" data-kind="fill"/>',
    );
    expect(svg).not.toContain('background-color');
  });

  it('paints a pattern through an SVG pattern tile with its background', () => {
    const svg = svgOf([
      {
        kind: 'fill',
        box: BODY,
        pattern: 'darkGrid',
        color: '#FF0000',
        background: '#00FF00',
      },
    ]);
    expect(svg).toContain('<pattern id="print-0-pattern-0" patternUnits="userSpaceOnUse"');
    // Eight CSS pixels of tile are six points on paper.
    expect(svg).toContain('width="6" height="6"');
    expect(svg).toContain('fill="#00FF00"');
    expect(svg).toContain('stroke="#FF0000"');
    expect(svg).toContain('fill="url(#print-0-pattern-0)"');
  });

  it('draws nothing for a fill the plan leaves unpainted', () => {
    const svg = svgOf([
      { kind: 'fill', box: BODY, pattern: 'solid', color: null, background: null },
    ]);
    expect(svg).not.toContain('data-kind="fill"');
  });
});

describe('lines', () => {
  it('carries stroke width and the dash array', () => {
    const svg = svgOf([
      {
        kind: 'line',
        x1: 10,
        y1: 20,
        x2: 110,
        y2: 20,
        color: '#000000',
        width: 1.5,
        dash: [4, 2],
      },
    ]);
    expect(svg).toContain(
      '<line x1="10" y1="20" x2="110" y2="20" stroke="#000000" stroke-width="1.5" stroke-linecap="butt" stroke-dasharray="4 2" data-kind="line"/>',
    );
  });

  it('omits the dash array for a solid line', () => {
    const svg = svgOf([
      { kind: 'line', x1: 0, y1: 0, x2: 10, y2: 10, color: '#000000', width: 0.75, dash: null },
    ]);
    expect(svg).not.toContain('stroke-dasharray');
  });
});

describe('text', () => {
  const text = (overrides: Partial<Extract<PrintItem, { kind: 'text' }>> = {}): PrintItem => ({
    kind: 'text',
    text: 'Total',
    x: 100,
    y: 200,
    anchor: 'middle',
    font: { family: 'Calibri', size: 11, bold: true, italic: true, color: '#202020' },
    rotation: 0,
    clip: { x: 50, y: 180, width: 100, height: 30 },
    ...overrides,
  });

  it('sets the anchor, the alphabetic baseline and the font', () => {
    const svg = svgOf([text()]);
    expect(svg).toContain('x="100"');
    expect(svg).toContain('y="200"');
    expect(svg).toContain('text-anchor="middle"');
    expect(svg).toContain('font-family="Calibri, sans-serif"');
    expect(svg).toContain('font-size="11"');
    expect(svg).toContain('font-weight="bold"');
    expect(svg).toContain('font-style="italic"');
    expect(svg).toContain('fill="#202020"');
    expect(svg).toContain('>Total</text>');
  });

  it('keeps a serif or monospace family with its generic fallback', () => {
    const serif = svgOf([
      text({
        font: { family: 'Times New Roman', size: 11, bold: false, italic: false, color: '#000000' },
      }),
    ]);
    expect(serif).toContain('font-family="Times New Roman, serif"');
    const mono = svgOf([
      text({
        font: { family: 'Courier New', size: 11, bold: false, italic: false, color: '#000000' },
      }),
    ]);
    expect(mono).toContain('font-family="Courier New, monospace"');
  });

  it('rotates clockwise about the anchor point', () => {
    const svg = svgOf([text({ rotation: 90 })]);
    expect(svg).toContain('transform="rotate(90 100 200)"');
  });

  it('clips the run to its clip box and shares a clip path', () => {
    const svg = svgOf([
      text(),
      text({ text: 'second' }),
      text({ text: 'other', clip: { x: 0, y: 0, width: 10, height: 10 } }),
    ]);
    expect([...svg.matchAll(/<clipPath /g)]).toHaveLength(2);
    expect([...svg.matchAll(/clip-path="url\(#print-0-clip-0\)"/g)]).toHaveLength(2);
    expect(svg).toContain(
      '<clipPath id="print-0-clip-0"><rect x="50" y="180" width="100" height="30"/></clipPath>',
    );
  });

  it('escapes sheet text so markup can never be injected', () => {
    const svg = svgOf([text({ text: '<script>alert("x") & co</script>' })]);
    expect(svg).not.toContain('<script>');
    expect(svg).toContain('&lt;script&gt;alert("x") &amp; co&lt;/script&gt;');
  });
});

describe('data bars', () => {
  it('paints a solid bar with its border and axis', () => {
    const svg = svgOf([
      {
        kind: 'dataBar',
        box: { x: 10, y: 20, width: 50, height: 15 },
        color: '#638EC6',
        gradient: null,
        borderColor: '#4A6A94',
        axisX: 35,
        axisColor: '#808080',
      },
    ]);
    expect(svg).toContain('data-kind="dataBar"');
    expect(svg).toContain('<rect x="10" y="20" width="50" height="15" fill="#638EC6"/>');
    expect(svg).toContain('fill="none" stroke="#4A6A94" stroke-width="1"');
    expect(svg).toContain(
      '<line x1="35" y1="20" x2="35" y2="35" stroke="#808080" stroke-width="1"/>',
    );
  });

  it('paints the gradient from the left to the right', () => {
    const svg = svgOf([
      {
        kind: 'dataBar',
        box: { x: 0, y: 0, width: 40, height: 12 },
        color: '#638EC6',
        gradient: { from: '#BBD0E8', to: '#638EC6' },
        borderColor: null,
        axisX: null,
        axisColor: '#808080',
      },
    ]);
    expect(svg).toContain('<linearGradient id="print-0-gradient-0" x1="0" y1="0" x2="1" y2="0">');
    expect(svg).toContain('<stop offset="0" stop-color="#BBD0E8"/>');
    expect(svg).toContain('<stop offset="1" stop-color="#638EC6"/>');
    expect(svg).toContain('fill="url(#print-0-gradient-0)"');
    expect(svg).not.toContain('stroke-width="1"');
  });
});

describe('icon glyphs', () => {
  const icon = (overrides: Partial<Extract<PrintItem, { kind: 'icon' }>>): PrintItem => ({
    kind: 'icon',
    box: { x: 40, y: 100, width: 9, height: 9 },
    family: 'arrow',
    color: '#F8696B',
    iconIndex: 0,
    iconCount: 3,
    ...overrides,
  });

  it('paints an arrow with the canvas quarter turn: up-right, sideways, down-right', () => {
    const first = svgOf([icon({ iconIndex: 0, iconCount: 3 })]);
    expect(first).toContain('data-icon-family="arrow"');
    expect(first).toContain('rotate(45 44.5 104.5)');
    const middle = svgOf([icon({ iconIndex: 1, iconCount: 3 })]);
    expect(middle).toContain('rotate(90 44.5 104.5)');
    const last = svgOf([icon({ iconIndex: 2, iconCount: 3 })]);
    expect(last).toContain('rotate(135 44.5 104.5)');
  });

  it('paints a one-icon arrow as the canvas draws it, turned a flat 90°', () => {
    const svg = svgOf([icon({ iconIndex: 0, iconCount: 1 })]);
    expect(svg).toContain('rotate(90 44.5 104.5)');
  });

  it('paints a traffic light as a ringed disc', () => {
    const svg = svgOf([icon({ family: 'traffic', color: '#63BE7B' })]);
    expect(svg).toContain('<circle');
    expect(svg).toContain('fill="#63BE7B"');
    expect(svg).toContain('stroke="#404040"');
  });

  it('paints the three sign faces', () => {
    const cross = svgOf([icon({ family: 'sign', iconIndex: 0, iconCount: 3 })]);
    expect(cross).toContain('fill="#F8696B"');
    expect(cross).toContain('stroke="#FFFFFF"');
    const triangle = svgOf([icon({ family: 'sign', iconIndex: 1, iconCount: 3 })]);
    expect(triangle).toContain('fill="#FFEB84"');
    expect(triangle).toContain('fill="#333333"');
    const tick = svgOf([icon({ family: 'sign', iconIndex: 2, iconCount: 3 })]);
    expect(tick).toContain('fill="#63BE7B"');
    expect(tick).toContain('stroke="#FFFFFF"');
  });

  it('paints a flag on a pole', () => {
    const svg = svgOf([icon({ family: 'flag', color: '#FFEB84' })]);
    expect(svg).toContain('stroke="#404040"');
    expect(svg).toContain('fill="#FFEB84"');
  });

  it('paints rating bars: the filled ones dark, the rest pale', () => {
    const svg = svgOf([icon({ family: 'rating', iconIndex: 0, iconCount: 3 })]);
    expect([...svg.matchAll(/fill="#404040"/g)]).toHaveLength(1);
    expect([...svg.matchAll(/fill="#E3E3E3"/g)]).toHaveLength(2);
  });

  it('paints a quarter disc with the canvas wedge fraction', () => {
    // Five quarters: the canvas is given index + 1 over (count - 1), so the
    // wedges are 1/4, 2/4, 3/4 and then two full discs.
    const quarter = svgOf([icon({ family: 'quarter', iconIndex: 0, iconCount: 5 })]);
    expect(quarter).toContain('fill="#E3E3E3"');
    expect(quarter).toContain('<path d="M 44.5 104.5');
    // size 9 → r 3.87; a quarter turn ends at 3 o'clock.
    expect(quarter).toContain('A 3.87 3.87 0 0 1 48.37 104.5 Z"');
    const half = svgOf([icon({ family: 'quarter', iconIndex: 1, iconCount: 5 })]);
    expect(half).toContain('A 3.87 3.87 0 0 1 44.5 108.37 Z"');
    const threeQuarters = svgOf([icon({ family: 'quarter', iconIndex: 2, iconCount: 5 })]);
    expect(threeQuarters).toContain('A 3.87 3.87 0 1 1 40.63 104.5 Z"');
    for (const index of [3, 4]) {
      const full = svgOf([icon({ family: 'quarter', iconIndex: index, iconCount: 5 })]);
      expect(full).not.toContain('<path');
      expect([...full.matchAll(/<circle /g)]).toHaveLength(2);
      expect(full).toContain('fill="#404040"');
    }
  });

  it('paints a dot', () => {
    const svg = svgOf([icon({ family: 'dot', color: '#000000' })]);
    expect(svg).toContain('r="2.07"');
    expect(svg).toContain('fill="#000000"');
  });
});

describe('image placeholders', () => {
  it('reserve the anchor and paint nothing in v1', () => {
    const svg = svgOf([
      {
        kind: 'image',
        box: { x: 10, y: 10, width: 100, height: 100 },
        anchor: { row: 1, col: 1 },
        placeholder: true,
      },
    ]);
    expect(svg).not.toContain('data-kind="image"');
    expect(svg).not.toContain('<image');
  });
});

describe('paint order and hostile values', () => {
  it('paints items in plan order', () => {
    const svg = svgOf([
      { kind: 'line', x1: 0, y1: 0, x2: 1, y2: 1, color: '#000000', width: 1, dash: null },
      { kind: 'fill', box: BODY, pattern: 'solid', color: '#FFFF00', background: null },
      {
        kind: 'text',
        text: 'x',
        x: 1,
        y: 1,
        anchor: 'start',
        font: { family: 'Calibri', size: 11, bold: false, italic: false, color: '#000000' },
        rotation: 0,
        clip: BODY,
      },
    ]);
    const line = svg.indexOf('data-kind="line"');
    const fill = svg.indexOf('data-kind="fill"');
    const text = svg.indexOf('data-kind="text"');
    expect(line).toBeGreaterThan(-1);
    expect(fill).toBeGreaterThan(line);
    expect(text).toBeGreaterThan(fill);
  });

  it('never emits a colour that is not a plain hex', () => {
    const svg = svgOf([
      {
        kind: 'fill',
        box: BODY,
        pattern: 'solid',
        color: 'red"/><script>alert(1)</script>',
        background: null,
      },
    ]);
    expect(svg).not.toContain('<script>');
    expect(svg).not.toContain('data-kind="fill"');
  });

  it('strips quotes from a hostile font family', () => {
    const svg = svgOf([
      {
        kind: 'text',
        text: 'x',
        x: 1,
        y: 1,
        anchor: 'start',
        font: {
          family: 'Arial" onload="alert(1)',
          size: 11,
          bold: false,
          italic: false,
          color: '#000000',
        },
        rotation: 0,
        clip: BODY,
      },
    ]);
    expect(svg).not.toContain('onload=');
    expect(svg).toContain('font-family="Arial onloadalert1, sans-serif"');
  });
});

describe('the print page box (the A4 trap)', () => {
  it('fits A4 inside Chromium’s real page box, not the nominal one', () => {
    const box = printPageBox(A4_PORTRAIT);
    expect(box.width).toBeLessThan(CHROMIUM_A4_PAGE.width);
    expect(box.height).toBeLessThan(CHROMIUM_A4_PAGE.height);
    // The spike's trap: nominal A4 is 0.32pt wider than Chromium's page.
    expect(box.width).toBeLessThan(595.28);
    expect(A4_PORTRAIT.width).toBeGreaterThan(CHROMIUM_A4_PAGE.width);
  });

  it('swaps the measured box for landscape A4', () => {
    const box = printPageBox(A4_LANDSCAPE);
    expect(box.width).toBeLessThan(CHROMIUM_A4_PAGE.height);
    expect(box.height).toBeLessThan(CHROMIUM_A4_PAGE.width);
    expect(box.width).toBeGreaterThan(box.height);
  });

  it('keeps every paper inside its nominal box', () => {
    for (const paper of [
      A4_PORTRAIT,
      A4_LANDSCAPE,
      LETTER,
      { ...LETTER, name: 'Custom', code: 256, width: 400, height: 500 },
    ]) {
      const box = printPageBox(paper);
      expect(box.width).toBeLessThan(paper.width);
      expect(box.height).toBeLessThan(paper.height);
    }
  });

  it('labels a page with its number and sheet', () => {
    expect(printPageLabel(2, 3, 'Sheet1')).toBe('Page 2 of 3, Sheet1');
    expect(printPageLabel(1, 1, '  ')).toBe('Page 1 of 1, Sheet');
  });
});

describe('N planned pages print as N pages', () => {
  it('keeps a real A4 landscape plan at two print pages inside the browser box', () => {
    const opened = openWorkbook(fixtureBytes('print-areas.xlsx'));
    const sheet = opened.workbook.sheets[0]!;
    const projection = createSheetProjection(opened.workbook, sheet.id, { version: 1 });
    const plan = createPrintPlan(createPrintSnapshot(opened.workbook, projection));
    const rendered = renderPrintPlan(plan);

    expect(plan.paper).toMatchObject({ name: 'A4', orientation: 'landscape' });
    expect(plan.pages).toHaveLength(2);
    expect(rendered.pages).toHaveLength(plan.pages.length);
    for (const page of rendered.pages) {
      // The trap: nominal A4 is wider than Chromium's page box. A page element
      // at the nominal size overflows by 0.32pt and spawns an empty page.
      expect(page.printBox.width).toBeLessThan(CHROMIUM_A4_PAGE.height);
      expect(page.printBox.height).toBeLessThan(CHROMIUM_A4_PAGE.width);
      expect(page.printBox.width).toBeLessThan(page.box.width);
      expect(page.printBox.height).toBeLessThan(page.box.height);
      expect(page.svg).toContain('data-print-page=');
    }
  });
});

// ---------------------------------------------------------------------------
// The glyphs agree with the grid's own artwork
// ---------------------------------------------------------------------------

/**
 * The grid's `paintCfIcon` is the authoritative icon geometry. These tests run
 * it against a recording canvas context and compare the numbers it draws with
 * the numbers the SVG renders: the shared ratios stay in one place (the canvas
 * file), and the mirror in `svg-renderer.ts` cannot drift silently.
 *
 * `cf-paint.ts` is read-only for this lane; nothing here imports a private
 * constant, only the exported painter.
 */

interface Point {
  readonly x: number;
  readonly y: number;
}

interface RecordedArc extends Point {
  readonly r: number;
  /** Canvas angles in radians. */
  readonly start: number;
  readonly end: number;
}

interface RecordedPaint {
  readonly points: readonly Point[];
  readonly closed: boolean;
  readonly arcs: readonly RecordedArc[];
  readonly fill: string | null;
  readonly stroke: string | null;
  readonly lineWidth: number;
}

/**
 * Just enough of `CanvasRenderingContext2D` for `paintCfIcon`, recording what
 * it draws in absolute page coordinates (the arrow's translate is applied).
 */
class RecordingContext {
  lineWidth = 1;
  fillStyle = '#000000';
  strokeStyle = '#000000';
  readonly paints: RecordedPaint[] = [];
  readonly rotations: number[] = [];
  readonly rects: { x: number; y: number; width: number; height: number; fill: string }[] = [];
  private points: Point[] = [];
  private arcs: RecordedArc[] = [];
  private closed = false;
  private offset: Point = { x: 0, y: 0 };
  private readonly stack: Point[] = [];

  save(): void {
    this.stack.push(this.offset);
  }

  restore(): void {
    this.offset = this.stack.pop() ?? { x: 0, y: 0 };
  }

  translate(x: number, y: number): void {
    this.offset = { x: this.offset.x + x, y: this.offset.y + y };
  }

  rotate(angle: number): void {
    this.rotations.push(angle);
  }

  beginPath(): void {
    this.points = [];
    this.arcs = [];
    this.closed = false;
  }

  moveTo(x: number, y: number): void {
    this.points.push({ x: x + this.offset.x, y: y + this.offset.y });
  }

  lineTo(x: number, y: number): void {
    this.points.push({ x: x + this.offset.x, y: y + this.offset.y });
  }

  closePath(): void {
    this.closed = true;
  }

  arc(x: number, y: number, r: number, start: number, end: number): void {
    this.arcs.push({ x: x + this.offset.x, y: y + this.offset.y, r, start, end });
  }

  fill(): void {
    this.paints.push(this.snapshot(this.fillStyle, null));
  }

  stroke(): void {
    this.paints.push(this.snapshot(null, this.strokeStyle));
  }

  fillRect(x: number, y: number, width: number, height: number): void {
    this.rects.push({ x, y, width, height, fill: this.fillStyle });
  }

  private snapshot(fill: string | null, stroke: string | null): RecordedPaint {
    return {
      points: [...this.points],
      arcs: [...this.arcs],
      closed: this.closed,
      fill,
      stroke,
      lineWidth: this.lineWidth,
    };
  }
}

/** The icon set whose family each print family mirrors. */
const GRID_SETS: Readonly<Record<PrintIconFamily, CfIconSetId>> = {
  arrow: '3Arrows',
  traffic: '3TrafficLights1',
  sign: '3Signs',
  flag: '3Flags',
  rating: '4Rating',
  quarter: '5Quarters',
  dot: '4RedToBlack',
};

const ICON_BOX = { x: 40, y: 100, width: 12, height: 12 };

function recordGridIcon(
  family: PrintIconFamily,
  iconIndex: number,
  iconCount: number,
): RecordingContext {
  const context = new RecordingContext();
  paintCfIcon(context as unknown as CanvasRenderingContext2D, {
    kind: 'icon',
    iconSet: GRID_SETS[family],
    iconIndex,
    iconCount,
    box: { x: ICON_BOX.x, y: ICON_BOX.y, size: ICON_BOX.width },
  });
  return context;
}

function iconSvg(family: PrintIconFamily, iconIndex: number, iconCount: number): string {
  return svgOf([
    {
      kind: 'icon',
      box: ICON_BOX,
      family,
      color: '#F8696B',
      iconIndex,
      iconCount,
    },
  ]);
}

function attribute(tag: string, name: string): string | null {
  return new RegExp(`${name}="([^"]*)"`).exec(tag)?.[1] ?? null;
}

interface SvgCircle {
  readonly cx: number;
  readonly cy: number;
  readonly r: number;
  readonly fill: string;
  readonly stroke: string | null;
}

function circles(svg: string): SvgCircle[] {
  return [...svg.matchAll(/<circle ([^>]*)\/>/g)].map((match) => {
    const tag = match[1] ?? '';
    return {
      cx: Number(attribute(tag, 'cx')),
      cy: Number(attribute(tag, 'cy')),
      r: Number(attribute(tag, 'r')),
      fill: attribute(tag, 'fill') ?? '',
      stroke: attribute(tag, 'stroke'),
    };
  });
}

interface SvgRect {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly fill: string;
}

function rects(svg: string): SvgRect[] {
  return [...svg.matchAll(/<rect ([^>]*)\/>/g)].map((match) => {
    const tag = match[1] ?? '';
    return {
      x: Number(attribute(tag, 'x')),
      y: Number(attribute(tag, 'y')),
      width: Number(attribute(tag, 'width')),
      height: Number(attribute(tag, 'height')),
      fill: attribute(tag, 'fill') ?? '',
    };
  });
}

interface SvgPath {
  readonly points: readonly Point[];
  readonly closed: boolean;
  readonly fill: string | null;
  readonly stroke: string | null;
}

/** Every `<path>` without an arc command, as absolute points. */
function paths(svg: string): SvgPath[] {
  return [...svg.matchAll(/<path ([^>]*)\/>/g)]
    .map((match) => match[1] ?? '')
    .filter((tag) => !(attribute(tag, 'd') ?? '').includes(' A '))
    .map((tag) => {
      const d = attribute(tag, 'd') ?? '';
      const numbers = (d.match(/-?\d+(?:\.\d+)?/g) ?? []).map(Number);
      const points: Point[] = [];
      for (let index = 0; index + 1 < numbers.length; index += 2) {
        points.push({ x: numbers[index]!, y: numbers[index + 1]! });
      }
      return {
        points,
        closed: d.trimEnd().endsWith('Z'),
        fill: attribute(tag, 'fill'),
        stroke: attribute(tag, 'stroke'),
      };
    });
}

function rotations(svg: string): { degrees: number; cx: number; cy: number }[] {
  return [...svg.matchAll(/rotate\((-?\d+(?:\.\d+)?) (-?\d+(?:\.\d+)?) (-?\d+(?:\.\d+)?)\)/g)].map(
    (match) => ({
      degrees: Number(match[1]),
      cx: Number(match[2]),
      cy: Number(match[3]),
    }),
  );
}

function closeTo(actual: number, expected: number): void {
  expect(Math.abs(actual - expected)).toBeLessThan(0.02);
}

function pointsCloseTo(actual: readonly Point[], expected: readonly Point[]): void {
  expect(actual).toHaveLength(expected.length);
  for (const [index, point] of expected.entries()) {
    const other = actual[index]!;
    closeTo(other.x, point.x);
    closeTo(other.y, point.y);
  }
}

describe('the glyphs match the grid’s paintCfIcon', () => {
  it('arrow: the same path points and the same rotation, for every icon', () => {
    for (let index = 0; index < 3; index += 1) {
      const grid = recordGridIcon('arrow', index, 3);
      const svg = iconSvg('arrow', index, 3);
      const gridPath = grid.paints.find((paint) => paint.fill !== null)!;
      closeTo(grid.rotations[0]! * (180 / Math.PI), 45 + 45 * index);
      const svgRotation = rotations(svg)[0]!;
      closeTo(svgRotation.degrees, (grid.rotations[0]! * 180) / Math.PI);
      pointsCloseTo(paths(svg)[0]!.points, gridPath.points);
    }
  });

  it('traffic: the same disc radius, centre and ring', () => {
    const grid = recordGridIcon('traffic', 0, 3);
    const disc = grid.paints.find((paint) => paint.arcs.length > 0)!.arcs[0]!;
    const circle = circles(iconSvg('traffic', 0, 3))[0]!;
    closeTo(circle.cx, disc.x);
    closeTo(circle.cy, disc.y);
    closeTo(circle.r, disc.r);
    expect(circle.stroke).toBe('#404040');
  });

  it('sign: every face keeps the canvas discs, triangle, bar and dot', () => {
    for (let index = 0; index < 3; index += 1) {
      const grid = recordGridIcon('sign', index, 3);
      const svg = iconSvg('sign', index, 3);
      const svgCircles = circles(svg);
      const gridArcs = grid.paints.flatMap((paint) => paint.arcs);
      // A disc face draws one arc; the middle face's dot is an arc too.
      expect(gridArcs.length).toBe(svgCircles.length);
      for (const [position, arc] of gridArcs.entries()) {
        const circle = svgCircles[position]!;
        closeTo(circle.cx, arc.x);
        closeTo(circle.cy, arc.y);
        closeTo(circle.r, arc.r);
      }
      const gridPaths = grid.paints.filter((paint) => paint.points.length > 0);
      const svgPaths = paths(svg);
      expect(gridPaths).toHaveLength(svgPaths.length);
      for (const [position, paint] of gridPaths.entries()) {
        pointsCloseTo(svgPaths[position]!.points, paint.points);
      }
    }
  });

  it('flag: the same pole and pennant', () => {
    const grid = recordGridIcon('flag', 0, 3);
    const gridPaths = grid.paints.filter((paint) => paint.points.length > 0);
    const svgPaths = paths(iconSvg('flag', 0, 3));
    expect(gridPaths).toHaveLength(svgPaths.length);
    for (const [position, paint] of gridPaths.entries()) {
      pointsCloseTo(svgPaths[position]!.points, paint.points);
    }
  });

  it('rating: the same bars, gaps and fills', () => {
    const grid = recordGridIcon('rating', 1, 3);
    const svgRects = rects(iconSvg('rating', 1, 3));
    expect(svgRects).toHaveLength(grid.rects.length);
    for (const [position, rect] of grid.rects.entries()) {
      const svgRect = svgRects[position]!;
      closeTo(svgRect.x, rect.x);
      closeTo(svgRect.y, rect.y);
      closeTo(svgRect.width, rect.width);
      closeTo(svgRect.height, rect.height);
      expect(svgRect.fill).toBe(rect.fill);
    }
  });

  it('quarter: the same disc and the same wedge fraction', () => {
    for (let index = 0; index < 5; index += 1) {
      const grid = recordGridIcon('quarter', index, 5);
      const svg = iconSvg('quarter', index, 5);
      const disc = grid.paints.find((paint) => paint.arcs.length > 0)!.arcs[0]!;
      const svgCircles = circles(svg);
      closeTo(svgCircles[0]!.r, disc.r);

      const wedge = grid.paints
        .flatMap((paint) => paint.arcs)
        .find((arc) => Math.abs(arc.start + Math.PI / 2) < 1e-6);
      if (wedge === undefined || wedge.end - wedge.start >= Math.PI * 2 - 1e-6) {
        // A full wedge is a solid disc: two circles, no arc path.
        expect(svgCircles).toHaveLength(2);
        expect(svgCircles[1]!.fill).toBe('#404040');
        continue;
      }
      const fraction = (wedge.end - wedge.start) / (Math.PI * 2);
      const arcPath = /A (-?[\d.]+) (-?[\d.]+) 0 (\d) 1 (-?[\d.]+) (-?[\d.]+) Z/.exec(svg)!;
      closeTo(Number(arcPath[1]), wedge.r);
      const endX = Number(arcPath[4]);
      const endY = Number(arcPath[5]);
      const endAngle = Math.atan2(endY - wedge.y, endX - wedge.x);
      const svgFraction = normalise(endAngle - wedge.start) / (Math.PI * 2);
      closeTo(svgFraction, fraction);
    }
  });

  it('dot: the same radius and centre', () => {
    const grid = recordGridIcon('dot', 0, 4);
    const arc = grid.paints.find((paint) => paint.arcs.length > 0)!.arcs[0]!;
    const circle = circles(iconSvg('dot', 0, 4))[0]!;
    closeTo(circle.cx, arc.x);
    closeTo(circle.cy, arc.y);
    closeTo(circle.r, arc.r);
  });

  it('clamps the index the way the grid does', () => {
    // A hostile index resolves to the last icon in both surfaces.
    const grid = recordGridIcon('quarter', 99, 5);
    const svg = iconSvg('quarter', 99, 5);
    expect(circles(svg)).toHaveLength(2);
    expect(grid.paints.some((paint) => (paint.arcs[0]?.end ?? 0) > 0)).toBe(true);
  });
});

/** An angle in [0, 2π): the canvas arc's start is -π/2. */
function normalise(angle: number): number {
  const full = Math.PI * 2;
  return ((angle % full) + full) % full;
}
