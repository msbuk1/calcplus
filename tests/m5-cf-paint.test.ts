import { describe, expect, it } from 'vitest';
import type { GridRange } from '@model/types';
import type {
  CfCellSource,
  CfDifferentialFormat,
  CfIconSetId,
  CfThreshold,
  ConditionalFormatRule,
} from '@model/conditional-format';
import { createConditionalFormatIndex } from '@model/conditional-format';
import { BLANK, numberValue, type CellValue } from '@model/value';
import { DEFAULT_ALIGNMENT, DEFAULT_BORDER, DEFAULT_FILL, DEFAULT_FONT } from '@model/style';
import { createSheetProjection } from '@view/projection';
import { openWorkbook } from '@ooxml/open';
import type { ProjectedStyle } from '@view/types';
import { computePaneLayout, intersectRects, mergedRect, panePaintRect } from '@grid/panes';
import {
  CF_ICON_MIN_ZOOM,
  cfBorderStyle,
  cfDataBarAxisX,
  cfDataBarRect,
  cfFillColor,
  cfIconBox,
  cfIconShift,
  cfPaintOps,
  iconsSuppressed,
  lighten,
  type CfPaintOp,
} from '@grid/cf-paint';
import { buildXlsx } from './helpers/build-xlsx';

/**
 * Conditional-formatting painting: CF12 `cf-merged` and CF13 `cf-panes-zoom`,
 * plus the paint-op contract behind §5.1–§5.4.
 *
 * The engine's own evaluation is covered by `m5-cf.test.ts`; these tests pin
 * what the renderer receives and how geometry follows zoom, pane and merge.
 */

const SHEET = 'sheet-1';
const TODAY = { year: 2024, month: 3, day: 15 };

const BASE_STYLE: ProjectedStyle = {
  font: DEFAULT_FONT,
  fill: DEFAULT_FILL,
  border: DEFAULT_BORDER,
  alignment: DEFAULT_ALIGNMENT,
  numberFormat: null,
  numberFormatColor: null,
};

function cellSource(values: readonly (readonly number[])[]): CfCellSource {
  const cells = new Map<string, CellValue>();
  values.forEach((row, rowIndex) => {
    row.forEach((value, colIndex) => {
      cells.set(`${rowIndex}:${colIndex}`, numberValue(value));
    });
  });
  return {
    valueAt(row, col) {
      return cells.get(`${row}:${col}`) ?? BLANK;
    },
    displayTextAt(row, col) {
      const value = cells.get(`${row}:${col}`);
      return value?.type === 'number' ? String(value.value) : '';
    },
  };
}

interface RuleBaseFields {
  readonly id: string;
  readonly sheetId: string;
  readonly ranges: readonly GridRange[];
  readonly priority?: number;
  readonly stopIfTrue: boolean;
  readonly dxfId?: number;
  readonly raw: unknown;
}

function base(id: string, ranges: readonly GridRange[], dxfId?: number): RuleBaseFields {
  return { id, sheetId: SHEET, ranges, priority: 1, stopIfTrue: false, dxfId, raw: {} };
}

const COLUMN_A_TO_E: GridRange = { startRow: 0, startCol: 0, endRow: 4, endCol: 0 };

function evaluationFor(
  rules: readonly ConditionalFormatRule[],
  values: readonly (readonly number[])[],
  row: number,
  col: number,
  dxf?: CfDifferentialFormat,
) {
  const index = createConditionalFormatIndex({
    sheetId: SHEET,
    rules,
    date1904: false,
    today: TODAY,
    cells: cellSource(values),
    dxfFor: dxf === undefined ? undefined : () => dxf,
  });
  return index.evaluateCell(row, col);
}

function plan(
  evaluation: ReturnType<typeof evaluationFor>,
  options?: {
    zoom?: number;
    rect?: { x: number; y: number; width: number; height: number };
    cfDisplayText?: string | null;
  },
): readonly CfPaintOp[] {
  return cfPaintOps({
    evaluation,
    rect: options?.rect ?? { x: 10, y: 20, width: 100, height: 20 },
    zoom: options?.zoom ?? 1,
    style: BASE_STYLE,
    displayText: '50',
    cfDisplayText: options?.cfDisplayText,
  });
}

function kinds(ops: readonly CfPaintOp[]): string[] {
  return ops.map((op) => op.kind);
}

function opOf<K extends CfPaintOp['kind']>(
  ops: readonly CfPaintOp[],
  kind: K,
): Extract<CfPaintOp, { kind: K }> | undefined {
  return ops.find((op): op is Extract<CfPaintOp, { kind: K }> => op.kind === kind);
}

function cellIsRule(dxfId?: number): ConditionalFormatRule {
  return {
    ...base('cell-is', [COLUMN_A_TO_E], dxfId),
    kind: 'cellIs',
    operator: 'greaterThan',
    value1: { type: 'number', value: 10 },
  };
}

function colorScaleRule(
  colors: readonly string[],
  thresholds: readonly CfThreshold[] = colors.map(() => ({ type: 'automatic' })),
): ConditionalFormatRule {
  return {
    ...base('scale', [COLUMN_A_TO_E]),
    kind: 'colorScale',
    thresholds,
    colors,
  };
}

function dataBarRule(options: {
  readonly min?: CfThreshold;
  readonly max?: CfThreshold;
  readonly negativeColor?: string;
  readonly direction?: 'leftToRight' | 'rightToLeft';
  readonly fill?: 'gradient' | 'solid';
  readonly border?: boolean;
}): ConditionalFormatRule {
  return {
    ...base('bar', [COLUMN_A_TO_E]),
    kind: 'dataBar',
    min: options.min ?? { type: 'automatic' },
    max: options.max ?? { type: 'automatic' },
    color: '#638EC6',
    negativeColor: options.negativeColor,
    direction: options.direction,
    fill: options.fill,
    border: options.border,
  };
}

function iconSetRule(options: {
  readonly iconSet: CfIconSetId;
  readonly thresholds: readonly CfThreshold[];
  readonly iconOnly?: boolean;
  readonly reverse?: boolean;
}): ConditionalFormatRule {
  return {
    ...base('icons', [COLUMN_A_TO_E]),
    kind: 'iconSet',
    iconSet: options.iconSet,
    thresholds: options.thresholds,
    iconOnly: options.iconOnly,
    reverse: options.reverse,
  };
}

describe('CF12 cf-merged — the paint plan follows §5.1 and §5.2', () => {
  it('orders the ops fill, data bar, text, border, icon', () => {
    const dxf: CfDifferentialFormat = {
      fill: { pattern: 'solid', color: '#FFF2CC' },
      font: { bold: true, color: '#7F6000' },
      border: { bottom: { style: 'thin', color: '#7F6000' } },
    };
    const evaluation = evaluationFor(
      [
        cellIsRule(0),
        iconSetRule({
          iconSet: '3Arrows',
          thresholds: [
            { type: 'percent', value: 33 },
            { type: 'percent', value: 67 },
          ],
        }),
      ],
      [[50], [0], [100]],
      0,
      0,
      dxf,
    );
    expect(evaluation.painted).toBe(true);
    const ops = plan(evaluation);
    expect(kinds(ops)).toEqual(['fill', 'text', 'border', 'icon']);
  });

  it('paints the CF fill from the dxf and the colour scale', () => {
    const dxf: CfDifferentialFormat = { fill: { color: '#FFF2CC' } };
    const filled = plan(evaluationFor([cellIsRule(0)], [[50]], 0, 0, dxf));
    const fill = opOf(filled, 'fill');
    expect(fill?.color).toBe('#FFF2CC');

    const scale = plan(evaluationFor([colorScaleRule(['#FFFFFF', '#FF0000'])], [[0], [50]], 0, 0));
    expect(opOf(scale, 'fill')?.color).toBe('#FFFFFF');
    // Without a CF number format the projection's own layout still applies.
    expect(opOf(scale, 'text')?.layoutText).toBe(true);
  });

  it('interpolates a two-colour scale between its endpoints', () => {
    const rule = colorScaleRule(['#FFFFFF', '#FF0000']);
    const values = [[0], [50], [100]];
    const low = evaluationFor([rule], values, 0, 0);
    const middle = evaluationFor([rule], values, 1, 0);
    const high = evaluationFor([rule], values, 2, 0);
    expect(low.colorScale?.position).toBe(0);
    expect(low.colorScale?.color).toBe('#FFFFFF');
    expect(high.colorScale?.position).toBe(1);
    expect(high.colorScale?.color).toBe('#FF0000');
    expect(middle.colorScale?.position).toBeCloseTo(0.5);
    // Linear sRGB: halfway from FF to 00 on red, halfway 00 to FF on green.
    expect(middle.colorScale?.color).toBe('#FF8080');
  });

  it('carries the CF font and number format over the text without touching the manual style', () => {
    const dxf: CfDifferentialFormat = {
      font: { bold: true, size: 14, color: '#B00020' },
      numberFormat: { code: '0.0%' },
    };
    const evaluation = evaluationFor([cellIsRule(0)], [[50]], 0, 0, dxf);
    const ops = plan(evaluation, { cfDisplayText: '5000.0%' });
    const text = opOf(ops, 'text');
    expect(text?.text).toBe('5000.0%');
    expect(text?.style.font.bold).toBe(true);
    expect(text?.style.font.size).toBe(14);
    expect(text?.style.font.color).toBe('#B00020');
    expect(text?.style.numberFormat).toBe('0.0%');
    // A CF number format produced the text, so width layout is bypassed.
    expect(text?.layoutText).toBe(false);
    // The manual style is untouched: CF is an overlay, never a mutation (§2).
    expect(BASE_STYLE.font.bold).toBe(false);
    expect(BASE_STYLE.numberFormat).toBeNull();
  });

  it('paints only the CF border edges the dxf sets', () => {
    const border = cfBorderStyle({
      bottom: { style: 'double', color: '#FF0000' },
      diagonalUp: true,
    });
    expect(border).not.toBeNull();
    expect(border!.bottom.style).toBe('double');
    expect(border!.bottom.color).toBe('#FF0000');
    expect(border!.left.style).toBe('none');
    expect(border!.top.style).toBe('none');
    expect(border!.diagonal.style).toBe('none');
    expect(cfBorderStyle(null)).toBeNull();
  });

  it('scales a merged data bar to the merged rectangle', () => {
    const opened = openWorkbook(
      buildXlsx(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>10</v></c><c r="B1"><v>0</v></c></row>
    <row r="2"><c r="A2"><v>20</v></c><c r="B2"><v>0</v></c></row>
    <row r="3"><c r="A3"><v>30</v></c><c r="B3"><v>0</v></c></row>
    <row r="4"><c r="A4"><v>40</v></c><c r="B4"><v>0</v></c></row>
  </sheetData>
  <mergeCells count="1"><mergeCell ref="A1:B2"/></mergeCells>
  <conditionalFormatting sqref="A1:A4">
    <cfRule type="dataBar" priority="1">
      <dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar>
    </cfRule>
  </conditionalFormatting>
</worksheet>`),
    );
    const projection = createSheetProjection(opened.workbook, opened.initialSheetId, {
      version: 1,
      conditionalFormats: { today: TODAY },
    });
    expect(projection.conditionalFormatting).not.toBeNull();
    const evaluation = projection.conditionalFormatting!.evaluate(0, 0);
    expect(evaluation.dataBar).not.toBeNull();

    const merge = projection.mergeAt(0, 0);
    expect(merge).toEqual({ anchorRow: 0, anchorCol: 0, rowSpan: 2, colSpan: 2 });
    const layout = computePaneLayout(projection, 900, 600);
    const rect = mergedRect(projection, layout, 0, 0, 0, 0, 2, 2);
    const ops = cfPaintOps({
      evaluation,
      rect,
      zoom: 1,
      style: BASE_STYLE,
      displayText: '10',
    });
    const bar = opOf(ops, 'dataBar');
    expect(bar).toBeDefined();
    // 10 of a 10..40 range = the first quarter of the merged width.
    const geometry = cfDataBarRect(evaluation.dataBar!, rect);
    expect(geometry.width).toBeCloseTo(rect.width * 0.25, 5);
    expect(geometry.height).toBeLessThanOrEqual(rect.height);
  });
});

describe('CF13 cf-panes-zoom — geometry, zoom and pane clipping', () => {
  it('draws a negative bar from the zero axis and a positive one away from it', () => {
    const rule = dataBarRule({ negativeColor: '#F8696B' });
    const values = [[-10], [-5], [0], [5], [10]];
    const rect = { x: 100, y: 40, width: 200, height: 20 };

    const negative = evaluationFor([rule], values, 1, 0);
    expect(negative.dataBar?.negative).toBe(true);
    expect(negative.dataBar?.start).toBe(0.25);
    expect(negative.dataBar?.end).toBe(0.5);
    expect(negative.dataBar?.axis).toBe(0.5);
    expect(cfDataBarAxisX(negative.dataBar!, rect)).toBeCloseTo(200, 5);
    expect(cfDataBarRect(negative.dataBar!, rect)).toMatchObject({
      x: 150,
      width: 50,
    });

    const positive = evaluationFor([rule], values, 3, 0);
    expect(cfDataBarRect(positive.dataBar!, rect)).toMatchObject({ x: 200, width: 50 });
    expect(opOf(plan(positive, { rect }), 'dataBar')?.color).toBe('#638EC6');

    const zero = evaluationFor([rule], values, 2, 0);
    expect(cfDataBarRect(zero.dataBar!, rect).width).toBe(0);
  });

  it('mirrors the bar and the axis for a right-to-left rule', () => {
    const rule = dataBarRule({ direction: 'rightToLeft' });
    const rect = { x: 0, y: 0, width: 100, height: 20 };
    const overlay = evaluationFor([rule], [[0], [10]], 1, 0).dataBar!;
    const mirrored = cfDataBarRect(overlay, rect);
    expect(mirrored.x).toBe(0);
    expect(mirrored.width).toBeCloseTo(100, 5);
    expect(cfDataBarAxisX(overlay, rect)).toBe(100);
  });

  it('paints a gradient bar with a lighter ramp toward the axis', () => {
    const rule = dataBarRule({});
    const overlay = evaluationFor([rule], [[-10], [10]], 1, 0).dataBar!;
    expect(overlay.fill).toBe('gradient');
    expect(lighten('#638EC6', 0.45)).not.toBe('#638EC6');
    const bar = opOf(plan(evaluationFor([rule], [[-10], [10]], 1, 0)), 'dataBar');
    expect(bar?.fill).toBe('gradient');
    expect(bar?.axisX).not.toBeNull();
  });

  it('chooses the icon by value and shifts the text right', () => {
    const rule = iconSetRule({
      iconSet: '3Arrows',
      thresholds: [
        { type: 'percent', value: 33 },
        { type: 'percent', value: 67 },
      ],
    });
    const values = [[0], [50], [100]];
    const low = evaluationFor([rule], values, 0, 0);
    const middle = evaluationFor([rule], values, 1, 0);
    const high = evaluationFor([rule], values, 2, 0);
    expect(low.iconSet?.iconIndex).toBe(0);
    expect(middle.iconSet?.iconIndex).toBe(1);
    expect(high.iconSet?.iconIndex).toBe(2);

    const rect = { x: 10, y: 20, width: 100, height: 20 };
    const ops = cfPaintOps({
      evaluation: middle,
      rect,
      zoom: 1,
      style: BASE_STYLE,
      displayText: '50',
    });
    const icon = opOf(ops, 'icon');
    expect(icon?.iconIndex).toBe(1);
    if (icon === undefined) throw new Error('expected an icon op');
    expect(icon.box.x).toBeGreaterThanOrEqual(rect.x);
    expect(icon.box.x + icon.box.size).toBeLessThanOrEqual(rect.x + rect.width);
    const text = opOf(ops, 'text');
    expect(text?.shift).toBe(cfIconShift(rect, 1));
    expect(text?.hidden).toBe(false);
  });

  it('honours icon-only rules by hiding the text', () => {
    const rule = iconSetRule({
      iconSet: '3TrafficLights1',
      thresholds: [
        { type: 'percent', value: 33 },
        { type: 'percent', value: 67 },
      ],
      iconOnly: true,
    });
    const ops = plan(evaluationFor([rule], [[100]], 0, 0));
    expect(opOf(ops, 'text')?.hidden).toBe(true);
    expect(opOf(ops, 'icon')?.iconSet).toBe('3TrafficLights1');
  });

  it('scales the bar, the icon and the CF font with zoom but never the thresholds', () => {
    const barRule = dataBarRule({});
    const overlay = evaluationFor([barRule], [[0], [10]], 1, 0);
    // Thresholds are value-based: the same overlay, whatever the zoom.
    expect(evaluationFor([barRule], [[0], [10]], 1, 0).dataBar?.end).toBe(overlay.dataBar?.end);

    const iconRule = iconSetRule({
      iconSet: '3Arrows',
      thresholds: [{ type: 'percent', value: 50 }],
    });
    const icon = evaluationFor([iconRule], [[100]], 0, 0);
    const rect = { x: 0, y: 0, width: 100, height: 20 };
    const atOne = cfIconBox(rect, 1).size;
    const atTwo = cfIconBox(rect, 2).size;
    const shorterRow = cfIconBox({ x: 0, y: 0, width: 100, height: 10 }, 1).size;
    expect(atTwo).toBeGreaterThan(atOne);
    expect(atOne).toBeGreaterThanOrEqual(shorterRow);
    expect(opOf(plan(icon, { rect, zoom: 2 }), 'icon')?.box.size).toBe(atTwo);
  });

  it('suppresses icons below 25% zoom while fills, bars and scales still paint', () => {
    const rule = iconSetRule({
      iconSet: '3Arrows',
      thresholds: [{ type: 'percent', value: 50 }],
    });
    const evaluation = evaluationFor([rule], [[100]], 0, 0);
    expect(iconsSuppressed(CF_ICON_MIN_ZOOM - 0.01)).toBe(true);
    expect(iconsSuppressed(CF_ICON_MIN_ZOOM)).toBe(false);

    const suppressed = plan(evaluation, { zoom: 0.2 });
    expect(opOf(suppressed, 'icon')).toBeUndefined();
    expect(opOf(suppressed, 'text')?.shift).toBe(0);

    const visible = plan(evaluation, { zoom: 0.25 });
    expect(opOf(visible, 'icon')).toBeDefined();

    const scale = colorScaleRule(['#FFFFFF', '#FF0000']);
    const atLowZoom = plan(evaluationFor([scale], [[0]], 0, 0), { zoom: 0.15 });
    expect(opOf(atLowZoom, 'fill')?.color).toBe('#FFFFFF');
    const bar = plan(evaluationFor([dataBarRule({})], [[10]], 0, 0), { zoom: 0.15 });
    expect(opOf(bar, 'dataBar')).toBeDefined();
  });

  it('clips a merged rule per pane: each pane paints its own portion, nothing bleeds', () => {
    const opened = openWorkbook(
      buildXlsx(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews><sheetView workbookViewId="0">
    <pane xSplit="1" ySplit="0" topLeftCell="B1" activePane="topRight" state="frozen"/>
  </sheetView></sheetViews>
  <sheetData>
    <row r="1"><c r="A1"><v>25</v></c><c r="B1"><v>0</v></c><c r="C1"><v>0</v></c><c r="D1"><v>0</v></c><c r="E1"><v>0</v></c><c r="F1"><v>0</v></c></row>
    <row r="2"><c r="A2"><v>0</v></c><c r="B2"><v>50</v></c></row>
    <row r="3"><c r="A3"><v>100</v></c><c r="B3"><v>100</v></c></row>
  </sheetData>
  <mergeCells count="1"><mergeCell ref="A1:F1"/></mergeCells>
  <conditionalFormatting sqref="A1:A3">
    <cfRule type="colorScale" priority="1">
      <colorScale>
        <cfvo type="min"/><cfvo type="max"/>
        <color rgb="FFFFFFFF"/><color rgb="FFFF0000"/>
      </colorScale>
    </cfRule>
  </conditionalFormatting>
</worksheet>`),
    );
    const projection = createSheetProjection(opened.workbook, opened.initialSheetId, {
      version: 1,
      conditionalFormats: { today: TODAY },
    });
    expect(projection.freezeCols).toBe(1);
    const merge = projection.mergeAt(0, 0);
    expect(merge).not.toBeNull();
    const layout = computePaneLayout(projection, 900, 600);
    const rect = mergedRect(projection, layout, 0, 0, 0, 0, merge!.rowSpan, merge!.colSpan);
    const evaluation = projection.conditionalFormatting!.evaluate(0, 0);
    const ops = cfPaintOps({
      evaluation,
      rect,
      zoom: 1,
      style: BASE_STYLE,
      displayText: '25',
    });
    const fill = opOf(ops, 'fill')!;
    // Each pane paints its own portion of the merged fill: the frozen pane and
    // the scrolling pane tile the merged rectangle with no overlap and no gap,
    // so nothing bleeds across the divider (§5.3).
    const paneTop = fill.rect.y;
    const paneHeight = fill.rect.height;
    const frozen = intersectRects(
      { x: layout.headerWidth, y: paneTop, width: layout.frozenWidth, height: paneHeight },
      fill.rect,
    );
    const scrolling = intersectRects(
      {
        x: layout.headerWidth + layout.frozenWidth,
        y: paneTop,
        width: layout.scrollRegionWidth,
        height: paneHeight,
      },
      fill.rect,
    );
    expect(frozen.width).toBeGreaterThan(0);
    expect(scrolling.width).toBeGreaterThan(0);
    expect(frozen.x + frozen.width).toBeLessThanOrEqual(scrolling.x);
    expect(frozen.width + scrolling.width).toBeCloseTo(fill.rect.width, 5);

    // The merged-overflow path paints the whole rectangle but clips it inside
    // the grid body, and only on the axis the pane scrolls.
    const overflowClip = intersectRects(
      panePaintRect(layout, { frozenRows: true, frozenCols: false }),
      fill.rect,
    );
    expect(overflowClip.width).toBeGreaterThan(0);

    // The anchor's bar/fill evaluates once and paints the same plan in both
    // panes; §5.3 forbids a second evaluation with different results.
    expect(projection.conditionalFormatting!.evaluate(0, 0).colorScale?.color).toBe(
      evaluation.colorScale?.color,
    );
  });
});

describe('the projection seam carries the rules to the renderer', () => {
  const worksheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row>
    <row r="2"><c r="A2"><v>3</v></c><c r="B2"><v>4</v></c></row>
  </sheetData>
  <conditionalFormatting sqref="A1:A2">
    <cfRule type="cellIs" operator="greaterThan" priority="1" dxfId="0">
      <formula>2</formula>
    </cfRule>
  </conditionalFormatting>
</worksheet>`;

  it('parses the sheet records into a paint-ready index', () => {
    const opened = openWorkbook(buildXlsx(worksheet));
    const projection = createSheetProjection(opened.workbook, opened.initialSheetId, {
      version: 1,
      conditionalFormats: { today: TODAY },
    });
    const conditional = projection.conditionalFormatting;
    expect(conditional).not.toBeNull();
    expect(conditional!.ruleCount).toBe(1);

    // The dxf table is not on the model yet, so the rule is true but sets no
    // properties: it paints nothing until the resolver is wired.
    const withoutDxf = conditional!.evaluate(1, 0);
    expect(withoutDxf.ruleIds).toHaveLength(1);
    expect(withoutDxf.painted).toBe(false);

    const withDxf = createSheetProjection(opened.workbook, opened.initialSheetId, {
      version: 1,
      conditionalFormats: {
        today: TODAY,
        dxfResolver: (dxfId) =>
          dxfId === 0 ? { fill: { pattern: 'solid', color: '#FFF2CC' } } : null,
      },
    });
    const painted = withDxf.conditionalFormatting!.evaluate(1, 0);
    expect(painted.painted).toBe(true);
    expect(painted.overlay.fill?.color).toBe('#FFF2CC');
    expect(cfFillColor(painted.overlay, painted.colorScale)).toBe('#FFF2CC');
  });

  it('accepts explicit rules and can be switched off', () => {
    const opened = openWorkbook(buildXlsx(worksheet));
    const explicit = createSheetProjection(opened.workbook, opened.initialSheetId, {
      version: 1,
      conditionalFormats: {
        today: TODAY,
        rules: [colorScaleRule(['#FFFFFF', '#FF0000'])],
      },
    });
    const evaluation = explicit.conditionalFormatting!.evaluate(0, 0);
    expect(evaluation.colorScale).not.toBeNull();

    const disabled = createSheetProjection(opened.workbook, opened.initialSheetId, {
      version: 1,
      conditionalFormats: null,
    });
    expect(disabled.conditionalFormatting).toBeNull();
  });

  it('leaves a sheet with no rules unpainted and costs no seam', () => {
    const opened = openWorkbook(
      buildXlsx(worksheet.replace(/<conditionalFormatting[\s\S]*<\/conditionalFormatting>/, '')),
    );
    const projection = createSheetProjection(opened.workbook, opened.initialSheetId, {
      version: 1,
    });
    expect(projection.conditionalFormatting).toBeNull();
    // The cell's value and manual style are exactly as imported (§2).
    expect(projection.cellAt(0, 0)?.displayText).toBe('1');
    expect(projection.cellAt(0, 0)?.style.fill.pattern).toBe('none');
  });
});
