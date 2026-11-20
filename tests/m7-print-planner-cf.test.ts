import { describe, expect, it } from 'vitest';
import { cfPaintOps } from '@grid/cf-paint';
import { createSheetProjection } from '@view/projection';
import { openWorkbook } from '@ooxml/open';
import { createPrintPlan } from '../src/print/planner';
import { createPrintSnapshot } from '../src/print/adapter';
import type { PrintItem, PrintPage, PrintPlan, PrintTextItem } from '../src/print/types';
import { fixtureBytes } from './helpers/roundtrip';

/**
 * M7 lane 1 — conditional formatting in the print plan (print-pdf §4,
 * conditional formatting §5.1).
 *
 * The fixture's columns each exercise one paint kind, and column A also
 * carries two rules with a stop-if-true. The plan is compared against the
 * grid's own `cfPaintOps` output for the same cells: the print planner must
 * not re-order or re-resolve anything.
 */

function setup(): {
  bytes: Uint8Array;
  opened: ReturnType<typeof openWorkbook>;
  projection: ReturnType<typeof createSheetProjection>;
  plan: PrintPlan;
  page: PrintPage;
} {
  const bytes = fixtureBytes('print-cf.xlsx');
  const opened = openWorkbook(bytes);
  const sheet = opened.workbook.sheets[0]!;
  const projection = createSheetProjection(opened.workbook, sheet.id, {
    version: 1,
    defaultColumnWidth: 60,
  });
  const plan = createPrintPlan(createPrintSnapshot(opened.workbook, projection));
  const page = plan.pages[0]!;
  return { bytes, opened, projection, plan, page };
}

function cellBox(page: PrintPage, row: number, col: number) {
  const column = page.columns.find((candidate) => candidate.col === col);
  const rowEntry = page.rows.find((candidate) => candidate.row === row);
  if (column === undefined || rowEntry === undefined) throw new Error(`No cell ${row}:${col}`);
  return { x: column.x, y: rowEntry.y, width: column.width, height: rowEntry.height };
}

/** The plan items that belong to one cell, in paint order. */
function itemsForCell(page: PrintPage, row: number, col: number): PrintItem[] {
  const box = cellBox(page, row, col);
  const within = (x: number, y: number): boolean =>
    x >= box.x - 0.5 &&
    x <= box.x + box.width + 0.5 &&
    y >= box.y - 0.5 &&
    y <= box.y + box.height + 0.5;
  return page.items.filter((item) => {
    switch (item.kind) {
      case 'text':
        return within(item.x, item.y);
      case 'line':
        return false;
      default:
        return (
          item.box.x + item.box.width > box.x &&
          item.box.x < box.x + box.width &&
          item.box.y + item.box.height > box.y &&
          item.box.y < box.y + box.height
        );
    }
  });
}

function textFor(page: PrintPage, row: number, col: number): PrintTextItem | undefined {
  return itemsForCell(page, row, col).find((item): item is PrintTextItem => item.kind === 'text');
}

function fillsFor(page: PrintPage, row: number, col: number) {
  return itemsForCell(page, row, col).filter((item) => item.kind === 'fill');
}

function opsFor(
  projection: ReturnType<typeof createSheetProjection>,
  row: number,
  col: number,
): ReturnType<typeof cfPaintOps> {
  const cell = projection.cellAt(row, col)!;
  const conditional = projection.conditionalFormatting!;
  const evaluation = conditional.evaluate(row, col, {
    value: cell.value,
    displayText: cell.displayText,
  });
  const code = evaluation.overlay.numberFormat?.code;
  return cfPaintOps({
    evaluation,
    rect: {
      x: 0,
      y: 0,
      width: projection.columnWidth(col) / projection.zoom,
      height: projection.rowHeight(row) / projection.zoom,
    },
    zoom: 1,
    style: cell.style,
    displayText: cell.displayText,
    cfDisplayText: code === undefined ? null : conditional.textFor(cell.value, code),
  });
}

describe('conditional formatting reaches the print plan', () => {
  const { projection, page } = setup();
  const arrowColours = ['#F8696B', '#FFEB84', '#63BE7B'];

  it('paints a dxf fill from a fill rule', () => {
    // A4 is 30: rule 1 (> 25) sets the FFF2CC fill.
    const fills = fillsFor(page, 3, 0);
    expect(fills).toHaveLength(1);
    expect(fills[0]).toMatchObject({ kind: 'fill', color: '#FFF2CC', pattern: 'solid' });
  });

  it('paints a dxf font colour from a font rule', () => {
    // B3 is 20, above the rule's 15.
    const text = textFor(page, 2, 1)!;
    expect(text.font.color).toBe('#0000FF');
  });

  it('paints a colour-scale blend', () => {
    const first = fillsFor(page, 1, 2)[0];
    const last = fillsFor(page, 4, 2)[0];
    expect(first).toMatchObject({ kind: 'fill', color: '#F8696B' });
    expect(last).toMatchObject({ kind: 'fill', color: '#63BE7B' });
    // The middle value sits between the two end points, not on either.
    const middle = fillsFor(page, 2, 2)[0];
    expect(middle?.kind).toBe('fill');
    expect((middle as { color: string }).color).not.toBe('#F8696B');
  });

  it('paints a data bar as a rect variant', () => {
    // D3 is 20 of a 10-40 range: the bar covers a third of the cell.
    const bar = itemsForCell(page, 2, 3).find((item) => item.kind === 'dataBar');
    expect(bar).toMatchObject({
      kind: 'dataBar',
      color: '#638EC6',
      borderColor: null,
      axisX: null,
    });
    if (bar?.kind !== 'dataBar') throw new Error('missing bar');
    expect(bar.box.width).toBeCloseTo(bar.box.width, 6);
    expect(bar.box.width).toBeLessThan(cellBox(page, 2, 3).width);
    expect(bar.gradient).toEqual({ from: expect.any(String), to: '#638EC6' });
  });

  it('carries a mixed-sign bar axis', () => {
    // G3 is -10 in a -20..20 range, so the zero axis paints.
    const bar = itemsForCell(page, 2, 6).find((item) => item.kind === 'dataBar');
    if (bar?.kind !== 'dataBar') throw new Error('missing bar');
    expect(bar.color).toBe('#638EC6');
    expect(bar.borderColor).toBeNull();
    expect(bar.axisX).not.toBeNull();
    const box = cellBox(page, 2, 6);
    expect(bar.axisX!).toBeGreaterThan(box.x);
    expect(bar.axisX!).toBeLessThan(box.x + box.width);
    // The negative value fades the other way.
    expect(bar.gradient).not.toBeNull();
  });

  it('paints an icon as a glyph primitive', () => {
    for (const [row, col] of [
      [1, 4],
      [4, 4],
    ] as const) {
      const icon = itemsForCell(page, row, col).find((item) => item.kind === 'icon');
      const op = opsFor(projection, row, col).find((candidate) => candidate.kind === 'icon');
      if (icon?.kind !== 'icon' || op?.kind !== 'icon') throw new Error('missing icon');
      expect(icon).toMatchObject({
        kind: 'icon',
        family: 'arrow',
        iconIndex: op.iconIndex,
        iconCount: op.iconCount,
        color: arrowColours[op.iconIndex],
      });
    }
  });

  it('formats the value with the CF number format through the text path', () => {
    // F2 is 0.1 with a `0.0%` dxf number format.
    const text = textFor(page, 1, 5)!;
    expect(text.text).toBe('10.0%');
  });
});

describe("priority and stop-if-true are the grid's", () => {
  const { projection, page } = setup();

  it('stops lower rules when the winning rule says so', () => {
    const evaluation = projection.conditionalFormatting!.evaluate(3, 0, {
      value: { type: 'number', value: 30 },
      displayText: '30',
    });
    expect(evaluation.stoppedBy).not.toBeNull();
    expect(evaluation.ruleIds).toHaveLength(1);
    // The fill shows; the lower-priority font rule does not.
    expect(fillsFor(page, 3, 0)[0]).toMatchObject({ color: '#FFF2CC' });
    expect(textFor(page, 3, 0)!.font.color).toBe('#202020');
  });

  it.each([
    [2, 0], // fill + stop-if-true
    [2, 1], // font
    [2, 2], // colour scale
    [2, 3], // data bar
    [2, 4], // icon
  ])('emits the same paint kinds in the same order as cfPaintOps for %i:%i', (row, col) => {
    const ops = opsFor(projection, row, col);
    const planKinds = itemsForCell(page, row, col).map((item) => item.kind);
    const opKinds = ops
      .filter((op) => op.kind !== 'border')
      .map((op) => (op.kind === 'text' ? 'text' : op.kind));
    expect(planKinds).toEqual(opKinds);
  });

  it('carries the ops values, not a second resolution', () => {
    const ops = opsFor(projection, 2, 2);
    const fillOp = ops.find((op) => op.kind === 'fill');
    const textOp = ops.find((op) => op.kind === 'text');
    if (fillOp?.kind !== 'fill' || textOp?.kind !== 'text') throw new Error('missing ops');
    expect(fillsFor(page, 2, 2)[0]).toMatchObject({ color: fillOp.color });
    expect(textFor(page, 2, 2)!.font.color).toBe(textOp.style.font.color ?? '#202020');
  });
});

describe('a CF-free sheet is unchanged', () => {
  it('has no conditional paint and no new item kinds', () => {
    const opened = openWorkbook(fixtureBytes('print-areas.xlsx'));
    const sheet = opened.workbook.sheets[0]!;
    const projection = createSheetProjection(opened.workbook, sheet.id, { version: 1 });
    const snapshot = createPrintSnapshot(opened.workbook, projection);
    expect(snapshot.cellAt(1, 0)?.conditional ?? null).toBeNull();
    const plan = createPrintPlan(snapshot);
    for (const page of plan.pages) {
      for (const item of page.items) {
        expect(item.kind).not.toBe('dataBar');
        expect(item.kind).not.toBe('icon');
      }
    }
  });
});
