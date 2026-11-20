import { describe, expect, it } from 'vitest';
import { createSheetProjection } from '@view/projection';
import { formatValue } from '@view/number-format';
import { openWorkbook } from '@ooxml/open';
import { saveWorkbook } from '@ooxml/save';
import { createPrintPlan } from '../src/print/planner';
import { createPrintSnapshot } from '../src/print/adapter';
import type { PrintItem, PrintPlan, PrintTextItem } from '../src/print/types';
import {
  assertByteIdentical,
  assertOnlyPartsChanged,
  fixtureBytes,
  openFixture,
  partText,
  saveNoOp,
} from './helpers/roundtrip';

/**
 * M7 lane 1 — the planner against real fixtures through the display-record
 * adapter, plus the `fitToPage` round-trip (M6 identity-harness style).
 */

const SHEET1 = 'xl/worksheets/sheet1.xml';

function texts(items: readonly PrintItem[]): PrintTextItem[] {
  return items.filter((item): item is PrintTextItem => item.kind === 'text');
}

function planFor(
  bytes: Uint8Array,
  options: {
    zoom?: number;
    showGridLines?: boolean;
    showHeadings?: boolean;
    defaultColumnWidth?: number;
  } = {},
): PrintPlan {
  const opened = openWorkbook(bytes);
  const sheet = opened.workbook.sheets[0]!;
  const projection = createSheetProjection(opened.workbook, sheet.id, {
    version: 1,
    zoom: options.zoom,
    showGridLines: options.showGridLines,
    showHeadings: options.showHeadings,
    defaultColumnWidth: options.defaultColumnWidth,
  });
  return createPrintPlan(createPrintSnapshot(opened.workbook, projection));
}

function allText(plan: PrintPlan): string[] {
  return plan.pages.flatMap((page) => texts(page.items).map((item) => item.text));
}

// ---------------------------------------------------------------------------
// print-areas.xlsx — two print areas, print options, styles, date and currency
// ---------------------------------------------------------------------------

describe('print-areas.xlsx through the display records', () => {
  const bytes = fixtureBytes('print-areas.xlsx');
  const plan = planFor(bytes);

  it('reads both print areas as separate page sets', () => {
    expect(plan.pageSets.map((set) => set.range)).toEqual([
      { startRow: 0, startCol: 0, endRow: 29, endCol: 2 },
      { startRow: 0, startCol: 4, endRow: 5, endCol: 5 },
    ]);
    expect(plan.pageSets.map((set) => set.pageCount)).toEqual([1, 1]);
    expect(plan.pages).toHaveLength(2);
    expect(plan.pages[1]!.band).toEqual({ startRow: 0, endRow: 5, startCol: 4, endCol: 5 });
  });

  it('takes gridlines and headings from printOptions, never the view', () => {
    const withViewOff = planFor(bytes, { showGridLines: false, showHeadings: false });
    expect(withViewOff.showGridLines).toBe(true);
    expect(withViewOff.showHeadings).toBe(true);
    expect(
      withViewOff.pages[0]!.items.filter((item) => item.kind === 'line').length,
    ).toBeGreaterThan(10);
    // 25pt of row-number band before the first column, 15pt of letters above.
    expect(withViewOff.pages[0]!.columns[0]!.x).toBeCloseTo(50.4 + 25, 6);
    expect(withViewOff.pages[0]!.rows[0]!.y).toBeCloseTo(54 + 15, 6);
    expect(allText(withViewOff)).toEqual(expect.arrayContaining(['A', 'B', 'C', '1', '30']));
  });

  it('carries the number-formatted text exactly as the grid paints it', () => {
    const opened = openWorkbook(bytes);
    const sheet = opened.workbook.sheets[0]!;
    const projection = createSheetProjection(opened.workbook, sheet.id, { version: 1 });
    const projectionC2 = projection.cellAt(1, 2)!.displayText;
    const projectionC7 = projection.cellAt(6, 2)!.displayText;
    expect(projectionC2).toBe('1,200.00');
    expect(allText(plan)).toContain(projectionC2);
    expect(allText(plan)).toContain(projectionC7);
    expect(projectionC7).toMatch(/\d{2}\/\d{2}\/\d{2}$/);
  });

  it('carries the imported font, fill and border', () => {
    const page = plan.pages[0]!;
    const bold = texts(page.items).find((item) => item.text === 'Total')!;
    expect(bold.font).toMatchObject({ family: 'Calibri', bold: true, size: 11 });
    const fill = page.items.find((item) => item.kind === 'fill' && item.color === '#FFFF00');
    expect(fill).toBeDefined();
    const border = page.items.find(
      (item) => item.kind === 'line' && item.width === 2.25 && item.color === '#000000',
    );
    expect(border).toBeDefined();
  });

  it('ignores the projection zoom: print geometry is the sheet at 100%', () => {
    const zoomed = planFor(bytes, { zoom: 2 });
    expect(zoomed.pages.map((page) => page.band)).toEqual(plan.pages.map((page) => page.band));
    expect(zoomed.pages[0]!.columns).toEqual(plan.pages[0]!.columns);
    expect(zoomed.pages[0]!.rows).toEqual(plan.pages[0]!.rows);
    expect(zoomed.scale).toBe(plan.scale);
  });
});

// ---------------------------------------------------------------------------
// date1904.xlsx — the 1900 and 1904 serials reach print already formatted
// ---------------------------------------------------------------------------

describe('dates in print', () => {
  it('formats 1904 serials through the same number-format path as the grid', () => {
    const bytes = fixtureBytes('date1904.xlsx');
    const opened = openWorkbook(bytes);
    const sheet = opened.workbook.sheets[0]!;
    expect(opened.workbook.date1904).toBe(true);
    const cell = sheet.cells.get('0:1');
    expect(cell).toBeDefined();
    const code = opened.workbook.styles[cell!.styleId]!.numFmtCode;
    const expected = formatValue(cell!.value, code, true);
    // The fixture's columns are narrow: the width-aware layout prints Excel's
    // `#` overflow marker, exactly as the grid paints it.
    expect(allText(planFor(bytes))).toContain('#######');
    const plan = planFor(bytes, { defaultColumnWidth: 120 });
    expect(allText(plan)).toContain(expected);

    // The same serial is a different date under the 1900 system: the flag is
    // what makes the printed text correct, and the plan carries the result.
    opened.workbook.date1904 = false;
    const projection = createSheetProjection(opened.workbook, sheet.id, {
      version: 2,
      defaultColumnWidth: 120,
    });
    const plan1900 = createPrintPlan(createPrintSnapshot(opened.workbook, projection));
    const changed = allText(plan1900);
    expect(changed).not.toEqual(allText(plan));
    expect(changed).toContain(formatValue(cell!.value, code, false));
  });
});

// ---------------------------------------------------------------------------
// print-setup.xlsx (F16) — orientation, fit-to-page, centring
// ---------------------------------------------------------------------------

describe('print-setup.xlsx honours the page setup', () => {
  const plan = planFor(fixtureBytes('print-setup.xlsx'));

  it('resolves A4 landscape, the authored print area and fit-to-page', () => {
    expect(plan.paper).toMatchObject({
      name: 'A4',
      orientation: 'landscape',
      width: 841.89,
      height: 595.28,
    });
    expect(plan.fitToPage).toBe(true);
    expect(plan.pageSets[0]!.range).toEqual({
      startRow: 0,
      startCol: 0,
      endRow: 19,
      endCol: 3,
    });
    expect(plan.pages[0]!.band).toEqual({ startRow: 0, endRow: 19, startCol: 0, endCol: 3 });
  });

  it('keeps 100% because the 20x4 area fits one page already', () => {
    expect(plan.scale).toBe(1);
    expect(plan.pages).toHaveLength(1);
  });

  it('honours printOptions horizontalCentered', () => {
    expect(plan.horizontalCentered).toBe(true);
    const page = plan.pages[0]!;
    const contentWidth = page.columns.reduce((sum, column) => sum + column.width, 0);
    expect(page.columns[0]!.x).toBeCloseTo(
      plan.margins.left + (page.body.width - contentWidth) / 2,
      6,
    );
  });
});

// ---------------------------------------------------------------------------
// fitToPage and print options round-trip
// ---------------------------------------------------------------------------

describe('fitToPage and printOptions round-trip', () => {
  it('leaves a sheet with printOptions and no sheetPr byte-identical', () => {
    const original = fixtureBytes('print-areas.xlsx');
    const saved = saveNoOp(openFixture('print-areas.xlsx'));
    expect(saved.noOp).toBe(true);
    assertByteIdentical(original, saved.bytes, 'print-areas no-op save');
    assertOnlyPartsChanged(original, saved.bytes, []);
  });

  it('reads fitToPage from sheetPr/pageSetUpPr and saves untouched bytes', () => {
    const original = fixtureBytes('print-setup.xlsx');
    const opened = openFixture('print-setup.xlsx');
    const sheet = opened.workbook.sheets[0]!;
    expect(sheet.printSetup).toMatchObject({
      fitToPage: true,
      scale: 85,
      options: {
        gridLines: false,
        gridLinesSet: false,
        headings: false,
        horizontalCentered: true,
        verticalCentered: false,
      },
    });
    const saved = saveNoOp(opened);
    expect(saved.noOp).toBe(true);
    assertByteIdentical(original, saved.bytes, 'print-setup no-op save after fitToPage parsing');
  });

  it('writes a changed fitToPage as one attribute', () => {
    const original = fixtureBytes('print-setup.xlsx');
    const opened = openFixture('print-setup.xlsx');
    const sheet = opened.workbook.sheets[0]!;
    const beforeXml = partText(original, SHEET1);
    sheet.printSetup = { ...sheet.printSetup!, fitToPage: false };
    const saved = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: new Map(),
    });
    assertOnlyPartsChanged(original, saved.bytes, [SHEET1]);
    const afterXml = partText(saved.bytes, SHEET1);
    expect(afterXml).toContain('<pageSetUpPr fitToPage="0"/>');
    // Nothing else in the part moved.
    expect(afterXml.replace('fitToPage="0"', 'fitToPage="1"')).toBe(beforeXml);
    expect(openWorkbook(saved.bytes).workbook.sheets[0]!.printSetup!.fitToPage).toBe(false);

    // The saved part is now the source, so a second save is byte-identical.
    const second = saveWorkbook({
      workbook: opened.workbook,
      package: saved.package,
      dirtyCells: new Map(),
    });
    assertByteIdentical(saved.bytes, second.bytes, 'second fitToPage save');
  });

  it('writes printOptions without touching pageMargins, pageSetup or sheetPr', () => {
    const original = fixtureBytes('print-setup.xlsx');
    const opened = openFixture('print-setup.xlsx');
    const sheet = opened.workbook.sheets[0]!;
    const setup = sheet.printSetup!;
    sheet.printSetup = {
      ...setup,
      options: { ...setup.options!, gridLines: true, gridLinesSet: true, headings: true },
    };
    const saved = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: new Map(),
    });
    assertOnlyPartsChanged(original, saved.bytes, [SHEET1]);
    const afterXml = partText(saved.bytes, SHEET1);
    expect(afterXml).toContain(
      '<printOptions horizontalCentered="1" verticalCentered="0" gridLines="1" gridLinesSet="1" headings="1"/>',
    );
    expect(afterXml).toContain('<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>');
    expect(afterXml).toContain(
      '<pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0" scale="85" firstPageNumber="1" useFirstPageNumber="1" cellComments="none"/>',
    );
    const reopened = openWorkbook(saved.bytes).workbook.sheets[0]!;
    expect(reopened.printSetup!.options).toMatchObject({
      gridLines: true,
      gridLinesSet: true,
      headings: true,
      horizontalCentered: true,
    });
  });

  it('writes the authored scale when fit-to-page is off', () => {
    const original = fixtureBytes('print-setup.xlsx');
    const opened = openFixture('print-setup.xlsx');
    const sheet = opened.workbook.sheets[0]!;
    sheet.printSetup = { ...sheet.printSetup!, fitToPage: false, scale: 120 };
    const saved = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: new Map(),
    });
    assertOnlyPartsChanged(original, saved.bytes, [SHEET1]);
    expect(partText(saved.bytes, SHEET1)).toContain('scale="120"');
    expect(openWorkbook(saved.bytes).workbook.sheets[0]!.printSetup).toMatchObject({
      fitToPage: false,
      scale: 120,
    });
  });

  it('removes fitToPage when the model clears it', () => {
    const original = fixtureBytes('print-setup.xlsx');
    const opened = openFixture('print-setup.xlsx');
    const sheet = opened.workbook.sheets[0]!;
    sheet.printSetup = { ...sheet.printSetup!, fitToPage: null };
    const saved = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: new Map(),
    });
    assertOnlyPartsChanged(original, saved.bytes, [SHEET1]);
    expect(partText(saved.bytes, SHEET1)).toContain('<pageSetUpPr/>');
    expect(openWorkbook(saved.bytes).workbook.sheets[0]!.printSetup!.fitToPage).toBeNull();
  });

  it('creates sheetPr/pageSetUpPr on a sheet that declares none', () => {
    const original = fixtureBytes('print-areas.xlsx');
    const opened = openFixture('print-areas.xlsx');
    const sheet = opened.workbook.sheets[0]!;
    expect(partText(original, SHEET1)).not.toContain('<sheetPr');
    expect(sheet.printSetup!.fitToPage).toBeNull();
    sheet.printSetup = { ...sheet.printSetup!, fitToPage: true };
    const saved = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: new Map(),
    });
    assertOnlyPartsChanged(original, saved.bytes, [SHEET1]);
    const afterXml = partText(saved.bytes, SHEET1);
    expect(afterXml).toContain(
      '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><dimension ref="A1:F30"/>',
    );
    expect(openWorkbook(saved.bytes).workbook.sheets[0]!.printSetup!.fitToPage).toBe(true);
  });
});
