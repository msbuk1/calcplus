import { describe, expect, it } from 'vitest';
import {
  MAX_FILL_CELLS,
  MERGED_FILL_MESSAGE,
  planFill,
  sourceCellFor,
  type FillPlan,
  type FillResult,
} from '@model/fill';
import type { Cell, GridRange, Sheet } from '@model/types';
import { cellStateEquals, cloneCell, createCell, getCell, setCellSnapshot } from '@model/types';
import { numberValue, textValue, type CellValue } from '@model/value';
import { createEmptyWorkbook, createSheet } from '@model/workbook';

/**
 * The model half of the fill handle (feature matrix §1 "Fill handle: copy,
 * relative-reference adjustment, double-click to adjacent data"). The fill
 * copies the source block with relative references adjusted; it never invents
 * a series (feature matrix §9 defers those to v2). Every assertion works on the
 * pure plan, so the command tests own the transaction wrapper.
 */

function workbookWithSheet() {
  const workbook = createEmptyWorkbook();
  const sheet = createSheet('Sheet1', 'xl/worksheets/sheet1.xml');
  workbook.sheets.push(sheet);
  return { workbook, sheet };
}

function range(startRow: number, startCol: number, endRow: number, endCol: number): GridRange {
  return { startRow, startCol, endRow, endCol };
}

function put(sheet: Sheet, row: number, col: number, value: CellValue, styleId = 0): Cell {
  const cell = createCell(value, styleId);
  setCellSnapshot(sheet, row, col, cell);
  return cell;
}

function putFormula(sheet: Sheet, row: number, col: number, formula: string): Cell {
  const cell: Cell = {
    value: numberValue(0),
    formula,
    formulaDerived: false,
    savedResult: numberValue(0),
    styleId: 0,
    richText: null,
    stale: false,
  };
  setCellSnapshot(sheet, row, col, cell);
  return cell;
}

/** A filled plan, failing loudly when the fixture was meant to fill. */
function planOf(result: FillResult): FillPlan {
  if (result.kind !== 'filled') throw new Error(`Expected a fill, got ${JSON.stringify(result)}`);
  return result.plan;
}

/** Apply a plan the way the command's `cell.set` patches do. */
function applyPlan(sheet: Sheet, plan: FillPlan): void {
  for (const band of plan.bands) {
    for (let row = 0; row < band.after.length; row += 1) {
      const afterRow = band.after[row] ?? [];
      for (let col = 0; col < afterRow.length; col += 1) {
        setCellSnapshot(
          sheet,
          band.range.startRow + row,
          band.range.startCol + col,
          cloneCell(afterRow[col] ?? null),
        );
      }
    }
  }
}

function valueAt(sheet: Sheet, row: number, col: number): string {
  const cell = getCell(sheet, row, col);
  if (cell === null) return '<empty>';
  if (cell.value.type === 'text') return cell.value.value;
  if (cell.value.type === 'blank') return '<blank>';
  return String(cell.value.value);
}

function formulaAt(sheet: Sheet, row: number, col: number): string | null {
  return getCell(sheet, row, col)?.formula ?? null;
}

// ---------------------------------------------------------------------------
// Block wrap and reference adjustment
// ---------------------------------------------------------------------------

describe('the block wrap', () => {
  it('repeats a two-row block down and shifts each formula by its own delta', () => {
    const { sheet } = workbookWithSheet();
    put(sheet, 0, 0, numberValue(10));
    putFormula(sheet, 0, 1, 'A1*2');
    put(sheet, 1, 0, numberValue(20));
    putFormula(sheet, 1, 1, 'A2*2');

    const plan = planOf(
      planFill({ sheet, source: range(0, 0, 1, 1), destination: range(0, 0, 5, 1) }),
    );
    applyPlan(sheet, plan);

    // Rows 3-6 repeat rows 1-2; the second tile shifts by four rows.
    expect(valueAt(sheet, 2, 0)).toBe('10');
    expect(formulaAt(sheet, 2, 1)).toBe('A3*2');
    expect(valueAt(sheet, 3, 0)).toBe('20');
    expect(formulaAt(sheet, 3, 1)).toBe('A4*2');
    expect(valueAt(sheet, 4, 0)).toBe('10');
    expect(formulaAt(sheet, 4, 1)).toBe('A5*2');
    expect(valueAt(sheet, 5, 0)).toBe('20');
    expect(formulaAt(sheet, 5, 1)).toBe('A6*2');
    expect(plan.cellCount).toBe(8);
  });

  it('fills a one-cell source over the whole destination', () => {
    const { sheet } = workbookWithSheet();
    put(sheet, 0, 0, numberValue(5));

    const plan = planOf(
      planFill({ sheet, source: range(0, 0, 0, 0), destination: range(0, 0, 1, 2) }),
    );
    applyPlan(sheet, plan);

    for (const [row, col] of [
      [0, 1],
      [0, 2],
      [1, 0],
      [1, 1],
      [1, 2],
    ] as const) {
      expect(valueAt(sheet, row, col)).toBe('5');
    }
  });

  it('repeats the block upward from the source, not from the destination', () => {
    const { sheet } = workbookWithSheet();
    put(sheet, 2, 0, numberValue(10));
    put(sheet, 3, 0, numberValue(20));

    const plan = planOf(
      planFill({ sheet, source: range(2, 0, 3, 0), destination: range(1, 0, 3, 0) }),
    );
    applyPlan(sheet, plan);

    // A2 receives the bottom of the two-row block, as Excel's copy fill does.
    expect(valueAt(sheet, 1, 0)).toBe('20');
    expect(plan.cellRecords[0]?.source).toEqual({ row: 3, col: 0 });
  });

  it('tiles by modulo in both directions', () => {
    const source = range(1, 1, 2, 2);
    const cell = (row: number, col: number) => sourceCellFor(source, row, col);
    expect(cell(1, 1)).toEqual({ row: 1, col: 1 });
    expect(cell(2, 2)).toEqual({ row: 2, col: 2 });
    expect(cell(3, 3)).toEqual({ row: 1, col: 1 });
    expect(cell(0, 0)).toEqual({ row: 2, col: 2 });
    expect(cell(-1, -1)).toEqual({ row: 1, col: 1 });
  });
});

describe('relative-reference adjustment', () => {
  it('shifts relative parts, keeps absolute and mixed anchors, and fills both axes', () => {
    const { sheet } = workbookWithSheet();
    putFormula(sheet, 0, 0, 'A1+$A$1+$A1+A$1');

    const down = planOf(
      planFill({ sheet, source: range(0, 0, 0, 0), destination: range(0, 0, 2, 0) }),
    );
    applyPlan(sheet, down);
    expect(formulaAt(sheet, 1, 0)).toBe('A2+$A$1+$A2+A$1');
    expect(formulaAt(sheet, 2, 0)).toBe('A3+$A$1+$A3+A$1');

    const across = planOf(
      planFill({ sheet, source: range(0, 0, 0, 0), destination: range(0, 0, 0, 2) }),
    );
    applyPlan(sheet, across);
    expect(formulaAt(sheet, 0, 1)).toBe('B1+$A$1+$A1+B$1');
    expect(formulaAt(sheet, 0, 2)).toBe('C1+$A$1+$A1+C$1');
  });

  it('turns a reference that leaves the grid into #REF!', () => {
    const { sheet } = workbookWithSheet();
    putFormula(sheet, 1, 0, 'A1');

    const plan = planOf(
      planFill({ sheet, source: range(1, 0, 1, 0), destination: range(0, 0, 1, 0) }),
    );
    applyPlan(sheet, plan);
    expect(formulaAt(sheet, 0, 0)).toBe('#REF!');
  });

  it('drops the cached result and marks a copied formula stale', () => {
    const { sheet } = workbookWithSheet();
    putFormula(sheet, 0, 0, 'A1*2');

    const plan = planOf(
      planFill({ sheet, source: range(0, 0, 0, 0), destination: range(0, 0, 1, 0) }),
    );
    const after = plan.cellRecords[0]?.after;
    expect(after?.stale).toBe(true);
    expect(after?.savedResult).toBeNull();
    expect(after?.value).toEqual({ type: 'blank' });
  });

  it('copies a shared formula group as its displayed value', () => {
    const { sheet } = workbookWithSheet();
    setCellSnapshot(sheet, 0, 0, {
      value: numberValue(42),
      formula: 'A1*2',
      formulaDerived: true,
      savedResult: numberValue(42),
      styleId: 0,
      richText: null,
      stale: false,
      formulaSource: { kind: 'shared-master', sharedIndex: '0', xml: '<f t="shared" si="0"/>' },
    });

    const plan = planOf(
      planFill({ sheet, source: range(0, 0, 0, 0), destination: range(0, 0, 1, 0) }),
    );
    const after = plan.cellRecords[0]?.after;
    expect(after?.formula).toBeNull();
    expect(after?.value).toEqual({ type: 'number', value: 42 });
    expect(after?.formulaSource).toBeUndefined();
  });

  it('copies rich text as plain text', () => {
    const { sheet } = workbookWithSheet();
    setCellSnapshot(sheet, 0, 0, {
      value: textValue('Bold'),
      formula: null,
      formulaDerived: false,
      savedResult: null,
      styleId: 0,
      richText: { xml: '<si><r><t>Bold</t></r></si>', plainText: 'Bold' },
      stale: false,
    });

    const plan = planOf(
      planFill({ sheet, source: range(0, 0, 0, 0), destination: range(0, 0, 1, 0) }),
    );
    const after = plan.cellRecords[0]?.after;
    expect(after?.richText).toBeNull();
    expect(after?.value).toEqual({ type: 'text', value: 'Bold' });
  });
});

// ---------------------------------------------------------------------------
// What travels, what is left alone
// ---------------------------------------------------------------------------

describe('styles and the source block', () => {
  it('copies the source cell style with its value', () => {
    const { sheet } = workbookWithSheet();
    put(sheet, 0, 0, numberValue(7), 3);

    const plan = planOf(
      planFill({ sheet, source: range(0, 0, 0, 0), destination: range(0, 0, 2, 0) }),
    );
    applyPlan(sheet, plan);
    expect(getCell(sheet, 1, 0)?.styleId).toBe(3);
    expect(getCell(sheet, 2, 0)?.styleId).toBe(3);
  });

  it('never writes the source cells and never reads them as destinations', () => {
    const { sheet } = workbookWithSheet();
    put(sheet, 0, 0, numberValue(1));
    put(sheet, 1, 0, numberValue(2));

    const plan = planOf(
      planFill({ sheet, source: range(0, 0, 1, 0), destination: range(0, 0, 3, 0) }),
    );
    for (const band of plan.bands) {
      expect(band.range.startRow).toBeGreaterThan(1);
    }
    for (const record of plan.cellRecords) {
      expect(record.address.row).toBeGreaterThan(1);
      expect(record.source.row).toBeLessThanOrEqual(1);
    }
  });

  it('records the exact prior cell so undo can restore it', () => {
    const { sheet } = workbookWithSheet();
    put(sheet, 0, 0, numberValue(1));
    put(sheet, 1, 0, textValue('old'), 4);

    const plan = planOf(
      planFill({ sheet, source: range(0, 0, 0, 0), destination: range(0, 0, 1, 0) }),
    );
    const record = plan.cellRecords[0];
    expect(record?.before).toEqual(getCell(sheet, 1, 0));
    expect(record?.before?.styleId).toBe(4);
    expect(record?.after?.value).toEqual({ type: 'number', value: 1 });
  });

  it('clears a destination cell when the source cell is empty', () => {
    const { sheet } = workbookWithSheet();
    put(sheet, 1, 0, textValue('gone'));

    const plan = planOf(
      planFill({ sheet, source: range(0, 0, 0, 0), destination: range(0, 0, 1, 0) }),
    );
    applyPlan(sheet, plan);
    expect(getCell(sheet, 1, 0)).toBeNull();
  });

  it('fills hidden destination rows and columns', () => {
    const { sheet } = workbookWithSheet();
    sheet.rowDimensions.set(2, { hidden: true });
    sheet.columnDimensions.set(1, { hidden: true });
    put(sheet, 0, 0, numberValue(9));

    const plan = planOf(
      planFill({ sheet, source: range(0, 0, 0, 0), destination: range(0, 0, 2, 1) }),
    );
    applyPlan(sheet, plan);
    expect(valueAt(sheet, 2, 0)).toBe('9');
    expect(valueAt(sheet, 1, 1)).toBe('9');
  });
});

// ---------------------------------------------------------------------------
// Rejected and no-op edge cases
// ---------------------------------------------------------------------------

describe('edge cases', () => {
  it('is a no-op when the destination is the source (a click without a drag)', () => {
    const { sheet } = workbookWithSheet();
    put(sheet, 0, 0, numberValue(1));
    const result = planFill({ sheet, source: range(0, 0, 0, 0), destination: range(0, 0, 0, 0) });
    expect(result).toMatchObject({ kind: 'unchanged', reason: 'no-destination' });
  });

  it('is a no-op when nothing changes', () => {
    const { sheet } = workbookWithSheet();
    const result = planFill({ sheet, source: range(0, 0, 0, 0), destination: range(0, 0, 2, 0) });
    expect(result).toMatchObject({ kind: 'unchanged', reason: 'nothing-changed' });
  });

  it("blocks a merge in the destination with Excel's message", () => {
    const { sheet } = workbookWithSheet();
    put(sheet, 0, 0, numberValue(1));
    sheet.merges.push({ range: range(3, 0, 3, 1) });

    const result = planFill({
      sheet,
      source: range(0, 0, 0, 0),
      destination: range(0, 0, 4, 0),
    });
    expect(result).toMatchObject({ kind: 'blocked', code: 'merged', message: MERGED_FILL_MESSAGE });
  });

  it('allows a merge wholly inside the source and does not fill it', () => {
    const { sheet } = workbookWithSheet();
    put(sheet, 0, 0, textValue('merged'));
    sheet.merges.push({ range: range(0, 0, 0, 1) });

    const result = planFill({
      sheet,
      source: range(0, 0, 0, 1),
      destination: range(0, 0, 2, 1),
    });
    expect(result.kind).toBe('filled');
    // The merge list is not part of a fill: only cells change.
    expect(sheet.merges).toHaveLength(1);
  });

  it('blocks a merge that only partly overlaps the source', () => {
    const { sheet } = workbookWithSheet();
    put(sheet, 0, 0, numberValue(1));
    put(sheet, 1, 0, numberValue(2));
    sheet.merges.push({ range: range(0, 0, 2, 0) });

    const result = planFill({
      sheet,
      source: range(0, 0, 1, 0),
      destination: range(0, 0, 3, 0),
    });
    expect(result).toMatchObject({ kind: 'blocked', code: 'merged' });
  });

  it('blocks a destination that does not contain the source', () => {
    const { sheet } = workbookWithSheet();
    put(sheet, 0, 0, numberValue(1));
    const result = planFill({ sheet, source: range(0, 0, 0, 0), destination: range(2, 0, 3, 0) });
    expect(result).toMatchObject({ kind: 'blocked', code: 'invalid-range' });
  });

  it('blocks a fill larger than the transaction budget', () => {
    const { sheet } = workbookWithSheet();
    put(sheet, 0, 0, numberValue(1));
    const result = planFill({
      sheet,
      source: range(0, 0, 0, 0),
      destination: range(0, 0, MAX_FILL_CELLS, 0),
    });
    expect(result).toMatchObject({ kind: 'blocked', code: 'too-large' });
  });

  it('rejects a range outside the grid', () => {
    const { sheet } = workbookWithSheet();
    const result = planFill({
      sheet,
      source: range(0, 0, 0, 0),
      destination: range(0, 0, 1_048_576, 0),
    });
    expect(result).toMatchObject({ kind: 'blocked', code: 'invalid-range' });
  });

  it('keeps every record and band consistent with the after state', () => {
    const { sheet } = workbookWithSheet();
    put(sheet, 0, 0, numberValue(1));
    putFormula(sheet, 0, 1, 'A1+1');
    put(sheet, 1, 0, numberValue(2));
    putFormula(sheet, 1, 1, 'A2+1');

    const plan = planOf(
      planFill({ sheet, source: range(0, 0, 1, 1), destination: range(0, 0, 3, 1) }),
    );
    applyPlan(sheet, plan);
    for (const record of plan.cellRecords) {
      expect(
        cellStateEquals(getCell(sheet, record.address.row, record.address.col), record.after),
      ).toBe(true);
    }
    expect(formulaAt(sheet, 2, 1)).toBe('A3+1');
    expect(formulaAt(sheet, 3, 1)).toBe('A4+1');
  });
});
