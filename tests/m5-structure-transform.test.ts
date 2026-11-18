import { describe, expect, it } from 'vitest';
import {
  planAutoFilterTransform,
  planRangeTransform,
  transformAttachmentMap,
  transformCellAddress,
  transformConditionalFormatRecord,
  transformDimensionMap,
  transformFormulaReferences,
  transformMergeRange,
  type StructuralEdit,
} from '@model/range-transform';
import { createAutoFilter, evaluateAutoFilter } from '@model/autofilter';
import type { Cell, GridRange, Sheet, ValidationRecord } from '@model/types';
import { cellKey, setCellSnapshot } from '@model/types';
import { numberValue, textValue } from '@model/value';
import { createEmptyWorkbook, createSheet } from '@model/workbook';

/**
 * The M5 range-transform surface, fixtures F11, F12, CF17 and CF20's transform
 * half. The command-level halves (one transaction, undo) live in
 * `m5-structure-commands.test.ts`.
 */

function workbookWithSheet(name = 'Sheet1') {
  const workbook = createEmptyWorkbook();
  const sheet = createSheet(name, 'xl/worksheets/sheet1.xml');
  workbook.sheets.push(sheet);
  return { workbook, sheet };
}

function putFormula(sheet: Sheet, row: number, col: number, formula: string): void {
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
}

function insertRow(sheetId: string, index: number, count = 1): StructuralEdit {
  return { kind: 'insert', sheetId, axis: 'row', index, count };
}

function deleteRow(sheetId: string, index: number, count = 1): StructuralEdit {
  return { kind: 'delete', sheetId, axis: 'row', index, count };
}

function insertColumn(sheetId: string, index: number, count = 1): StructuralEdit {
  return { kind: 'insert', sheetId, axis: 'column', index, count };
}

function range(startRow: number, startCol: number, endRow: number, endCol: number): GridRange {
  return { startRow, startCol, endRow, endCol };
}

// ---------------------------------------------------------------------------
// Formulas, names and merged ranges
// ---------------------------------------------------------------------------

describe('M5 structural formulas', () => {
  it('rewrites relative, absolute, mixed and cross-sheet references', () => {
    const { workbook, sheet } = workbookWithSheet();
    putFormula(sheet, 1, 0, 'A2+$B$2+C$3+$D3');
    putFormula(sheet, 1, 1, 'Sheet1!A2');
    const edit = insertRow(sheet.id, 1);
    expect(transformFormulaReferences('A2+$B$2+C$3+$D3', edit, { workbook }, sheet.id)).toBe(
      'A3+$B$3+C$4+$D4',
    );
    expect(transformFormulaReferences('Sheet1!A2', edit, { workbook }, sheet.id)).toBe('Sheet1!A3');
  });

  it('rewrites a cross-sheet reference to the edited sheet from another sheet', () => {
    const { workbook, sheet } = workbookWithSheet();
    const second = createSheet('Sheet2', 'xl/worksheets/sheet2.xml');
    workbook.sheets.push(second);
    putFormula(second, 0, 0, 'Sheet1!A1+SUM(Sheet1!A1:B2)');
    const plan = planRangeTransform(workbook, insertRow(sheet.id, 1));
    expect(plan.formulaRewrites.map((rewrite) => rewrite.after)).toEqual([
      'Sheet1!A1+SUM(Sheet1!A1:B3)',
    ]);
    expect(plan.cellMoves).toHaveLength(1);
    expect(plan.cellMoves[0]?.sheetId).toBe(second.id);
    expect(plan.cellMoves[0]?.from).toEqual({ row: 0, col: 0 });
    expect(plan.cellMoves[0]?.to).toEqual({ row: 0, col: 0 });
  });

  it('collapses a reference the edit deletes to #REF! and keeps the row', () => {
    const { workbook, sheet } = workbookWithSheet();
    putFormula(sheet, 1, 0, 'A1');
    const collapsed = planRangeTransform(workbook, deleteRow(sheet.id, 0));
    expect(
      transformFormulaReferences('A1+B2', deleteRow(sheet.id, 0), { workbook }, sheet.id),
    ).toBe('#REF!+B1');
    expect(collapsed.formulaRewrites).toHaveLength(1);
  });

  it('renumbers the row dimension map without touching widths', () => {
    const { workbook, sheet } = workbookWithSheet();
    sheet.rowDimensions.set(2, { height: 30, hidden: true });
    sheet.rowDimensions.set(4, { height: 12 });
    sheet.columnDimensions.set(1, { width: 20, hidden: true });
    const plan = planRangeTransform(workbook, insertRow(sheet.id, 1));
    expect(plan.rowDimensions.after).toEqual([
      [3, { height: 30, hidden: true }],
      [5, { height: 12 }],
    ]);
    expect(plan.columnDimensions.after).toEqual([[1, { width: 20, hidden: true }]]);
    expect(
      new Map(transformDimensionMap(sheet.rowDimensions, deleteRow(sheet.id, 2), 'row')),
    ).toEqual(new Map([[3, { height: 12 }]]));
  });

  it('moves comment and hyperlink keys with their cells and drops deleted ones', () => {
    const { workbook, sheet } = workbookWithSheet();
    sheet.comments.set(cellKey(1, 0), { text: 'note', threaded: false });
    sheet.comments.set(cellKey(4, 2), { text: 'gone', threaded: false });
    sheet.hyperlinks.set(cellKey(1, 1), { target: 'https://example.com' });
    const plan = planRangeTransform(workbook, deleteRow(sheet.id, 4));
    expect(plan.comments.after).toEqual([[cellKey(1, 0), { text: 'note', threaded: false }]]);
    expect(plan.hyperlinks.after).toEqual([[cellKey(1, 1), { target: 'https://example.com' }]]);
    expect(new Map(transformAttachmentMap(sheet.comments, insertRow(sheet.id, 1)))).toEqual(
      new Map([
        [cellKey(2, 0), { text: 'note', threaded: false }],
        [cellKey(5, 2), { text: 'gone', threaded: false }],
      ]),
    );
  });
});

describe('M5 structural merges and names', () => {
  it('shifts, expands, clips and drops merged ranges', () => {
    expect(transformMergeRange(range(0, 0, 1, 1), insertRow('s', 1))).toEqual(range(0, 0, 2, 1));
    expect(transformMergeRange(range(0, 0, 1, 1), deleteRow('s', 0))).toEqual(range(0, 0, 0, 1));
    expect(transformMergeRange(range(0, 0, 0, 1), deleteRow('s', 0))).toBeNull();
  });

  it('blocks a cell edit that would change part of a merged cell', () => {
    const { workbook, sheet } = workbookWithSheet();
    sheet.merges.push({ range: range(0, 0, 1, 1) });
    const plan = planRangeTransform(workbook, {
      kind: 'insert',
      sheetId: sheet.id,
      axis: 'row',
      index: 1,
      count: 1,
      band: [0, 0],
    });
    expect(plan.block?.code).toBe('merge-split');
    expect(plan.block?.message).toBe('Cannot change part of a merged cell.');
  });

  it('moves a merge only when the cell edit covers the whole merged range', () => {
    const merge = range(0, 0, 0, 1);
    const partial: StructuralEdit = {
      kind: 'insert',
      sheetId: 's',
      axis: 'column',
      index: 0,
      count: 1,
      band: [0, 0],
    };
    expect(transformMergeRange(merge, partial)).toEqual(merge);
    const whole: StructuralEdit = { ...partial, count: 2 };
    expect(transformMergeRange(merge, whole)).toEqual(range(0, 2, 0, 3));
    const wholeDelete: StructuralEdit = {
      kind: 'delete',
      sheetId: 's',
      axis: 'column',
      index: 0,
      count: 2,
      band: [0, 0],
    };
    expect(transformMergeRange(merge, wholeDelete)).toBeNull();
  });

  it('blocks a cell delete that removes only part of a merge', () => {
    const { workbook, sheet } = workbookWithSheet();
    sheet.merges.push({ range: range(0, 0, 1, 0) });
    const plan = planRangeTransform(workbook, {
      kind: 'delete',
      sheetId: sheet.id,
      axis: 'row',
      index: 0,
      count: 1,
      band: [0, 0],
    });
    expect(plan.block?.code).toBe('merge-split');
  });

  it('rewrites workbook and sheet-scoped defined names', () => {
    const { workbook, sheet } = workbookWithSheet();
    const second = createSheet('Sheet2', 'xl/worksheets/sheet2.xml');
    workbook.sheets.push(second);
    workbook.names.push({ name: 'TaxRange', formula: 'Sheet1!$A$1:$B$2', sheetId: null });
    second.names.push({ name: 'LocalAmount', formula: '$C$3', sheetId: second.id });
    const firstPlan = planRangeTransform(workbook, insertRow(sheet.id, 1));
    expect(firstPlan.nameRewrites.map((rewrite) => rewrite.after)).toEqual(['Sheet1!$A$1:$B$3']);
    const secondPlan = planRangeTransform(workbook, insertRow(second.id, 1));
    expect(secondPlan.nameRewrites.map((rewrite) => rewrite.after)).toEqual(['$C$4']);
  });
});

// ---------------------------------------------------------------------------
// CF17 cf-structural
// ---------------------------------------------------------------------------

describe('CF17 cf-structural', () => {
  const CF_RULE =
    '<cfRule type="cellIs" dxfId="3" priority="1" operator="greaterThan"><formula>A2&gt;10</formula></cfRule>';

  it('transforms the rule range, its sqref and its formula references', () => {
    const { workbook, sheet } = workbookWithSheet();
    sheet.conditionalFormats.push({
      range: range(0, 0, 9, 0),
      rawXml: `<conditionalFormatting sqref="A1:A10" priority="1">${CF_RULE}</conditionalFormatting>`,
    });
    const plan = planRangeTransform(workbook, insertRow(sheet.id, 1));
    expect(plan.conditionalFormats.after).toHaveLength(1);
    const record = plan.conditionalFormats.after[0]!;
    expect(record.range).toEqual(range(0, 0, 10, 0));
    expect(record.rawXml).toContain('sqref="A1:A11"');
    expect(record.rawXml).toContain('<formula>A3&gt;10</formula>');
    // Priority and every unknown attribute survive untouched.
    expect(record.rawXml).toContain('priority="1"');
    expect(record.rawXml).toContain('dxfId="3"');
  });

  it('removes a rule whose range collapses and preserves the remaining order', () => {
    const { workbook, sheet } = workbookWithSheet();
    sheet.conditionalFormats.push(
      { range: range(1, 0, 1, 0), rawXml: '<cfRule priority="1"/>' },
      { range: range(0, 2, 5, 2), rawXml: '<cfRule priority="2"/>' },
    );
    const plan = planRangeTransform(workbook, deleteRow(sheet.id, 1));
    expect(plan.conditionalFormats.after).toHaveLength(1);
    expect(plan.conditionalFormats.after[0]?.rawXml).toBe('<cfRule priority="2"/>');
    expect(plan.conditionalFormats.after[0]?.range).toEqual(range(0, 2, 4, 2));
  });

  it('transforms validation ranges and formulas through the same path', () => {
    const { workbook, sheet } = workbookWithSheet();
    const record: ValidationRecord = {
      range: range(0, 0, 4, 0),
      rawXml:
        '<dataValidation type="list" sqref="A1:A5"><formula1>$B$1:$B$5</formula1></dataValidation>',
    };
    sheet.validation.push(record);
    const plan = planRangeTransform(workbook, insertColumn(sheet.id, 1));
    const next = plan.validation.after[0]!;
    expect(next.range).toEqual(range(0, 0, 4, 0));
    expect(next.rawXml).toContain('sqref="A1:A5"');
    expect(next.rawXml).toContain('$C$1:$C$5');
  });

  it('drops a CF record whose range collapses but keeps the rule list stable', () => {
    const { workbook, sheet } = workbookWithSheet();
    const kept = transformConditionalFormatRecord(
      { range: range(0, 0, 0, 0), rawXml: CF_RULE },
      deleteRow(sheet.id, 0),
      { workbook },
      sheet.id,
    );
    expect(kept).toBeNull();
  });
});

// ---------------------------------------------------------------------------
// F11 filter-structural
// ---------------------------------------------------------------------------

describe('F11 filter-structural', () => {
  function filteredSheet() {
    const { workbook, sheet } = workbookWithSheet();
    for (let row = 0; row < 6; row += 1) {
      setCellSnapshot(sheet, row, 0, {
        value: textValue(`r${row}`),
        formula: null,
        formulaDerived: false,
        savedResult: null,
        styleId: 0,
        richText: null,
      });
      setCellSnapshot(sheet, row, 1, {
        value: numberValue(row * 10),
        formula: null,
        formulaDerived: false,
        savedResult: null,
        styleId: 0,
        richText: null,
      });
    }
    sheet.autoFilter = createAutoFilter(sheet.id, range(0, 0, 5, 2), [
      [
        1,
        {
          kind: 'conditions',
          join: 'and',
          conditions: [{ kind: 'number', operator: '>', value: 20 }],
        },
      ],
      [2, { kind: 'values', selected: [textValue('x')], includeBlanks: false }],
    ]);
    return { workbook, sheet };
  }

  it('shifts the range when rows are inserted above it and keeps the offsets', () => {
    const { workbook, sheet } = filteredSheet();
    const plan = planRangeTransform(workbook, insertRow(sheet.id, 1));
    expect(plan.autoFilter.after?.range).toEqual(range(0, 0, 6, 2));
    expect([...(plan.autoFilter.after?.columns.keys() ?? [])]).toEqual([1, 2]);
  });

  it('moves a criterion with its column when a column is inserted inside the range', () => {
    const { workbook, sheet } = filteredSheet();
    const plan = planRangeTransform(workbook, insertColumn(sheet.id, 1));
    expect(plan.autoFilter.after?.range).toEqual(range(0, 0, 5, 3));
    expect([...(plan.autoFilter.after?.columns.keys() ?? [])]).toEqual([2, 3]);
  });

  it('keeps the offsets when a column is inserted before the range', () => {
    const { workbook, sheet } = filteredSheet();
    const plan = planRangeTransform(workbook, insertColumn(sheet.id, 0));
    expect(plan.autoFilter.after?.range).toEqual(range(0, 1, 5, 3));
    expect([...(plan.autoFilter.after?.columns.keys() ?? [])]).toEqual([1, 2]);
  });

  it('drops the criterion whose column is deleted and shifts the rest', () => {
    const { workbook, sheet } = filteredSheet();
    const plan = planRangeTransform(workbook, {
      kind: 'delete',
      sheetId: sheet.id,
      axis: 'column',
      index: 1,
      count: 1,
    });
    expect(plan.autoFilter.after?.range).toEqual(range(0, 0, 5, 1));
    expect([...(plan.autoFilter.after?.columns.keys() ?? [])]).toEqual([1]);
  });

  it('refuses a cell edit that would split only part of the filter rectangle', () => {
    const { workbook, sheet } = filteredSheet();
    const plan = planRangeTransform(workbook, {
      kind: 'insert',
      sheetId: sheet.id,
      axis: 'row',
      index: 2,
      count: 1,
      band: [0, 0],
    });
    expect(plan.block?.code).toBe('filter-split');
  });
});

// ---------------------------------------------------------------------------
// F12 filter-header-delete
// ---------------------------------------------------------------------------

describe('F12 filter-header-delete', () => {
  it('clears the filter and reports the filter-hidden rows to restore', () => {
    const { sheet } = workbookWithSheet();
    for (let row = 0; row < 5; row += 1) {
      setCellSnapshot(sheet, row, 1, {
        value: numberValue(row === 2 || row === 4 ? 5 : 15),
        formula: null,
        formulaDerived: false,
        savedResult: null,
        styleId: 0,
        richText: null,
      });
    }
    sheet.autoFilter = createAutoFilter(sheet.id, range(0, 0, 4, 1), [
      [
        1,
        {
          kind: 'conditions',
          join: 'and',
          conditions: [{ kind: 'number', operator: '>', value: 10 }],
        },
      ],
    ]);
    sheet.rowDimensions.set(2, { hidden: true });
    sheet.rowDimensions.set(4, { hidden: true });
    expect([...evaluateAutoFilter(sheet, sheet.autoFilter).filterHidden]).toEqual([2, 4]);

    const result = planAutoFilterTransform(sheet, deleteRow(sheet.id, 0));
    expect(result.after).toBeNull();
    // The surviving rows are now 1 and 3, so those flags must be cleared.
    expect(result.restoredHiddenRows).toEqual([1, 3]);
  });

  it('clears a filter that collapses below two rows', () => {
    const { sheet } = workbookWithSheet();
    sheet.autoFilter = createAutoFilter(sheet.id, range(0, 0, 4, 1));
    expect(planAutoFilterTransform(sheet, deleteRow(sheet.id, 1, 4)).after).toBeNull();
  });

  it('leaves a filter entirely below the edit alone', () => {
    const { workbook, sheet } = workbookWithSheet();
    sheet.autoFilter = createAutoFilter(sheet.id, range(3, 0, 8, 1), [
      [0, { kind: 'values', selected: [textValue('a')], includeBlanks: false }],
    ]);
    const plan = planRangeTransform(workbook, insertRow(sheet.id, 10));
    expect(sheet.autoFilter?.range).toEqual(range(3, 0, 8, 1));
    expect(plan.autoFilter.after).toEqual(sheet.autoFilter);
    expect(plan.changed).toBe(false);
  });
});

// ---------------------------------------------------------------------------
// Cell-shift semantics
// ---------------------------------------------------------------------------

describe('M5 cell insert/delete shift semantics', () => {
  it('maps a cell address through a banded shift', () => {
    const edit: StructuralEdit = {
      kind: 'insert',
      sheetId: 's',
      axis: 'column',
      index: 1,
      count: 1,
      band: [0, 1],
    };
    expect(transformCellAddress(0, 1, edit)).toEqual({ row: 0, col: 2 });
    expect(transformCellAddress(0, 0, edit)).toEqual({ row: 0, col: 0 });
    // Rows outside the band are untouched even for the same column.
    expect(transformCellAddress(2, 1, edit)).toEqual({ row: 2, col: 1 });
  });

  it('shifts only the selected rows for Insert Cells & Shift Right', () => {
    const { workbook, sheet } = workbookWithSheet();
    setCellSnapshot(sheet, 0, 0, {
      value: textValue('inside'),
      formula: null,
      formulaDerived: false,
      savedResult: null,
      styleId: 0,
      richText: null,
    });
    setCellSnapshot(sheet, 5, 0, {
      value: textValue('outside'),
      formula: null,
      formulaDerived: false,
      savedResult: null,
      styleId: 0,
      richText: null,
    });
    const plan = planRangeTransform(workbook, {
      kind: 'insert',
      sheetId: sheet.id,
      axis: 'column',
      index: 0,
      count: 1,
      band: [0, 1],
    });
    expect(plan.cellMoves.map((move) => [move.from, move.to])).toEqual([
      [
        { row: 0, col: 0 },
        { row: 0, col: 1 },
      ],
    ]);
    expect(plan.rowDimensions.after).toEqual(plan.rowDimensions.before);
  });
});
