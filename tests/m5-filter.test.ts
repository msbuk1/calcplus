import { describe, expect, it } from 'vitest';
import fc from 'fast-check';
import {
  MAX_CONDITIONS_PER_COLUMN,
  clearFilterLayers,
  cloneAutoFilter,
  collectValueChecklist,
  compareFilterValues,
  computeHiddenLayers,
  createAutoFilter,
  evaluateAutoFilter,
  evaluateColumnFilter,
  evaluateCondition,
  isHeaderRow,
  planAutoFilterChange,
  planFilteredSort,
  restoreHiddenLayers,
  searchChecklist,
  visibleRowPredicate,
  visibleRowsInRange,
  type AutoFilter,
  type ColumnOffset,
  type ColumnFilter,
  type Condition,
  type DateComparisonOperator,
  type NumberComparisonOperator,
  type TextOperator,
  type ValueFilter,
} from '@model/autofilter';
import { cellKey, setCellSnapshot, type Cell, type GridRange, type Sheet } from '@model/types';
import { createEmptyWorkbook, createSheet } from '@model/workbook';
import {
  BLANK,
  booleanValue,
  errorValue,
  formatCellValue,
  numberValue,
  textValue,
  type CellValue,
} from '@model/value';

/**
 * Model-level AutoFilter conformance: sort-autofilter §4 fixtures F1-F10.
 * F11-F18 need the range-transform, OOXML and command layers and are covered
 * in a later pass. Everything here is pure: no DOM, no I/O, no clock.
 */

function cell(value: CellValue): Cell {
  return {
    value,
    formula: null,
    formulaDerived: false,
    savedResult: null,
    styleId: 0,
    richText: null,
    stale: false,
  };
}

function sheetOf(rows: readonly (readonly CellValue[])[]): Sheet {
  const sheet = createSheet('Sheet1', 'xl/worksheets/sheet1.xml');
  rows.forEach((values, row) => {
    values.forEach((value, col) => setCellSnapshot(sheet, row, col, cell(value)));
  });
  return sheet;
}

function filterFor(
  sheet: Sheet,
  range: GridRange,
  columns: readonly (readonly [ColumnOffset, ColumnFilter])[],
): AutoFilter {
  return createAutoFilter(sheet.id, range, columns);
}

function valuesFilter(selected: readonly CellValue[], includeBlanks = false): ValueFilter {
  return { kind: 'values', selected, includeBlanks };
}

function conditionsFilter(join: 'and' | 'or', ...conditions: Condition[]): ColumnFilter {
  return { kind: 'conditions', join, conditions };
}

function textCondition(operator: TextOperator, value: string): Condition {
  return { kind: 'text', operator, value };
}

function numCondition(operator: NumberComparisonOperator, value: number): Condition {
  return { kind: 'number', operator, value };
}

function numBetween(operator: 'between' | 'not-between', from: number, to: number): Condition {
  return { kind: 'number', operator, from, to };
}

function dateCondition(operator: DateComparisonOperator, value: number): Condition {
  return { kind: 'date', operator, value };
}

function dateBetween(operator: 'between' | 'not-between', from: number, to: number): Condition {
  return { kind: 'date', operator, from, to };
}

function rangeOf(endRow: number, endCol = 0, startRow = 0, startCol = 0): GridRange {
  return { startRow, startCol, endRow, endCol };
}

/** Data rows (header excluded) that the filter leaves visible, in order. */
function visibleDataRows(sheet: Sheet, filter: AutoFilter): number[] {
  const hidden = evaluateAutoFilter(sheet, filter).filterHidden;
  const rows: number[] = [];
  for (let row = filter.range.startRow + 1; row <= filter.range.endRow; row += 1) {
    if (!hidden.has(row)) rows.push(row);
  }
  return rows;
}

function visibleWithCondition(sheet: Sheet, range: GridRange, condition: Condition): number[] {
  return visibleDataRows(sheet, filterFor(sheet, range, [[0, conditionsFilter('and', condition)]]));
}

describe('F1 filter-values', () => {
  const sheet = sheetOf([
    [textValue('Value')],
    [numberValue(10)],
    [textValue('10')],
    [numberValue(10)],
    [numberValue(3)],
    [booleanValue(true)],
    [errorValue('#N/A')],
    [BLANK],
  ]);
  const range = rangeOf(7);
  const filter = filterFor(sheet, range, []);

  it('lists distinct typed values in comparator order with blanks separate', () => {
    const checklist = collectValueChecklist(sheet, filter, 0);
    expect(checklist.hasBlanks).toBe(true);
    expect(checklist.rowCount).toBe(7);
    expect(checklist.entries.map((entry) => entry.display)).toEqual([
      '3',
      '10',
      '10',
      'TRUE',
      '#N/A',
    ]);
    expect(checklist.entries.map((entry) => entry.values[0]?.type)).toEqual([
      'number',
      'number',
      'text',
      'boolean',
      'error',
    ]);
    expect(checklist.entries.some((entry) => entry.display === 'Value')).toBe(false);
  });

  it('selects a typed value only, so number 10 is not text "10"', () => {
    const selected = filterFor(sheet, range, [[0, valuesFilter([numberValue(10)])]]);
    expect(visibleDataRows(sheet, selected)).toEqual([1, 3]);

    const text = filterFor(sheet, range, [[0, valuesFilter([textValue('10')])]]);
    expect(visibleDataRows(sheet, text)).toEqual([2]);
  });

  it('selects errors as checklist values', () => {
    const selected = filterFor(sheet, range, [[0, valuesFilter([errorValue('#N/A')])]]);
    expect(visibleDataRows(sheet, selected)).toEqual([6]);
  });

  it('Select All is the whole checklist plus blanks', () => {
    const checklist = collectValueChecklist(sheet, filter, 0);
    const everything = checklist.entries.flatMap((entry) => entry.values);
    const selectAll = filterFor(sheet, range, [[0, valuesFilter(everything, true)]]);
    expect(visibleDataRows(sheet, selectAll)).toEqual([1, 2, 3, 4, 5, 6, 7]);
  });

  it('merges values that format identically into one entry selecting both', () => {
    const formatted = sheetOf([
      [textValue('Amount')],
      [numberValue(1.001)],
      [numberValue(1.002)],
      [numberValue(2.5)],
    ]);
    const formattedRange = rangeOf(3);
    const checklist = collectValueChecklist(
      formatted,
      filterFor(formatted, formattedRange, []),
      0,
      {
        displayText: (_cell, value) =>
          value.type === 'number' ? value.value.toFixed(2) : formatCellValue(value),
      },
    );
    expect(checklist.entries.map((entry) => entry.display)).toEqual(['1.00', '2.50']);
    expect(checklist.entries[0]?.values).toHaveLength(2);

    const firstEntry = checklist.entries[0];
    expect(firstEntry).toBeDefined();
    const grouped = filterFor(formatted, formattedRange, [
      [0, valuesFilter(firstEntry?.values ?? [])],
    ]);
    expect(visibleDataRows(formatted, grouped)).toEqual([1, 2]);
  });

  it('filters the checklist search case-insensitively', () => {
    const checklist = collectValueChecklist(sheet, filter, 0);
    expect(searchChecklist(checklist.entries, '10').map((entry) => entry.display)).toEqual([
      '10',
      '10',
    ]);
    expect(searchChecklist(checklist.entries, 'true').map((entry) => entry.display)).toEqual([
      'TRUE',
    ]);
    expect(searchChecklist(checklist.entries, 'zzz')).toEqual([]);
    expect(searchChecklist(checklist.entries, '')).toHaveLength(checklist.entries.length);
  });

  it('evaluates the current value of a formula cell and never recalculates', () => {
    const formulaSheet = sheetOf([[textValue('Value')]]);
    setCellSnapshot(formulaSheet, 1, 0, {
      value: numberValue(5),
      formula: 'A1*2',
      formulaDerived: false,
      savedResult: numberValue(5),
      styleId: 0,
      richText: null,
      stale: false,
    });
    const formulaRange = rangeOf(1);
    const selected = filterFor(formulaSheet, formulaRange, [[0, valuesFilter([numberValue(5)])]]);
    expect(visibleDataRows(formulaSheet, selected)).toEqual([1]);
  });
});

describe('F2 filter-text', () => {
  const sheet = sheetOf([
    [textValue('Text')],
    [textValue('Apple')],
    [textValue('apricot')],
    [textValue('Banana')],
    [textValue('a-b')],
    [textValue('a*b')],
    [textValue('ab')],
    [numberValue(42)],
    [BLANK],
  ]);
  const range = rangeOf(8);

  it('matches equals case-insensitively', () => {
    expect(visibleWithCondition(sheet, range, textCondition('equals', 'APPLE'))).toEqual([1]);
  });

  it('supports ?, * and the ~ escape', () => {
    expect(visibleWithCondition(sheet, range, textCondition('equals', 'a?ple'))).toEqual([1]);
    expect(visibleWithCondition(sheet, range, textCondition('equals', 'a~*b'))).toEqual([5]);
    expect(visibleWithCondition(sheet, range, textCondition('begins-with', 'a*'))).toEqual([
      1, 2, 4, 5, 6,
    ]);
  });

  it('implements begins with, ends with, contains and their negations', () => {
    expect(visibleWithCondition(sheet, range, textCondition('begins-with', 'ap'))).toEqual([1, 2]);
    expect(visibleWithCondition(sheet, range, textCondition('ends-with', 'na'))).toEqual([3]);
    expect(visibleWithCondition(sheet, range, textCondition('contains', 'b'))).toEqual([
      3, 4, 5, 6,
    ]);
    expect(visibleWithCondition(sheet, range, textCondition('does-not-contain', 'an'))).toEqual([
      1, 2, 4, 5, 6,
    ]);
    expect(visibleWithCondition(sheet, range, textCondition('not-equals', 'Apple'))).toEqual([
      2, 3, 4, 5, 6,
    ]);
  });

  it('never matches a number or a blank, even with a negated operator', () => {
    expect(visibleWithCondition(sheet, range, textCondition('contains', '4'))).toEqual([]);
    expect(visibleWithCondition(sheet, range, textCondition('not-equals', ''))).toEqual([
      1, 2, 3, 4, 5, 6,
    ]);
    expect(visibleWithCondition(sheet, range, textCondition('contains', ''))).toEqual([
      1, 2, 3, 4, 5, 6,
    ]);
    expect(evaluateCondition(textCondition('equals', '42'), numberValue(42))).toBe(false);
    expect(evaluateCondition(textCondition('not-equals', 'x'), BLANK)).toBe(false);
  });
});

describe('F3 filter-number', () => {
  const sheet = sheetOf([
    [textValue('Number')],
    [numberValue(5)],
    [numberValue(10)],
    [numberValue(15)],
    [numberValue(20)],
    [textValue('10')],
    [BLANK],
  ]);
  const range = rangeOf(6);

  it('implements the comparison operators over numbers only', () => {
    expect(visibleWithCondition(sheet, range, numCondition('>', 10))).toEqual([3, 4]);
    expect(visibleWithCondition(sheet, range, numCondition('>=', 10))).toEqual([2, 3, 4]);
    expect(visibleWithCondition(sheet, range, numCondition('<', 10))).toEqual([1]);
    expect(visibleWithCondition(sheet, range, numCondition('<=', 10))).toEqual([1, 2]);
    expect(visibleWithCondition(sheet, range, numCondition('=', 15))).toEqual([3]);
    expect(visibleWithCondition(sheet, range, numCondition('<>', 15))).toEqual([1, 2, 4]);
  });

  it('implements inclusive between and not between', () => {
    expect(visibleWithCondition(sheet, range, numBetween('between', 10, 15))).toEqual([2, 3]);
    expect(visibleWithCondition(sheet, range, numBetween('between', 15, 10))).toEqual([2, 3]);
    expect(visibleWithCondition(sheet, range, numBetween('not-between', 10, 15))).toEqual([1, 4]);
  });

  it('ignores text that looks numeric and blanks', () => {
    expect(visibleWithCondition(sheet, range, numCondition('<', 100))).toEqual([1, 2, 3, 4]);
    expect(evaluateCondition(numCondition('=', 10), textValue('10'))).toBe(false);
    expect(evaluateCondition(numCondition('=', 10), booleanValue(true))).toBe(false);
  });
});

describe('F4 filter-date', () => {
  function dateSheet(date1904: boolean): Sheet {
    const workbook = createEmptyWorkbook();
    workbook.date1904 = date1904;
    const sheet = createSheet(
      date1904 ? 'D1904' : 'D1900',
      `xl/worksheets/dates${date1904 ? '1904' : '1900'}.xml`,
    );
    workbook.sheets.push(sheet);
    const rows: readonly (readonly CellValue[])[] = [
      [textValue('Date')],
      [numberValue(43831)],
      [numberValue(43832)],
      [numberValue(43831.5)],
      [numberValue(44196)],
      [textValue('43831')],
    ];
    rows.forEach((values, row) => {
      values.forEach((value, col) => setCellSnapshot(sheet, row, col, cell(value)));
    });
    return sheet;
  }

  const sheet = dateSheet(false);
  const range = rangeOf(5);

  it('compares date conditions as serial numbers', () => {
    expect(visibleWithCondition(sheet, range, dateCondition('equals', 43831))).toEqual([1]);
    expect(visibleWithCondition(sheet, range, dateCondition('not-equals', 43831))).toEqual([
      2, 3, 4,
    ]);
    expect(visibleWithCondition(sheet, range, dateCondition('before', 43832))).toEqual([1, 3]);
    expect(visibleWithCondition(sheet, range, dateCondition('after', 43831))).toEqual([2, 3, 4]);
    expect(visibleWithCondition(sheet, range, dateBetween('between', 43831, 43832))).toEqual([
      1, 2, 3,
    ]);
    expect(visibleWithCondition(sheet, range, dateBetween('not-between', 43831, 43832))).toEqual([
      4,
    ]);
  });

  it('gives the same result in a 1900 and a 1904 workbook', () => {
    const modern = dateSheet(false);
    const mac = dateSheet(true);
    for (const condition of [
      dateCondition('equals', 43831),
      dateCondition('before', 43832),
      dateBetween('between', 43831, 43832),
    ]) {
      expect(visibleWithCondition(modern, range, condition)).toEqual(
        visibleWithCondition(mac, range, condition),
      );
    }
  });
});

describe('F5 filter-and-or', () => {
  const sheet = sheetOf([
    [textValue('Number')],
    [numberValue(5)],
    [numberValue(10)],
    [numberValue(15)],
    [numberValue(20)],
    [numberValue(25)],
  ]);
  const range = rangeOf(5);

  it('joins two conditions with AND', () => {
    const filter = filterFor(sheet, range, [
      [0, conditionsFilter('and', numCondition('>', 5), numCondition('<', 20))],
    ]);
    expect(visibleDataRows(sheet, filter)).toEqual([2, 3]);
  });

  it('joins two conditions with OR', () => {
    const filter = filterFor(sheet, range, [
      [0, conditionsFilter('or', numCondition('<', 10), numCondition('>', 20))],
    ]);
    expect(visibleDataRows(sheet, filter)).toEqual([1, 5]);
  });

  it('caps a column at two conditions', () => {
    expect(MAX_CONDITIONS_PER_COLUMN).toBe(2);
  });
});

describe('F6 filter-columns', () => {
  const sheet = sheetOf([
    [textValue('Region'), textValue('Amount')],
    [textValue('x'), numberValue(10)],
    [textValue('x'), numberValue(20)],
    [textValue('y'), numberValue(10)],
    [textValue('y'), numberValue(20)],
    [textValue('y'), numberValue(20)],
  ]);
  const range = rangeOf(5, 1);

  it('combines filtered columns with AND', () => {
    const both = filterFor(sheet, range, [
      [0, valuesFilter([textValue('x')])],
      [1, valuesFilter([numberValue(20)])],
    ]);
    expect(visibleDataRows(sheet, both)).toEqual([2]);
  });

  it('combines a value checklist with a condition on another column', () => {
    const mixed = filterFor(sheet, range, [
      [0, valuesFilter([textValue('y')])],
      [1, conditionsFilter('and', numCondition('>=', 20))],
    ]);
    expect(visibleDataRows(sheet, mixed)).toEqual([4, 5]);
  });

  it('keeps the header visible even when its cells fail every column', () => {
    const filter = filterFor(sheet, range, [
      [0, valuesFilter([textValue('x')])],
      [1, valuesFilter([numberValue(20)])],
    ]);
    const evaluation = evaluateAutoFilter(sheet, filter);
    expect(evaluation.filterHidden.has(0)).toBe(false);
    expect(isHeaderRow(filter, 0)).toBe(true);
    expect(isHeaderRow(filter, 1)).toBe(false);
  });
});

describe('F7 filter-blanks', () => {
  const sheet = sheetOf([
    [textValue('Value')],
    [numberValue(1)],
    [BLANK],
    [numberValue(2)],
    [BLANK],
    [textValue('x')],
  ]);
  // Row 6 has no cell record at all; it is blank for filtering too.
  const range = rangeOf(6);
  const filter = filterFor(sheet, range, []);
  const checklist = collectValueChecklist(sheet, filter, 0);

  it('reports blanks and treats an absent cell as blank', () => {
    expect(checklist.hasBlanks).toBe(true);
    expect(checklist.rowCount).toBe(6);
  });

  it('hides and shows blanks with includeBlanks', () => {
    const withoutBlanks = filterFor(sheet, range, [
      [0, valuesFilter([numberValue(1), numberValue(2), textValue('x')], false)],
    ]);
    expect(visibleDataRows(sheet, withoutBlanks)).toEqual([1, 3, 5]);

    const withBlanks = filterFor(sheet, range, [
      [0, valuesFilter([numberValue(1), numberValue(2), textValue('x')], true)],
    ]);
    expect(visibleDataRows(sheet, withBlanks)).toEqual([1, 2, 3, 4, 5, 6]);
  });

  it('shows only blanks when Blanks is the only selection', () => {
    const blanksOnly = filterFor(sheet, range, [[0, valuesFilter([], true)]]);
    expect(visibleDataRows(sheet, blanksOnly)).toEqual([2, 4, 6]);
  });

  it('never matches a blank with a condition', () => {
    expect(visibleWithCondition(sheet, range, textCondition('contains', 'x'))).toEqual([5]);
  });
});

describe('F8 filter-errors', () => {
  const sheet = sheetOf([
    [textValue('Value')],
    [numberValue(1)],
    [errorValue('#N/A')],
    [errorValue('#DIV/0!')],
    [textValue('x')],
    [BLANK],
  ]);
  const range = rangeOf(5);

  it('lists errors as values and selects them individually', () => {
    const checklist = collectValueChecklist(sheet, filterFor(sheet, range, []), 0);
    expect(checklist.entries.map((entry) => entry.display)).toEqual(['1', 'x', '#N/A', '#DIV/0!']);

    const selected = filterFor(sheet, range, [[0, valuesFilter([errorValue('#N/A')])]]);
    expect(visibleDataRows(sheet, selected)).toEqual([2]);
  });

  it('never lets an error match a condition', () => {
    expect(visibleWithCondition(sheet, range, numCondition('>', 0))).toEqual([1]);
    expect(visibleWithCondition(sheet, range, dateCondition('after', 0))).toEqual([1]);
    expect(visibleWithCondition(sheet, range, textCondition('contains', ''))).toEqual([4]);
  });

  it('reports an unsupported criterion as a typed outcome', () => {
    expect(
      evaluateColumnFilter({ kind: 'unsupported', raw: '<colourFilter/>' }, numberValue(1)),
    ).toBe('unsupported');
    const unsupported = filterFor(sheet, range, [
      [0, { kind: 'unsupported', raw: '<top10 percent="1"/>' }],
    ]);
    const evaluation = evaluateAutoFilter(sheet, unsupported);
    expect(evaluation.unsupportedColumns).toEqual([0]);
    expect(evaluation.filterHidden.size).toBe(0);
  });
});

describe('F9 filter-hidden-layers', () => {
  const sheet = sheetOf([
    [numberValue(20)],
    [numberValue(10)],
    [numberValue(20)],
    [numberValue(10)],
    [numberValue(30)],
  ]);
  const range = rangeOf(4);
  const filter = filterFor(sheet, range, [[0, valuesFilter([numberValue(10)])]]);
  const manual = new Set([2, 5]);

  it('keeps manual-hidden and filter-hidden rows as independent layers', () => {
    const layers = computeHiddenLayers(sheet, filter, manual);
    expect([...layers.filterHidden].sort((a, b) => a - b)).toEqual([2, 4]);
    expect([...layers.manualHidden].sort((a, b) => a - b)).toEqual([2, 5]);
    expect([...layers.effectiveHidden].sort((a, b) => a - b)).toEqual([2, 4, 5]);
    expect(layers.filterHidden.has(0)).toBe(false);
  });

  it('clearing removes only the filter layer', () => {
    const layers = computeHiddenLayers(sheet, filter, manual);
    const cleared = clearFilterLayers(layers);
    expect(cleared.filterHidden.size).toBe(0);
    expect([...cleared.effectiveHidden].sort((a, b) => a - b)).toEqual([2, 5]);
    expect([...cleared.manualHidden].sort((a, b) => a - b)).toEqual([2, 5]);

    const noFilter = computeHiddenLayers(sheet, null, manual);
    expect(noFilter.filterHidden.size).toBe(0);
    expect([...noFilter.effectiveHidden].sort((a, b) => a - b)).toEqual([2, 5]);
  });

  it('rebuilds the layers from saved hidden flags', () => {
    const restored = restoreHiddenLayers(sheet, filter, [2, 4, 5]);
    expect([...restored.filterHidden].sort((a, b) => a - b)).toEqual([2, 4]);
    expect([...restored.manualHidden].sort((a, b) => a - b)).toEqual([5]);
    expect([...restored.effectiveHidden].sort((a, b) => a - b)).toEqual([2, 4, 5]);

    // A row the saved file left visible but the criteria now fail becomes
    // filter-hidden, not manual.
    const stricter = restoreHiddenLayers(sheet, filter, [2, 5]);
    expect([...stricter.manualHidden]).toEqual([5]);
    expect([...stricter.filterHidden].sort((a, b) => a - b)).toEqual([2, 4]);
  });

  it('records the effective hidden-row changes for one undo step', () => {
    const applied = planAutoFilterChange(sheet, null, filter, manual);
    expect(applied.sheetId).toBe(sheet.id);
    expect(applied.before).toBeNull();
    expect(applied.after).toBe(filter);
    expect(applied.hiddenChanges).toEqual([{ row: 4, before: false, after: true }]);

    const cleared = planAutoFilterChange(sheet, filter, null, manual);
    expect(cleared.hiddenChanges).toEqual([{ row: 4, before: true, after: false }]);

    const unchanged = planAutoFilterChange(sheet, filter, filter, manual);
    expect(unchanged.hiddenChanges).toEqual([]);
  });

  it('snapshots criteria independently for the transaction', () => {
    const raw = { xml: '<autoFilter ref="A1:A5"/>' };
    const original = createAutoFilter(sheet.id, range, [[0, valuesFilter([numberValue(10)])]], raw);
    const snapshot = cloneAutoFilter(original);
    original.columns.set(1, conditionsFilter('and', numCondition('>', 0)));

    expect(snapshot.columns.size).toBe(1);
    expect(snapshot.range).toEqual(range);
    expect(snapshot.raw).toBe(raw);
    expect(snapshot.columns.get(0)).toEqual(valuesFilter([numberValue(10)]));
  });
});

describe('F10 filter-sort', () => {
  const sheet = sheetOf([
    [textValue('Value')],
    [numberValue(10)],
    [numberValue(30)],
    [numberValue(20)],
    [numberValue(40)],
    [numberValue(50)],
  ]);
  const range = rangeOf(5);
  const filter = filterFor(sheet, range, [
    [0, valuesFilter([numberValue(10), numberValue(20), numberValue(30)])],
  ]);
  const manual = new Set([2]);
  const target = rangeOf(5, 0, 1);

  it('reorders only visible rows and leaves hidden rows in place', () => {
    const plan = planFilteredSort(sheet, filter, manual, target);
    expect(plan.filter).toBe(filter);
    expect(plan.visibleRows).toEqual([1, 3]);
    expect(plan.fixedRows).toEqual([2, 4, 5]);

    const predicate = visibleRowPredicate(sheet, filter, manual);
    expect([1, 2, 3, 4, 5].filter(predicate)).toEqual([1, 3]);
    expect(visibleRowsInRange(sheet, filter, manual, target)).toEqual([1, 3]);
  });

  it('keeps the same filter and the same visible set after the sort', () => {
    // Simulate the sort lane permuting values among the visible slots.
    setCellSnapshot(sheet, 1, 0, cell(numberValue(20)));
    setCellSnapshot(sheet, 3, 0, cell(numberValue(10)));

    const after = planFilteredSort(sheet, filter, manual, target);
    expect(after.filter).toBe(filter);
    expect(after.visibleRows).toEqual([1, 3]);
    expect(after.fixedRows).toEqual([2, 4, 5]);
    expect(visibleDataRows(sheet, filter)).toEqual([1, 2, 3]);

    // Filter-hidden rows never move, so their values are untouched.
    expect(sheet.cells.get(cellKey(4, 0))?.value).toEqual(numberValue(40));
    expect(sheet.cells.get(cellKey(5, 0))?.value).toEqual(numberValue(50));
  });
});

describe('filter property invariants (testing-conformance §10)', () => {
  it('every data row is visible or filter-hidden, the header stays visible, and clearing restores the manual layer', () => {
    const arbitrary = fc.record({
      values: fc.array(fc.option(fc.integer({ min: 0, max: 5 }), { nil: null }), {
        minLength: 1,
        maxLength: 12,
      }),
      selectedFlags: fc.array(fc.boolean(), { minLength: 6, maxLength: 6 }),
      includeBlanks: fc.boolean(),
      manualFlags: fc.array(fc.boolean(), { minLength: 12, maxLength: 12 }),
    });

    fc.assert(
      fc.property(arbitrary, ({ values, selectedFlags, includeBlanks, manualFlags }) => {
        const sheet = sheetOf([
          [textValue('Value')],
          ...values.map((value) => [value === null ? BLANK : numberValue(value)]),
        ]);
        const range = rangeOf(values.length);
        const selected = selectedFlags.flatMap((chosen, index) =>
          chosen ? [numberValue(index)] : [],
        );
        const filter = filterFor(sheet, range, [[0, valuesFilter(selected, includeBlanks)]]);
        const manual = new Set<number>();
        manualFlags.slice(0, values.length).forEach((flag, index) => {
          if (flag) manual.add(index + 1);
        });

        const layers = computeHiddenLayers(sheet, filter, manual);
        const evaluation = evaluateAutoFilter(sheet, filter);
        const predicate = visibleRowPredicate(sheet, filter, manual);
        for (let row = 1; row <= values.length; row += 1) {
          const expectedHidden = layers.filterHidden.has(row) || manual.has(row);
          expect(layers.effectiveHidden.has(row)).toBe(expectedHidden);
          expect(predicate(row)).toBe(!expectedHidden);
        }
        expect(evaluation.filterHidden.has(0)).toBe(false);
        expect(isHeaderRow(filter, 0)).toBe(true);

        const cleared = clearFilterLayers(layers);
        expect(cleared.filterHidden.size).toBe(0);
        expect([...cleared.effectiveHidden].sort((a, b) => a - b)).toEqual(
          [...manual].sort((a, b) => a - b),
        );

        const restored = restoreHiddenLayers(sheet, filter, layers.effectiveHidden);
        expect([...restored.effectiveHidden].sort((a, b) => a - b)).toEqual(
          [...layers.effectiveHidden].sort((a, b) => a - b),
        );
        expect([...evaluateAutoFilter(sheet, filter).filterHidden].sort((a, b) => a - b)).toEqual(
          [...evaluation.filterHidden].sort((a, b) => a - b),
        );
      }),
      { seed: 20251118, numRuns: 200 },
    );
  });

  it('the §2.4 comparator is antisymmetric, transitive and keeps blanks last', () => {
    const arbitraryValue = fc.oneof(
      fc.integer({ min: -100, max: 100 }).map((value) => numberValue(value)),
      fc.string().map((value) => textValue(value)),
      fc.boolean().map((value) => booleanValue(value)),
      fc.constantFrom(errorValue('#N/A'), errorValue('#DIV/0!'), errorValue('#VALUE!')),
      fc.constant(BLANK),
    );

    fc.assert(
      fc.property(fc.array(arbitraryValue, { maxLength: 30 }), (values) => {
        for (const value of values) {
          expect(compareFilterValues(value, value)).toBe(0);
          if (value.type !== 'blank') {
            expect(compareFilterValues(value, BLANK)).toBeLessThan(0);
            expect(compareFilterValues(BLANK, value)).toBeGreaterThan(0);
          }
        }
        for (const left of values) {
          for (const right of values) {
            expect(
              Math.sign(compareFilterValues(left, right)) +
                Math.sign(compareFilterValues(right, left)),
            ).toBe(0);
          }
        }
        const sorted = [...values].sort(compareFilterValues);
        for (let index = 1; index < sorted.length; index += 1) {
          expect(compareFilterValues(sorted[index - 1]!, sorted[index]!)).toBeLessThanOrEqual(0);
        }
      }),
      { seed: 20251118, numRuns: 100 },
    );
  });
});
