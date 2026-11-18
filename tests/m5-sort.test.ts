import { describe, expect, it } from 'vitest';
import fc from 'fast-check';
import type { Cell, GridRange, Sheet } from '@model/types';
import { createCell, getCell, setCellSnapshot } from '@model/types';
import {
  ARRAY_BLOCK_MESSAGE,
  HIDDEN_ROWS_WARNING,
  KEY_OUTSIDE_RANGE_MESSAGE,
  MERGED_SIZE_MESSAGE,
  SUB_RANGE_WARNING,
  TABLE_BLOCK_MESSAGE,
  TOO_MANY_KEYS_MESSAGE,
  compareCellValues,
  compareText,
  guessHeaderRow,
  planSort,
  resolveTargetRange,
  type SortDirection,
  type SortPlan,
  type SortResult,
} from '@model/sort';
import type { CellErrorCode, CellValue } from '@model/value';
import {
  BLANK,
  booleanValue,
  errorValue,
  formatCellValue,
  numberValue,
  textValue,
} from '@model/value';
import { appendStyle, createEmptyWorkbook, createSheet } from '@model/workbook';
import { DEFAULT_BORDER, DEFAULT_FONT } from '@model/style';

/**
 * Sort conformance fixtures S1-S14 of sort-autofilter §4, plus the property
 * invariants of testing-conformance §10. S15 (undo/redo) is command wiring and
 * lives with the command tests.
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

function put(sheet: Sheet, row: number, col: number, value: CellValue, styleId = 0): void {
  setCellSnapshot(sheet, row, col, createCell(value, styleId));
}

function putFormula(sheet: Sheet, row: number, col: number, formula: string): void {
  setCellSnapshot(sheet, row, col, {
    value: numberValue(0),
    formula,
    formulaDerived: false,
    savedResult: numberValue(0),
    styleId: 0,
    richText: null,
  });
}

function sortedPlan(result: SortResult): SortPlan {
  if (result.kind !== 'sorted') {
    throw new Error(`expected a sorted plan, got ${result.kind}`);
  }
  return result.plan;
}

/** The value at a destination address after the plan is applied. */
function valueAfter(sheet: Sheet, plan: SortPlan, row: number, col: number): CellValue {
  const record = plan.cellRecords.find(
    (candidate) => candidate.address.row === row && candidate.address.col === col,
  );
  if (record !== undefined) return record.after?.value ?? BLANK;
  return getCell(sheet, row, col)?.value ?? BLANK;
}

function columnAfter(sheet: Sheet, plan: SortPlan, col: number): CellValue[] {
  const values: CellValue[] = [];
  for (let row = plan.range.startRow; row <= plan.range.endRow; row += 1) {
    values.push(valueAfter(sheet, plan, row, col));
  }
  return values;
}

function labelsAfter(sheet: Sheet, plan: SortPlan, col: number): string[] {
  return columnAfter(sheet, plan, col).map(formatCellValue);
}

function sameRange(left: GridRange, right: GridRange): boolean {
  return (
    left.startRow === right.startRow &&
    left.startCol === right.startCol &&
    left.endRow === right.endRow &&
    left.endCol === right.endCol
  );
}

/** Apply the cell and merge records, as the sort command will. */
function applyPlan(sheet: Sheet, plan: SortPlan): void {
  for (const record of plan.cellRecords) {
    setCellSnapshot(sheet, record.address.row, record.address.col, record.after);
  }
  // The plan's `before` ranges describe the pre-sort sheet, so locate them
  // against a snapshot of the original merge list.
  const original = sheet.merges.map((merge) => merge.range);
  for (const move of plan.mergeRewrites) {
    const index = original.findIndex((merge) => sameRange(merge, move.before));
    if (index >= 0) sheet.merges[index] = { range: move.after };
  }
}

function stylePair() {
  const workbook = createEmptyWorkbook();
  const sheet = createSheet('Sheet1', 'xl/worksheets/sheet1.xml');
  workbook.sheets.push(sheet);
  const bold = appendStyle(workbook, {
    numFmtId: 0,
    fontId: 1,
    fillId: 0,
    borderId: 0,
    xfId: 0,
    rawXml: '',
    font: { ...DEFAULT_FONT, bold: true },
  });
  const filled = appendStyle(workbook, {
    numFmtId: 0,
    fontId: 0,
    fillId: 2,
    borderId: 0,
    xfId: 0,
    rawXml: '',
    fill: { pattern: 'solid', color: 'FFFF00' },
  });
  const bordered = appendStyle(workbook, {
    numFmtId: 0,
    fontId: 0,
    fillId: 0,
    borderId: 1,
    xfId: 0,
    rawXml: '',
    border: { ...DEFAULT_BORDER, bottom: { style: 'thin', color: null, weight: 'thin' } },
  });
  return { workbook, sheet, bold, filled, bordered };
}

describe('S1 sort-values', () => {
  it('orders one key ascending and descending', () => {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, numberValue(30));
    put(sheet, 1, 0, numberValue(10));
    put(sheet, 2, 0, numberValue(20));
    const selection = range(0, 0, 2, 0);

    const ascending = sortedPlan(
      planSort({
        workbook,
        sheetId: sheet.id,
        selection,
        keys: [{ column: 0, direction: 'ascending' }],
        hasHeaders: false,
      }),
    );
    expect(columnAfter(sheet, ascending, 0)).toEqual([
      numberValue(10),
      numberValue(20),
      numberValue(30),
    ]);
    expect(ascending.rowPermutation).toEqual([1, 2, 0]);

    const descending = sortedPlan(
      planSort({
        workbook,
        sheetId: sheet.id,
        selection,
        keys: [{ column: 0, direction: 'descending' }],
        hasHeaders: false,
      }),
    );
    expect(columnAfter(sheet, descending, 0)).toEqual([
      numberValue(30),
      numberValue(20),
      numberValue(10),
    ]);
    expect(descending.rowPermutation).toEqual([0, 2, 1]);
  });
});

describe('S2 sort-multi', () => {
  it('applies three keys in order with a tie-break', () => {
    const { workbook, sheet } = workbookWithSheet();
    const rows: readonly [string, number, number, string][] = [
      ['b', 2, 1, 'r0'],
      ['a', 2, 5, 'r1'],
      ['a', 1, 9, 'r2'],
      ['a', 2, 1, 'r3'],
    ];
    rows.forEach(([letter, second, third, label], row) => {
      put(sheet, row, 0, textValue(letter));
      put(sheet, row, 1, numberValue(second));
      put(sheet, row, 2, numberValue(third));
      put(sheet, row, 3, textValue(label));
    });

    const plan = sortedPlan(
      planSort({
        workbook,
        sheetId: sheet.id,
        selection: range(0, 0, 3, 3),
        keys: [
          { column: 0, direction: 'ascending' },
          { column: 1, direction: 'ascending' },
          { column: 2, direction: 'descending' },
        ],
        hasHeaders: false,
      }),
    );

    expect(labelsAfter(sheet, plan, 3)).toEqual(['r2', 'r1', 'r3', 'r0']);
    expect(plan.rowPermutation).toEqual([2, 1, 3, 0]);
    expect(plan.keys).toHaveLength(3);
  });
});

describe('S3 sort-header', () => {
  function headerTable() {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, textValue('Name'));
    put(sheet, 0, 1, textValue('Qty'));
    put(sheet, 1, 0, textValue('pear'));
    put(sheet, 1, 1, numberValue(2));
    put(sheet, 2, 0, textValue('apple'));
    put(sheet, 2, 1, numberValue(1));
    return { workbook, sheet };
  }

  it('guesses the header, keeps it in place and labels keys from the header row', () => {
    const { workbook, sheet } = headerTable();
    const plan = sortedPlan(
      planSort({
        workbook,
        sheetId: sheet.id,
        selection: range(0, 0, 2, 1),
        keys: [{ column: 1, direction: 'ascending' }],
        hasHeaders: null,
      }),
    );

    expect(plan.hasHeaders).toBe(true);
    expect(plan.rowPermutation[0]).toBe(0);
    expect(plan.keyLabels).toEqual(['Qty']);
    expect(labelsAfter(sheet, plan, 0)).toEqual(['Name', 'apple', 'pear']);
  });

  it('sorts the first row as data and labels keys by column letter with headers off', () => {
    const { workbook, sheet } = headerTable();
    const plan = sortedPlan(
      planSort({
        workbook,
        sheetId: sheet.id,
        selection: range(0, 0, 2, 1),
        keys: [{ column: 1, direction: 'ascending' }],
        hasHeaders: false,
      }),
    );

    expect(plan.hasHeaders).toBe(false);
    expect(plan.keyLabels).toEqual(['B']);
    // The text "Qty" sorts after every number, so the old header row lands last.
    expect(plan.rowPermutation).toEqual([2, 1, 0]);
    expect(labelsAfter(sheet, plan, 0)).toEqual(['apple', 'pear', 'Name']);
  });

  it('honours the explicit override when the guess is off', () => {
    const { workbook, sheet } = headerTable();
    const plan = sortedPlan(
      planSort({
        workbook,
        sheetId: sheet.id,
        selection: range(0, 0, 2, 1),
        keys: [{ column: 1, direction: 'ascending' }],
        hasHeaders: true,
      }),
    );
    expect(plan.hasHeaders).toBe(true);
    expect(plan.rowPermutation[0]).toBe(0);
  });

  it('guesses from a distinct style when every data cell is text', () => {
    const plain = workbookWithSheet();
    put(plain.sheet, 0, 0, textValue('Name'));
    put(plain.sheet, 1, 0, textValue('alpha'));
    put(plain.sheet, 2, 0, textValue('beta'));
    const plainRange = range(0, 0, 2, 0);
    expect(
      guessHeaderRow({ workbook: plain.workbook, sheet: plain.sheet, range: plainRange }),
    ).toBe(false);

    const styled = stylePair();
    put(styled.sheet, 0, 0, textValue('Name'), styled.bold);
    put(styled.sheet, 1, 0, textValue('alpha'));
    put(styled.sheet, 2, 0, textValue('beta'));
    expect(
      guessHeaderRow({ workbook: styled.workbook, sheet: styled.sheet, range: plainRange }),
    ).toBe(true);

    const filled = stylePair();
    put(filled.sheet, 0, 0, textValue('Name'), filled.filled);
    put(filled.sheet, 1, 0, textValue('alpha'));
    put(filled.sheet, 2, 0, textValue('beta'));
    expect(
      guessHeaderRow({ workbook: filled.workbook, sheet: filled.sheet, range: plainRange }),
    ).toBe(true);
  });

  it('refuses a header when the first row is not all text', () => {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, textValue('Name'));
    put(sheet, 1, 0, textValue('alpha'));
    put(sheet, 2, 0, textValue('beta'));
    const withGap = range(0, 0, 2, 1);
    expect(guessHeaderRow({ workbook, sheet, range: withGap })).toBe(false);
  });
});

describe('S4 sort-mixed-types', () => {
  function mixedTable() {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, numberValue(10));
    put(sheet, 1, 0, textValue('abc'));
    put(sheet, 2, 0, booleanValue(true));
    put(sheet, 3, 0, errorValue('#N/A'));
    // Row 4 has no value in the key column: a blank.
    put(sheet, 5, 0, numberValue(2));
    put(sheet, 6, 0, textValue('5'));
    put(sheet, 7, 0, errorValue('#DIV/0!'));
    put(sheet, 8, 0, textValue(' '));
    put(sheet, 9, 0, booleanValue(false));
    for (let row = 0; row <= 9; row += 1) put(sheet, row, 1, textValue(`r${row}`));
    return { workbook, sheet };
  }

  it('orders numbers, text, logicals, errors and blanks with blanks last', () => {
    const { workbook, sheet } = mixedTable();
    const plan = sortedPlan(
      planSort({
        workbook,
        sheetId: sheet.id,
        selection: range(0, 0, 9, 1),
        keys: [{ column: 0, direction: 'ascending' }],
        hasHeaders: false,
      }),
    );

    expect(labelsAfter(sheet, plan, 1)).toEqual([
      'r5', // 2
      'r0', // 10
      'r6', // "5", a number stored as text, sorts as text after every number
      'r8', // " ", a spaces-only cell is text and sorts before punctuation and letters
      'r1', // "abc"
      'r9', // FALSE
      'r2', // TRUE
      'r3', // #N/A
      'r7', // #DIV/0!, equal to #N/A so the stable tie-break keeps the order
      'r4', // blank
    ]);
  });

  it('keeps blanks last when descending and reverses only ranks 1-4', () => {
    const { workbook, sheet } = mixedTable();
    const plan = sortedPlan(
      planSort({
        workbook,
        sheetId: sheet.id,
        selection: range(0, 0, 9, 1),
        keys: [{ column: 0, direction: 'descending' }],
        hasHeaders: false,
      }),
    );

    expect(labelsAfter(sheet, plan, 1)).toEqual([
      'r3', // #N/A
      'r7', // #DIV/0!
      'r2', // TRUE
      'r9', // FALSE
      'r1', // "abc"
      'r8', // " "
      'r6', // "5"
      'r0', // 10
      'r5', // 2
      'r4', // blank last in both directions
    ]);
  });
});

describe('S5 sort-formulas', () => {
  it('shifts relative references and keeps absolute, mixed and sheet prefixes', () => {
    const { workbook, sheet } = workbookWithSheet();
    workbook.sheets.push(createSheet('Data Sheet', 'xl/worksheets/sheet2.xml'));
    put(sheet, 0, 0, numberValue(3));
    putFormula(sheet, 0, 1, '$B$1+B1');
    putFormula(sheet, 0, 2, '$B$1');
    put(sheet, 1, 0, numberValue(1));
    putFormula(sheet, 1, 1, "'Data Sheet'!B2+A3");
    put(sheet, 2, 0, numberValue(2));
    putFormula(sheet, 2, 1, 'SUM(B$1:B3)');

    const plan = sortedPlan(
      planSort({
        workbook,
        sheetId: sheet.id,
        selection: range(0, 0, 2, 2),
        keys: [{ column: 0, direction: 'ascending' }],
        hasHeaders: false,
      }),
    );

    expect(plan.rowPermutation).toEqual([1, 2, 0]);
    expect(
      plan.formulaRewrites.map((rewrite) => [
        rewrite.address.row,
        rewrite.address.col,
        rewrite.after,
      ]),
    ).toEqual([
      [0, 1, "'Data Sheet'!B1+A2"],
      [1, 1, 'SUM(B$1:B2)'],
      [2, 1, '$B$1+B3'],
    ]);
    // The absolute-only formula moved but its text did not change, so it is not
    // a rewrite and the cell keeps its saved result unstale.
    expect(plan.formulaRewrites.some((rewrite) => rewrite.address.col === 2)).toBe(false);

    applyPlan(sheet, plan);
    expect(getCell(sheet, 0, 1)?.formula).toBe("'Data Sheet'!B1+A2");
    expect(getCell(sheet, 2, 2)?.formula).toBe('$B$1');
    expect(getCell(sheet, 1, 1)?.formula).toBe('SUM(B$1:B2)');
    expect(getCell(sheet, 2, 1)?.formula).toBe('$B$1+B3');
    expect(getCell(sheet, 2, 1)?.stale).toBe(true);
    expect(getCell(sheet, 2, 2)?.stale ?? false).toBe(false);
    // Every moved cell keeps its saved result beside the shifted formula.
    expect(getCell(sheet, 2, 1)?.savedResult).toEqual(numberValue(0));
  });

  it('turns a relative reference that leaves the grid into #REF!', () => {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, numberValue(2));
    put(sheet, 1, 0, numberValue(1));
    putFormula(sheet, 1, 1, 'A1');

    const plan = sortedPlan(
      planSort({
        workbook,
        sheetId: sheet.id,
        selection: range(0, 0, 1, 1),
        keys: [{ column: 0, direction: 'ascending' }],
        hasHeaders: false,
      }),
    );

    applyPlan(sheet, plan);
    expect(getCell(sheet, 0, 1)?.formula).toBe('#REF!');
  });
});

describe('S6 sort-outside-refs', () => {
  it('does not rewrite a formula outside the sorted range', () => {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, numberValue(2));
    put(sheet, 1, 0, numberValue(1));
    putFormula(sheet, 4, 3, 'A1+A2');
    putFormula(sheet, 5, 3, 'SUM(A1:A2)');

    const plan = sortedPlan(
      planSort({
        workbook,
        sheetId: sheet.id,
        selection: range(0, 0, 1, 0),
        keys: [{ column: 0, direction: 'ascending' }],
        hasHeaders: false,
      }),
    );

    expect(plan.formulaRewrites).toHaveLength(0);
    expect(getCell(sheet, 4, 3)?.formula).toBe('A1+A2');
    expect(getCell(sheet, 5, 3)?.formula).toBe('SUM(A1:A2)');
  });

  it('does not rewrite a cross-sheet formula pointing into the range', () => {
    const { workbook, sheet } = workbookWithSheet();
    const other = createSheet('Sheet2', 'xl/worksheets/sheet2.xml');
    workbook.sheets.push(other);
    put(sheet, 0, 0, numberValue(2));
    put(sheet, 1, 0, numberValue(1));
    putFormula(other, 0, 0, 'Sheet1!A1');

    const plan = sortedPlan(
      planSort({
        workbook,
        sheetId: sheet.id,
        selection: range(0, 0, 1, 0),
        keys: [{ column: 0, direction: 'ascending' }],
        hasHeaders: false,
      }),
    );

    expect(plan.formulaRewrites).toHaveLength(0);
    expect(getCell(other, 0, 0)?.formula).toBe('Sheet1!A1');
  });
});

describe('S7 sort-hidden', () => {
  it('keeps manual- and filter-hidden rows in place and warns once', () => {
    const { workbook, sheet } = workbookWithSheet();
    const values = [50, 40, 30, 20, 10];
    values.forEach((value, row) => put(sheet, row, 0, numberValue(value)));
    sheet.rowDimensions.set(1, { hidden: true });

    const result = planSort({
      workbook,
      sheetId: sheet.id,
      selection: range(0, 0, 4, 0),
      keys: [{ column: 0, direction: 'ascending' }],
      hasHeaders: false,
      filterHiddenRows: new Set([3]),
    });
    const plan = sortedPlan(result);

    // Visible rows 0, 2 and 4 reorder into their own slots.
    expect(plan.rowPermutation).toEqual([4, 1, 2, 3, 0]);
    expect(plan.crossesHiddenRows).toBe(true);
    expect(plan.hiddenRowCount).toBe(2);
    expect(plan.visibleRowCount).toBe(3);
    expect(result.kind === 'sorted' ? result.warnings : []).toEqual([
      { code: 'hidden-rows', message: HIDDEN_ROWS_WARNING, count: 2 },
    ]);

    applyPlan(sheet, plan);
    expect([0, 1, 2, 3, 4].map((row) => getCell(sheet, row, 0)?.value)).toEqual([
      numberValue(10),
      numberValue(40),
      numberValue(30),
      numberValue(20),
      numberValue(50),
    ]);
  });

  it('is silent when no row is hidden', () => {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, numberValue(2));
    put(sheet, 1, 0, numberValue(1));
    const result = planSort({
      workbook,
      sheetId: sheet.id,
      selection: range(0, 0, 1, 0),
      keys: [{ column: 0, direction: 'ascending' }],
      hasHeaders: false,
    });
    expect(result.kind === 'sorted' ? result.warnings : []).toEqual([]);
  });
});

describe('S8 sort-merges-uniform', () => {
  it('moves one-row merges fully inside the range with their rows', () => {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, numberValue(3));
    put(sheet, 1, 0, numberValue(2));
    put(sheet, 2, 0, numberValue(1));
    sheet.merges.push({ range: range(0, 0, 0, 1) }, { range: range(2, 0, 2, 1) });

    const plan = sortedPlan(
      planSort({
        workbook,
        sheetId: sheet.id,
        selection: range(0, 0, 2, 1),
        keys: [{ column: 0, direction: 'ascending' }],
        hasHeaders: false,
      }),
    );

    expect(plan.mergeRewrites).toEqual([
      { before: range(0, 0, 0, 1), after: range(2, 0, 2, 1) },
      { before: range(2, 0, 2, 1), after: range(0, 0, 0, 1) },
    ]);

    applyPlan(sheet, plan);
    expect(sheet.merges.map((merge) => merge.range)).toEqual([
      range(2, 0, 2, 1),
      range(0, 0, 0, 1),
    ]);
    expect(getCell(sheet, 0, 0)?.value).toEqual(numberValue(1));
    expect(getCell(sheet, 2, 0)?.value).toEqual(numberValue(3));
  });
});

describe('S9 sort-merges-mixed', () => {
  it('blocks a merge that spans rows with the Excel message', () => {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, numberValue(2));
    put(sheet, 1, 0, numberValue(1));
    sheet.merges.push({ range: range(0, 0, 1, 1) });

    const result = planSort({
      workbook,
      sheetId: sheet.id,
      selection: range(0, 0, 1, 1),
      keys: [{ column: 0, direction: 'ascending' }],
      hasHeaders: false,
    });

    expect(result.kind).toBe('blocked');
    expect(result.kind === 'blocked' ? result.code : null).toBe('merged-cells');
    expect(result.kind === 'blocked' ? result.message : null).toBe(MERGED_SIZE_MESSAGE);
    expect(MERGED_SIZE_MESSAGE).toBe(
      'This operation requires the merged cells to be identically sized.',
    );
  });

  it('blocks a merge that is only partly inside the range', () => {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, numberValue(2));
    put(sheet, 1, 0, numberValue(1));
    put(sheet, 2, 0, numberValue(3));
    sheet.merges.push({ range: range(2, 0, 4, 0) });

    const result = planSort({
      workbook,
      sheetId: sheet.id,
      selection: range(0, 0, 2, 0),
      keys: [{ column: 0, direction: 'ascending' }],
      hasHeaders: false,
    });

    expect(result.kind).toBe('blocked');
    expect(result.kind === 'blocked' ? result.message : null).toBe(MERGED_SIZE_MESSAGE);
  });
});

describe('S10 sort-row-props', () => {
  it('moves styles with cells and leaves row heights and hidden flags in place', () => {
    const { workbook, sheet, bold, filled } = stylePair();
    put(sheet, 0, 0, numberValue(3));
    put(sheet, 0, 1, textValue('r0'), bold);
    put(sheet, 1, 0, numberValue(2));
    put(sheet, 1, 1, textValue('r1'), filled);
    put(sheet, 2, 0, numberValue(1));
    put(sheet, 2, 1, textValue('r2'));
    sheet.rowDimensions.set(0, { height: 30 });
    sheet.rowDimensions.set(1, { hidden: false, height: 12 });
    sheet.rowDimensions.set(2, { height: 45 });
    sheet.columnDimensions.set(1, { width: 24 });

    const plan = sortedPlan(
      planSort({
        workbook,
        sheetId: sheet.id,
        selection: range(0, 0, 2, 1),
        keys: [{ column: 0, direction: 'ascending' }],
        hasHeaders: false,
      }),
    );

    applyPlan(sheet, plan);

    // Styles travel with the row that owned them.
    expect(getCell(sheet, 0, 1)?.styleId).toBe(0);
    expect(getCell(sheet, 1, 1)?.styleId).toBe(filled);
    expect(getCell(sheet, 2, 1)?.styleId).toBe(bold);

    // Row dimensions are address-based, so they stay where they were.
    expect(sheet.rowDimensions.get(0)?.height).toBe(30);
    expect(sheet.rowDimensions.get(1)?.height).toBe(12);
    expect(sheet.rowDimensions.get(2)?.height).toBe(45);
    expect(sheet.columnDimensions.get(1)?.width).toBe(24);
    expect(plan.cellRecords.some((record) => record.address.col === 1)).toBe(true);
  });

  it('moves comments and hyperlinks with their row', () => {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, numberValue(2));
    put(sheet, 1, 0, numberValue(1));
    sheet.comments.set('0:1', { text: 'first', threaded: false });
    sheet.hyperlinks.set('0:2', { target: 'https://example.invalid/report' });

    const plan = sortedPlan(
      planSort({
        workbook,
        sheetId: sheet.id,
        selection: range(0, 0, 1, 2),
        keys: [{ column: 0, direction: 'ascending' }],
        hasHeaders: false,
      }),
    );

    expect(plan.commentMoves).toEqual([
      {
        from: { row: 0, col: 1 },
        to: { row: 1, col: 1 },
        record: { text: 'first', threaded: false },
      },
    ]);
    expect(plan.hyperlinkMoves).toEqual([
      {
        from: { row: 0, col: 2 },
        to: { row: 1, col: 2 },
        record: { target: 'https://example.invalid/report' },
      },
    ]);
  });
});

describe('S11 sort-array', () => {
  function arrayCell(formula: string, covered: string): Cell {
    return {
      value: numberValue(6),
      formula,
      formulaDerived: false,
      savedResult: numberValue(6),
      styleId: 0,
      richText: null,
      formulaSource: { kind: 'array', range: covered, xml: `<f t="array" ref="${covered}">` },
    };
  }

  it('blocks a range that intersects a legacy array formula', () => {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, numberValue(2));
    put(sheet, 1, 0, numberValue(1));
    setCellSnapshot(sheet, 1, 1, arrayCell('SUM(B1:B3)', 'B1:B3'));

    const result = planSort({
      workbook,
      sheetId: sheet.id,
      selection: range(0, 0, 1, 1),
      keys: [{ column: 0, direction: 'ascending' }],
      hasHeaders: false,
    });

    expect(result.kind).toBe('blocked');
    expect(result.kind === 'blocked' ? result.code : null).toBe('array-formula');
    expect(result.kind === 'blocked' ? result.message : null).toBe(ARRAY_BLOCK_MESSAGE);
  });

  it('allows a sort beside an array formula', () => {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, numberValue(2));
    put(sheet, 1, 0, numberValue(1));
    setCellSnapshot(sheet, 1, 3, arrayCell('SUM(D1:D3)', 'D1:D3'));

    const result = planSort({
      workbook,
      sheetId: sheet.id,
      selection: range(0, 0, 1, 0),
      keys: [{ column: 0, direction: 'ascending' }],
      hasHeaders: false,
    });

    expect(result.kind).toBe('sorted');
  });
});

describe('S12 sort-table', () => {
  it('blocks a range that intersects a table', () => {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, numberValue(2));
    put(sheet, 1, 0, numberValue(1));

    const result = planSort({
      workbook,
      sheetId: sheet.id,
      selection: range(0, 0, 1, 1),
      keys: [{ column: 0, direction: 'ascending' }],
      hasHeaders: false,
      tableRanges: [range(0, 0, 4, 3)],
    });

    expect(result.kind).toBe('blocked');
    expect(result.kind === 'blocked' ? result.code : null).toBe('table');
    expect(result.kind === 'blocked' ? result.message : null).toBe(TABLE_BLOCK_MESSAGE);
  });

  it('allows a sort beside a table', () => {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, numberValue(2));
    put(sheet, 1, 0, numberValue(1));

    const result = planSort({
      workbook,
      sheetId: sheet.id,
      selection: range(0, 0, 1, 0),
      keys: [{ column: 0, direction: 'ascending' }],
      hasHeaders: false,
      tableRanges: [range(10, 0, 14, 3)],
    });

    expect(result.kind).toBe('sorted');
  });
});

describe('S13 sort-64-keys', () => {
  it('accepts 64 keys and orders by them', () => {
    const { workbook, sheet } = workbookWithSheet();
    const values = [9, 1, 5];
    values.forEach((value, row) => {
      for (let col = 0; col < 64; col += 1) put(sheet, row, col, numberValue(value));
    });

    const result = planSort({
      workbook,
      sheetId: sheet.id,
      selection: range(0, 0, 2, 63),
      keys: Array.from({ length: 64 }, (_unused, col) => ({
        column: col,
        direction: 'ascending' as const,
      })),
      hasHeaders: false,
    });
    const plan = sortedPlan(result);

    expect(plan.keys).toHaveLength(64);
    expect(plan.rowPermutation).toEqual([1, 2, 0]);
  });

  it('blocks a 65th key', () => {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, numberValue(2));
    put(sheet, 1, 0, numberValue(1));

    const result = planSort({
      workbook,
      sheetId: sheet.id,
      selection: range(0, 0, 1, 1),
      keys: Array.from({ length: 65 }, (_unused, index) => ({
        column: index % 2,
        direction: 'ascending' as const,
      })),
      hasHeaders: false,
    });

    expect(result.kind).toBe('blocked');
    expect(result.kind === 'blocked' ? result.code : null).toBe('too-many-keys');
    expect(result.kind === 'blocked' ? result.message : null).toBe(TOO_MANY_KEYS_MESSAGE);
  });

  it('blocks a key column outside the range', () => {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, numberValue(2));
    put(sheet, 1, 0, numberValue(1));

    const result = planSort({
      workbook,
      sheetId: sheet.id,
      selection: range(0, 0, 1, 0),
      keys: [{ column: 4, direction: 'ascending' }],
      hasHeaders: false,
    });

    expect(result.kind).toBe('blocked');
    expect(result.kind === 'blocked' ? result.code : null).toBe('key-outside-range');
    expect(result.kind === 'blocked' ? result.message : null).toBe(KEY_OUTSIDE_RANGE_MESSAGE);
  });
});

describe('S14 sort-stable', () => {
  function tiedTable() {
    const { workbook, sheet } = workbookWithSheet();
    const keys = [2, 1, 2, 1];
    keys.forEach((value, row) => {
      put(sheet, row, 0, numberValue(value));
      put(sheet, row, 1, textValue(`r${row}`));
    });
    return { workbook, sheet };
  }

  it('breaks a tie on the next key before falling back to row order', () => {
    const { workbook, sheet } = workbookWithSheet();
    const secondKeys = [2, 1, 2, 1];
    secondKeys.forEach((value, row) => {
      put(sheet, row, 0, numberValue(1));
      put(sheet, row, 1, numberValue(value));
      put(sheet, row, 2, textValue(`r${row}`));
    });

    const plan = sortedPlan(
      planSort({
        workbook,
        sheetId: sheet.id,
        selection: range(0, 0, 3, 2),
        keys: [
          { column: 0, direction: 'ascending' },
          { column: 1, direction: 'ascending' },
        ],
        hasHeaders: false,
      }),
    );
    // The first key ties everywhere, so the second key decides; its own ties
    // keep the original row order.
    expect(labelsAfter(sheet, plan, 2)).toEqual(['r1', 'r3', 'r0', 'r2']);
  });

  it('keeps equal keys in their original order in both directions', () => {
    const { workbook, sheet } = tiedTable();
    const ascending = sortedPlan(
      planSort({
        workbook,
        sheetId: sheet.id,
        selection: range(0, 0, 3, 1),
        keys: [{ column: 0, direction: 'ascending' }],
        hasHeaders: false,
      }),
    );
    expect(labelsAfter(sheet, ascending, 1)).toEqual(['r1', 'r3', 'r0', 'r2']);

    const descending = sortedPlan(
      planSort({
        workbook,
        sheetId: sheet.id,
        selection: range(0, 0, 3, 1),
        keys: [{ column: 0, direction: 'descending' }],
        hasHeaders: false,
      }),
    );
    expect(labelsAfter(sheet, descending, 1)).toEqual(['r0', 'r2', 'r1', 'r3']);
  });
});

describe('§2.1 target range', () => {
  it('expands a single cell to the current contiguous region', () => {
    const { sheet } = workbookWithSheet();
    for (let row = 0; row <= 2; row += 1) {
      for (let col = 0; col <= 1; col += 1) put(sheet, row, col, numberValue(row + col));
    }
    // A cell outside the block, separated by an empty row and column.
    put(sheet, 4, 3, numberValue(99));

    expect(resolveTargetRange(sheet, range(1, 1, 1, 1))).toEqual(range(0, 0, 2, 1));
    expect(resolveTargetRange(sheet, range(0, 0, 2, 1))).toEqual(range(0, 0, 2, 1));
  });

  it('leaves an empty single cell as a no-op', () => {
    const { workbook, sheet } = workbookWithSheet();
    const result = planSort({
      workbook,
      sheetId: sheet.id,
      selection: range(6, 6, 6, 6),
      keys: [{ column: 6, direction: 'ascending' }],
      hasHeaders: false,
    });
    expect(result.kind).toBe('unchanged');
    expect(result.kind === 'unchanged' ? result.reason : null).toBe('too-few-rows');
    expect(result.kind === 'unchanged' ? result.range : null).toEqual(range(6, 6, 6, 6));
  });

  it('is a no-op for fewer than two data rows', () => {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, textValue('Name'));
    put(sheet, 0, 1, textValue('Qty'));
    put(sheet, 1, 0, textValue('apple'));
    put(sheet, 1, 1, numberValue(1));

    const result = planSort({
      workbook,
      sheetId: sheet.id,
      selection: range(0, 0, 1, 1),
      keys: [{ column: 1, direction: 'ascending' }],
      hasHeaders: null,
    });
    expect(result.kind).toBe('unchanged');
    expect(result.kind === 'unchanged' ? result.reason : null).toBe('too-few-rows');
  });

  it('warns once when the selection is a sub-range of a larger block', () => {
    const { workbook, sheet } = workbookWithSheet();
    for (let row = 0; row <= 3; row += 1) put(sheet, row, 1, numberValue(row));

    const result = planSort({
      workbook,
      sheetId: sheet.id,
      selection: range(0, 0, 2, 1),
      keys: [{ column: 1, direction: 'ascending' }],
      hasHeaders: false,
    });
    expect(result.kind).toBe('sorted');
    expect(result.kind === 'sorted' ? result.warnings : []).toEqual([
      { code: 'sub-range', message: SUB_RANGE_WARNING },
    ]);
  });

  it('is a no-op when every data row is hidden', () => {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, numberValue(2));
    put(sheet, 1, 0, numberValue(1));
    sheet.rowDimensions.set(0, { hidden: true });
    sheet.rowDimensions.set(1, { hidden: true });

    const result = planSort({
      workbook,
      sheetId: sheet.id,
      selection: range(0, 0, 1, 0),
      keys: [{ column: 0, direction: 'ascending' }],
      hasHeaders: false,
    });
    expect(result.kind).toBe('unchanged');
    expect(result.kind === 'unchanged' ? result.reason : null).toBe('nothing-to-order');
  });
});

describe('§2.4 comparator', () => {
  it('orders digits, space, punctuation and letters in Excel order', () => {
    expect(compareText('9', ' ')).toBeLessThan(0);
    expect(compareText(' ', '!')).toBeLessThan(0);
    expect(compareText('!', '#')).toBeLessThan(0);
    expect(compareText('~', '+')).toBeLessThan(0);
    expect(compareText('+', '<')).toBeLessThan(0);
    expect(compareText('<', 'A')).toBeLessThan(0);
    expect(compareText('Z', 'z')).toBe(0);
  });

  it('ignores case and the apostrophe and hyphen markers', () => {
    expect(compareText('ABC', 'abc')).toBe(0);
    expect(compareText('coop', 'co-op')).toBeLessThan(0);
    expect(compareText('co-op', 'coop')).toBeGreaterThan(0);
    expect(compareText('dont', "don't")).toBeLessThan(0);
    expect(compareText('CO-OP', 'coop')).toBeGreaterThan(0);
  });

  it('falls back to code-point order for non-ASCII text', () => {
    expect(compareText('À', 'É')).toBeLessThan(0);
    expect(compareText('a', 'À')).toBeLessThan(0);
    expect(compareText('Z', 'É')).toBeLessThan(0);
  });

  it('ranks a spaces-only cell as text after every number', () => {
    expect(compareCellValues(textValue(' '), numberValue(1000), 'ascending')).toBeGreaterThan(0);
    expect(compareCellValues(textValue(' '), BLANK, 'ascending')).toBeLessThan(0);
  });

  it('compares errors equal and keeps blanks last both ways', () => {
    expect(compareCellValues(errorValue('#N/A'), errorValue('#DIV/0!'), 'ascending')).toBe(0);
    expect(compareCellValues(errorValue('#N/A'), errorValue('#DIV/0!'), 'descending')).toBe(0);
    expect(compareCellValues(BLANK, numberValue(0), 'ascending')).toBeGreaterThan(0);
    expect(compareCellValues(BLANK, numberValue(0), 'descending')).toBeGreaterThan(0);
  });

  it('reverses ranks 1-4 but not the blank rank', () => {
    expect(compareCellValues(numberValue(1), textValue('1'), 'ascending')).toBeLessThan(0);
    expect(compareCellValues(numberValue(1), textValue('1'), 'descending')).toBeGreaterThan(0);
    expect(compareCellValues(textValue('1'), booleanValue(true), 'ascending')).toBeLessThan(0);
    expect(compareCellValues(booleanValue(false), errorValue('#N/A'), 'ascending')).toBeLessThan(0);
  });
});

const valueArbitrary: fc.Arbitrary<CellValue> = fc.oneof(
  fc.integer({ min: -1000, max: 1000 }).map((value) => numberValue(value)),
  fc.string({ maxLength: 8 }).map((value) => textValue(value)),
  fc.boolean().map((value) => booleanValue(value)),
  fc
    .constantFrom<CellErrorCode>('#N/A', '#DIV/0!', '#VALUE!', '#REF!')
    .map((code) => errorValue(code)),
  fc.constant(BLANK),
);

const DIRECTIONS: readonly SortDirection[] = ['ascending', 'descending'];

describe('property: comparator is a total order with blanks last (§10)', () => {
  it('is antisymmetric', () => {
    fc.assert(
      fc.property(valueArbitrary, valueArbitrary, (left, right) => {
        for (const direction of DIRECTIONS) {
          const forward = Math.sign(compareCellValues(left, right, direction));
          const backward = Math.sign(compareCellValues(right, left, direction));
          // Summing avoids the +0 / -0 distinction of `toBe`.
          expect(forward + backward).toBe(0);
        }
      }),
      { numRuns: 200, seed: 42 },
    );
  });

  it('is transitive', () => {
    fc.assert(
      fc.property(valueArbitrary, valueArbitrary, valueArbitrary, (first, second, third) => {
        for (const direction of DIRECTIONS) {
          if (
            compareCellValues(first, second, direction) <= 0 &&
            compareCellValues(second, third, direction) <= 0
          ) {
            expect(compareCellValues(first, third, direction)).toBeLessThanOrEqual(0);
          }
        }
      }),
      { numRuns: 200, seed: 42 },
    );
  });

  it('puts a blank after every non-blank value in both directions', () => {
    fc.assert(
      fc.property(
        valueArbitrary.filter((value) => value.type !== 'blank'),
        (value) => {
          for (const direction of DIRECTIONS) {
            expect(compareCellValues(value, BLANK, direction)).toBeLessThan(0);
            expect(compareCellValues(BLANK, value, direction)).toBeGreaterThan(0);
          }
        },
      ),
      { numRuns: 100, seed: 42 },
    );
  });
});

describe('property: sort is a permutation and is stable (§10)', () => {
  const rowsArbitrary = fc.array(
    fc.tuple(fc.integer({ min: -3, max: 3 }), fc.integer({ min: 0, max: 99 })),
    { minLength: 2, maxLength: 8 },
  );

  it('preserves the row multiset and orders the key column', () => {
    fc.assert(
      fc.property(
        rowsArbitrary,
        fc.constantFrom<SortDirection>('ascending', 'descending'),
        (rows, direction) => {
          const { workbook, sheet } = workbookWithSheet();
          rows.forEach(([key, payload], row) => {
            put(sheet, row, 0, numberValue(key));
            put(sheet, row, 1, numberValue(payload));
          });
          const plan = sortedPlan(
            planSort({
              workbook,
              sheetId: sheet.id,
              selection: range(0, 0, rows.length - 1, 1),
              keys: [{ column: 0, direction }],
              hasHeaders: false,
            }),
          );

          const before = rows.map(([key, payload]) => `${key}:${payload}`).sort();
          const after: string[] = [];
          for (let row = 0; row < rows.length; row += 1) {
            const key = formatCellValue(valueAfter(sheet, plan, row, 0));
            const payload = formatCellValue(valueAfter(sheet, plan, row, 1));
            after.push(`${key}:${payload}`);
          }
          expect(after.slice().sort()).toEqual(before);

          for (let row = 1; row < rows.length; row += 1) {
            const previous = valueAfter(sheet, plan, row - 1, 0);
            const current = valueAfter(sheet, plan, row, 0);
            expect(compareCellValues(previous, current, direction)).toBeLessThanOrEqual(0);
          }
        },
      ),
      { numRuns: 60, seed: 42 },
    );
  });

  it('keeps the original order when every key ties', () => {
    fc.assert(
      fc.property(
        fc.integer({ min: 2, max: 20 }),
        fc.constantFrom<SortDirection>('ascending', 'descending'),
        (count, direction) => {
          const { workbook, sheet } = workbookWithSheet();
          for (let row = 0; row < count; row += 1) {
            put(sheet, row, 0, numberValue(7));
            put(sheet, row, 1, numberValue(row));
          }
          const plan = sortedPlan(
            planSort({
              workbook,
              sheetId: sheet.id,
              selection: range(0, 0, count - 1, 1),
              keys: [{ column: 0, direction }],
              hasHeaders: false,
            }),
          );
          const payloads: string[] = [];
          for (let row = 0; row < count; row += 1) {
            payloads.push(formatCellValue(valueAfter(sheet, plan, row, 1)));
          }
          expect(payloads).toEqual(Array.from({ length: count }, (_unused, index) => `${index}`));
        },
      ),
      { numRuns: 20, seed: 42 },
    );
  });
});
