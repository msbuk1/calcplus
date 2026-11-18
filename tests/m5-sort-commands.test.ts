import { describe, expect, it } from 'vitest';
import { DocumentSession } from '@commands/session';
import { computeHiddenLayers, createAutoFilter } from '@model/autofilter';
import {
  ARRAY_BLOCK_MESSAGE,
  HIDDEN_ROWS_WARNING,
  KEY_OUTSIDE_RANGE_MESSAGE,
  MAX_SORT_KEYS,
  MERGED_SIZE_MESSAGE,
  SUB_RANGE_WARNING,
  TABLE_BLOCK_MESSAGE,
  TOO_MANY_KEYS_MESSAGE,
  type SortKey,
} from '@model/sort';
import type { GridRange, Sheet } from '@model/types';
import { cellKey, cloneCell, createCell, getCell, setCellSnapshot } from '@model/types';
import type { CellValue } from '@model/value';
import { numberValue, textValue } from '@model/value';
import { appendStyle, createEmptyWorkbook, createSheet } from '@model/workbook';
import { DEFAULT_FONT } from '@model/style';

/**
 * Fixture S15 of sort-autofilter §4 through the command registry: one sort is
 * one transaction, undo restores the exact prior state and redo reapplies the
 * same permutation. The blocking messages, the no-op outcomes and the two
 * warnings are proven here too because they are command outcomes, not model
 * results. S12 (tables) needs the table model and is out of this lane.
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
    stale: false,
  });
}

function key(column: number, direction: 'ascending' | 'descending' = 'ascending'): SortKey {
  return { column, direction };
}

function dispatchSort(
  session: DocumentSession,
  sheet: Sheet,
  selection: GridRange,
  args: SortCallArgs,
) {
  return session.dispatch('data.sort', args, { activeSheetId: sheet.id, selection });
}

interface SortCallArgs {
  readonly keys?: readonly SortKey[];
  readonly hasHeaders?: boolean | null;
  readonly range?: GridRange;
}

/**
 * Everything the sort must restore, deep enough that a lost style, formula,
 * merge or attachment fails the comparison.
 */
function sheetState(sheet: Sheet) {
  return {
    cells: [...sheet.cells.entries()]
      .sort(([left], [right]) => left.localeCompare(right))
      .map(([address, cell]) => [address, cloneCell(cell)] as const),
    merges: sheet.merges.map((merge) => ({ ...merge.range })),
    comments: [...sheet.comments.entries()]
      .sort(([left], [right]) => left.localeCompare(right))
      .map(([address, record]) => [address, { ...record }] as const),
    hyperlinks: [...sheet.hyperlinks.entries()]
      .sort(([left], [right]) => left.localeCompare(right))
      .map(([address, record]) => [address, { ...record }] as const),
    rows: [...sheet.rowDimensions.entries()]
      .sort(([left], [right]) => left - right)
      .map(([row, dimension]) => [row, { ...dimension }] as const),
  };
}

/**
 * S15's workbook: Name/Qty/Calc/Tag columns, an Alice row that moves down and a
 * Dave row that moves up, a bold style, a one-row merge, a comment, a hyperlink
 * and a row height. Bob and Carol keep their positions, so the test also proves
 * a stable permutation leaves settled rows alone.
 */
function sortUndoWorkbook() {
  const { workbook, sheet } = workbookWithSheet();
  const bold = appendStyle(workbook, {
    numFmtId: 0,
    fontId: 1,
    fillId: 0,
    borderId: 0,
    xfId: 0,
    rawXml: '',
    font: { ...DEFAULT_FONT, bold: true },
  });

  for (const [col, label] of ['Name', 'Qty', 'Calc', 'Extra', 'Tag', 'Tag2'].entries()) {
    put(sheet, 0, col, textValue(label));
  }
  put(sheet, 1, 0, textValue('Alice'), bold);
  put(sheet, 1, 1, numberValue(30), bold);
  putFormula(sheet, 1, 2, '=B2*2');
  sheet.cells.get(cellKey(1, 2))!.styleId = bold;
  putFormula(sheet, 2, 2, '=B3*2');
  putFormula(sheet, 3, 2, '=B4*2');
  putFormula(sheet, 4, 2, '=B5*2');
  put(sheet, 2, 0, textValue('Bob'));
  put(sheet, 2, 1, numberValue(10));
  put(sheet, 3, 0, textValue('Carol'));
  put(sheet, 3, 1, numberValue(20));
  put(sheet, 4, 0, textValue('Dave'));
  put(sheet, 4, 1, numberValue(5));

  put(sheet, 1, 4, textValue('tagged'));
  sheet.merges.push({ range: range(1, 4, 1, 5) });
  sheet.comments.set(cellKey(1, 0), { text: 'Alice note', threaded: false });
  sheet.hyperlinks.set(cellKey(4, 1), { target: 'https://example.com/dave' });
  sheet.rowDimensions.set(2, { height: 30 });
  return { workbook, sheet, bold };
}

function bootstrapSheet() {
  const { workbook, sheet } = workbookWithSheet();
  for (const [col, label] of ['Name', 'Qty'].entries()) put(sheet, 0, col, textValue(label));
  put(sheet, 1, 0, textValue('Alice'));
  put(sheet, 1, 1, numberValue(30));
  put(sheet, 2, 0, textValue('Bob'));
  put(sheet, 2, 1, numberValue(10));
  put(sheet, 3, 0, textValue('Carol'));
  put(sheet, 3, 1, numberValue(20));
  return { workbook, sheet };
}

describe('S15 sort-undo', () => {
  it('applies one transaction and undo and redo restore the exact state', () => {
    const { workbook, sheet, bold } = sortUndoWorkbook();
    const session = new DocumentSession(workbook);
    const before = sheetState(sheet);

    const result = dispatchSort(session, sheet, range(0, 0, 4, 5), {
      keys: [key(1)],
      hasHeaders: true,
    });

    expect(result.sort?.state).toBe('sorted');
    expect(result.sort?.hasHeaders).toBe(true);
    expect(result.sort?.keyLabels).toEqual(['Qty']);
    expect(result.changes.filter((change) => change.type === 'sort.apply')).toHaveLength(1);
    expect(session.transactionCount).toBe(1);

    // Values, styles and formulas move with their rows.
    expect(getCell(sheet, 1, 0)?.value).toEqual(textValue('Dave'));
    expect(getCell(sheet, 2, 0)?.value).toEqual(textValue('Bob'));
    expect(getCell(sheet, 3, 0)?.value).toEqual(textValue('Carol'));
    expect(getCell(sheet, 4, 0)?.value).toEqual(textValue('Alice'));
    expect(getCell(sheet, 4, 0)?.styleId).toBe(bold);
    expect(getCell(sheet, 4, 1)?.styleId).toBe(bold);
    // The moved formula shifts its relative references by the row delta.
    expect(getCell(sheet, 4, 2)?.formula).toBe('=B5*2');
    expect(getCell(sheet, 1, 2)?.formula).toBe('=B2*2');

    // The single-row merge moved with Alice, and the attachments followed.
    expect(sheet.merges).toEqual([{ range: range(4, 4, 4, 5) }]);
    expect(sheet.comments.get(cellKey(4, 0))).toEqual({ text: 'Alice note', threaded: false });
    expect(sheet.comments.has(cellKey(1, 0))).toBe(false);
    expect(sheet.hyperlinks.get(cellKey(1, 1))).toEqual({ target: 'https://example.com/dave' });
    expect(sheet.hyperlinks.has(cellKey(4, 1))).toBe(false);
    // Row heights stay with positions, never with content.
    expect(sheet.rowDimensions.get(2)).toEqual({ height: 30 });

    const after = sheetState(sheet);

    session.dispatch('edit.undo', undefined);
    expect(sheetState(sheet)).toEqual(before);
    expect(session.transactionCount).toBe(0);
    expect(session.canRedo).toBe(true);

    session.dispatch('edit.redo', undefined);
    expect(sheetState(sheet)).toEqual(after);
    expect(session.transactionCount).toBe(1);
  });

  it('expands a single-cell selection to the current region before sorting', () => {
    const { workbook, sheet } = sortUndoWorkbook();
    const session = new DocumentSession(workbook);

    const result = dispatchSort(session, sheet, range(0, 0, 0, 0), { keys: [key(1)] });

    expect(result.sort?.range).toEqual(range(0, 0, 4, 5));
    expect(result.sort?.state).toBe('sorted');
    expect(getCell(sheet, 1, 1)?.value).toEqual(numberValue(5));
  });
});

describe('sort blocking outcomes', () => {
  it('blocks mixed merges with Excel’s exact message and records nothing', () => {
    const { workbook, sheet } = bootstrapSheet();
    sheet.merges.push({ range: range(1, 0, 2, 1) });
    const session = new DocumentSession(workbook);

    const result = dispatchSort(session, sheet, range(0, 0, 3, 1), { keys: [key(0)] });

    expect(result.changes).toHaveLength(0);
    expect(result.ui).toEqual({ type: 'toast', message: MERGED_SIZE_MESSAGE });
    expect(result.sort?.state).toBe('blocked');
    expect(result.sort?.code).toBe('merged-cells');
    expect(session.transactionCount).toBe(0);
  });

  it('blocks a legacy array formula with Excel’s exact message', () => {
    const { workbook, sheet } = bootstrapSheet();
    setCellSnapshot(sheet, 1, 0, {
      value: numberValue(3),
      formula: 'SUM(B2:B3)',
      formulaDerived: false,
      savedResult: numberValue(3),
      styleId: 0,
      richText: null,
      formulaSource: { kind: 'array', range: 'A2:B3', xml: '<f t="array" ref="A2:B3"/>' },
    });
    const session = new DocumentSession(workbook);

    const result = dispatchSort(session, sheet, range(0, 0, 3, 1), { keys: [key(1)] });

    expect(result.ui).toEqual({ type: 'toast', message: ARRAY_BLOCK_MESSAGE });
    expect(result.sort?.code).toBe('array-formula');
    expect(session.transactionCount).toBe(0);
  });

  it('keeps the table block message available to the command layer', () => {
    // Tables are preserved as opaque parts until the table model lands, so the
    // command cannot be handed a table range yet; the contract's message is
    // still exported and asserted here so the wiring cannot drift.
    expect(TABLE_BLOCK_MESSAGE).toBe('A sort cannot include cells that belong to a table.');
  });

  it('blocks more than 64 keys and a key outside the range with their messages', () => {
    const { workbook, sheet } = bootstrapSheet();
    const session = new DocumentSession(workbook);
    const keys: SortKey[] = [];
    for (let index = 0; index <= MAX_SORT_KEYS; index += 1) keys.push(key(0));

    const tooMany = dispatchSort(session, sheet, range(0, 0, 3, 1), { keys });
    expect(tooMany.ui).toEqual({ type: 'toast', message: TOO_MANY_KEYS_MESSAGE });
    expect(tooMany.sort?.code).toBe('too-many-keys');

    const outside = dispatchSort(session, sheet, range(0, 0, 3, 1), { keys: [key(4)] });
    expect(outside.ui).toEqual({ type: 'toast', message: KEY_OUTSIDE_RANGE_MESSAGE });
    expect(outside.sort?.code).toBe('key-outside-range');
    expect(session.transactionCount).toBe(0);
  });
});

describe('sort no-op outcomes', () => {
  it('asks the Sort dialog for keys instead of guessing them', () => {
    const { workbook, sheet } = bootstrapSheet();
    const session = new DocumentSession(workbook);

    const result = dispatchSort(session, sheet, range(0, 0, 3, 1), {});

    expect(result.changes).toHaveLength(0);
    expect(result.ui?.type).toBe('dialog');
    expect(result.ui?.message).toMatch(/Choose at least one sort key/);
    expect(result.sort?.state).toBe('needs-keys');
    expect(result.sort?.range).toEqual(range(0, 0, 3, 1));
    expect(session.transactionCount).toBe(0);
  });

  it('creates no transaction when the range has fewer than two data rows', () => {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, textValue('Name'));
    put(sheet, 0, 1, textValue('Qty'));
    put(sheet, 1, 0, textValue('Alice'));
    put(sheet, 1, 1, numberValue(30));
    const session = new DocumentSession(workbook);

    const result = dispatchSort(session, sheet, range(0, 0, 1, 1), {
      keys: [key(1)],
      hasHeaders: true,
    });

    expect(result.changes).toHaveLength(0);
    expect(result.sort?.state).toBe('unchanged');
    expect(result.sort?.reason).toBe('too-few-rows');
    expect(session.transactionCount).toBe(0);
  });

  it('creates no transaction when the rows are already in order', () => {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, textValue('Name'));
    put(sheet, 0, 1, textValue('Qty'));
    put(sheet, 1, 0, textValue('Alice'));
    put(sheet, 1, 1, numberValue(10));
    put(sheet, 2, 0, textValue('Bob'));
    put(sheet, 2, 1, numberValue(20));
    const session = new DocumentSession(workbook);

    const result = dispatchSort(session, sheet, range(0, 0, 2, 1), { keys: [key(1)] });

    expect(result.changes).toHaveLength(0);
    expect(result.sort?.state).toBe('unchanged');
    expect(result.sort?.reason).toBe('nothing-to-order');
    expect(session.transactionCount).toBe(0);
  });
});

describe('sort while filtered (§3.4)', () => {
  it('reorders visible rows only and leaves the filter state and visible set alone', () => {
    const { workbook, sheet } = bootstrapSheet();
    sheet.autoFilter = createAutoFilter(sheet.id, range(0, 0, 3, 1), [
      [
        1,
        {
          kind: 'values',
          selected: [numberValue(30), numberValue(20)],
          includeBlanks: false,
        },
      ],
    ]);
    const hiddenBefore = computeHiddenLayers(sheet, sheet.autoFilter).effectiveHidden;
    expect([...hiddenBefore]).toEqual([2]);
    const session = new DocumentSession(workbook);

    const result = dispatchSort(session, sheet, range(0, 0, 3, 1), {
      keys: [key(1)],
      hasHeaders: true,
    });

    expect(result.sort?.state).toBe('sorted');
    // Carol (20) and Alice (30) swap; Bob's filter-hidden row stays put.
    expect(getCell(sheet, 1, 0)?.value).toEqual(textValue('Carol'));
    expect(getCell(sheet, 2, 0)?.value).toEqual(textValue('Bob'));
    expect(getCell(sheet, 3, 0)?.value).toEqual(textValue('Alice'));
    expect(sheet.autoFilter?.columns.get(1)).toEqual({
      kind: 'values',
      selected: [numberValue(30), numberValue(20)],
      includeBlanks: false,
    });
    const hiddenAfter = computeHiddenLayers(sheet, sheet.autoFilter).effectiveHidden;
    expect([...hiddenAfter]).toEqual([2]);
    expect(session.dispatch('edit.undo', undefined).changes).toHaveLength(0);
    expect(getCell(sheet, 1, 0)?.value).toEqual(textValue('Alice'));
  });
});

describe('sort warnings', () => {
  it('warns once about hidden rows and leaves them in place', () => {
    const { workbook, sheet } = sortUndoWorkbook();
    sheet.rowDimensions.set(2, { height: 30, hidden: true });
    const session = new DocumentSession(workbook);

    const result = dispatchSort(session, sheet, range(0, 0, 4, 5), {
      keys: [key(1)],
      hasHeaders: true,
    });

    expect(result.sort?.warnings).toEqual([
      { code: 'hidden-rows', message: HIDDEN_ROWS_WARNING, count: 1 },
    ]);
    // Bob's row is hidden and stays exactly where it was.
    expect(getCell(sheet, 2, 0)?.value).toEqual(textValue('Bob'));
    expect(sheet.rowDimensions.get(2)?.hidden).toBe(true);
    expect(session.transactionCount).toBe(1);
  });

  it('warns once when the selection is a sub-range of a larger block', () => {
    const { workbook, sheet } = workbookWithSheet();
    for (const [col, label] of ['Name', 'Qty', 'Note'].entries()) {
      put(sheet, 0, col, textValue(label));
    }
    put(sheet, 1, 0, textValue('Alice'));
    put(sheet, 1, 1, numberValue(30));
    put(sheet, 1, 2, textValue('a'));
    put(sheet, 2, 0, textValue('Bob'));
    put(sheet, 2, 1, numberValue(10));
    put(sheet, 2, 2, textValue('b'));
    put(sheet, 3, 0, textValue('Carol'));
    put(sheet, 3, 1, numberValue(20));
    put(sheet, 3, 2, textValue('c'));
    const session = new DocumentSession(workbook);

    const result = dispatchSort(session, sheet, range(0, 0, 3, 1), { keys: [key(1)] });

    expect(result.sort?.warnings).toEqual([{ code: 'sub-range', message: SUB_RANGE_WARNING }]);
    expect(result.sort?.state).toBe('sorted');
  });
});
