import { describe, expect, it } from 'vitest';
import { DocumentSession } from '@commands/session';
import type { Cell, GridRange, Sheet } from '@model/types';
import { cellKey, cloneCell, setCellSnapshot } from '@model/types';
import { numberValue, textValue } from '@model/value';
import { createEmptyWorkbook, createSheet } from '@model/workbook';

/**
 * The two drag commands (ticket 49, feature matrix §1 "Drag-and-drop move;
 * Shift-drag insert"): each gesture is one transaction that moves the cells,
 * rewrites every reference that points at a moved cell and undoes in one step.
 * `tests/e2e/drag-move.spec.ts` owns the gestures themselves.
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

function put(sheet: Sheet, row: number, col: number, value: ReturnType<typeof textValue>): void {
  setCellSnapshot(sheet, row, col, {
    value,
    formula: null,
    formulaDerived: false,
    savedResult: null,
    styleId: 0,
    richText: null,
  });
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

function cellsOf(sheet: Sheet) {
  return [...sheet.cells.entries()]
    .sort(([left], [right]) => left.localeCompare(right))
    .map(([address, cell]) => [address, cloneCell(cell)] as const);
}

function namesOf(session: DocumentSession) {
  return session.workbook.names.map((name) => ({ ...name }));
}

function dispatch(
  session: DocumentSession,
  id: string,
  sheet: Sheet,
  args: Record<string, unknown>,
) {
  return session.dispatch(id, args, {
    activeSheetId: sheet.id,
    selection: null,
    activeCell: { row: 0, col: 0 },
  });
}

// ---------------------------------------------------------------------------
// edit.dragMove
// ---------------------------------------------------------------------------

describe('edit.dragMove', () => {
  it('moves the range, rewrites references and is one undo transaction', () => {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, numberValue(1));
    put(sheet, 1, 0, numberValue(2));
    putFormula(sheet, 1, 1, 'A2*10');
    workbook.names.push({ name: 'Anchor', formula: 'Sheet1!$A$2', sheetId: null });
    const session = new DocumentSession(workbook);
    const cellsBefore = cellsOf(sheet);
    const namesBefore = namesOf(session);

    // A2:B2 moves to A5. The formula in B2 rides along and its reference to
    // the carried A2 follows; the defined name follows too.
    const result = dispatch(session, 'edit.dragMove', sheet, {
      source: range(1, 0, 1, 1),
      to: { row: 4, col: 0 },
    });
    expect(result.structure?.state).toBe('applied');
    expect(session.transactionCount).toBe(1);

    expect(sheet.cells.get(cellKey(1, 0))).toBeUndefined();
    expect(sheet.cells.get(cellKey(1, 1))).toBeUndefined();
    expect(sheet.cells.get(cellKey(4, 0))?.value).toEqual(numberValue(2));
    expect(sheet.cells.get(cellKey(4, 1))?.formula).toBe('A5*10');
    expect(session.workbook.names[0]?.formula).toBe('Sheet1!$A$5');
    expect(result.selectionAfter).toEqual(range(4, 0, 4, 1));

    const cellsAfter = cellsOf(sheet);
    const namesAfter = namesOf(session);
    dispatch(session, 'edit.undo', sheet, {});
    expect(cellsOf(sheet)).toEqual(cellsBefore);
    expect(namesOf(session)).toEqual(namesBefore);
    expect(session.transactionCount).toBe(0);

    dispatch(session, 'edit.redo', sheet, {});
    expect(cellsOf(sheet)).toEqual(cellsAfter);
    expect(namesOf(session)).toEqual(namesAfter);
  });

  it('rewrites formulas on other sheets that point at the moved cells', () => {
    const { workbook, sheet } = workbookWithSheet();
    const other = createSheet('Sheet2', 'xl/worksheets/sheet2.xml');
    workbook.sheets.push(other);
    put(sheet, 0, 0, numberValue(5));
    putFormula(other, 0, 0, 'Sheet1!A1*2');
    putFormula(other, 1, 0, 'Sheet1!B1+1');
    const session = new DocumentSession(workbook);

    dispatch(session, 'edit.dragMove', sheet, {
      source: range(0, 0, 0, 0),
      to: { row: 2, col: 2 },
    });
    // The reference to the carried A1 follows to C3; the reference to the
    // untouched B1 keeps its authored text.
    expect(other.cells.get(cellKey(0, 0))?.formula).toBe('Sheet1!C3*2');
    expect(other.cells.get(cellKey(1, 0))?.formula).toBe('Sheet1!B1+1');

    dispatch(session, 'edit.undo', sheet, {});
    expect(other.cells.get(cellKey(0, 0))?.formula).toBe('Sheet1!A1*2');
    expect(other.cells.get(cellKey(1, 0))?.formula).toBe('Sheet1!B1+1');
  });

  it('keeps a moved formula whose reference points outside the carried cells', () => {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, numberValue(4));
    putFormula(sheet, 1, 0, 'A1+1');
    const session = new DocumentSession(workbook);

    // B1's formula points at A1, which stays put, so the authored text is
    // unchanged — Excel's rule for a moved formula.
    dispatch(session, 'edit.dragMove', sheet, {
      source: range(1, 0, 1, 0),
      to: { row: 3, col: 2 },
    });
    expect(sheet.cells.get(cellKey(3, 2))?.formula).toBe('A1+1');
  });

  it('a drop back on the source changes nothing', () => {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, numberValue(1));
    put(sheet, 1, 0, numberValue(2));
    const session = new DocumentSession(workbook);
    const cellsBefore = cellsOf(sheet);

    const result = dispatch(session, 'edit.dragMove', sheet, {
      source: range(0, 0, 1, 0),
      to: { row: 0, col: 0 },
    });
    expect(result.changes).toHaveLength(0);
    expect(session.transactionCount).toBe(0);
    expect(cellsOf(sheet)).toEqual(cellsBefore);
  });

  it('overwrites the destination cells and leaves references to them alone', () => {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, numberValue(1));
    put(sheet, 3, 0, numberValue(99));
    putFormula(sheet, 4, 0, 'A4*2');
    const session = new DocumentSession(workbook);

    dispatch(session, 'edit.dragMove', sheet, {
      source: range(0, 0, 0, 0),
      to: { row: 3, col: 0 },
    });
    expect(sheet.cells.get(cellKey(3, 0))?.value).toEqual(numberValue(1));
    // The overwritten A4 is not a reference rewrite: the formula keeps its
    // text and now reads the moved value, exactly like any cell edit.
    expect(sheet.cells.get(cellKey(4, 0))?.formula).toBe('A4*2');

    dispatch(session, 'edit.undo', sheet, {});
    expect(sheet.cells.get(cellKey(3, 0))?.value).toEqual(numberValue(99));
  });

  it('a drop that carries nothing records nothing', () => {
    const { workbook, sheet } = workbookWithSheet();
    const session = new DocumentSession(workbook);

    const result = dispatch(session, 'edit.dragMove', sheet, {
      source: range(5, 5, 6, 6),
      to: { row: 0, col: 0 },
    });
    expect(result.changes).toHaveLength(0);
    expect(result.selectionAfter).toEqual(range(0, 0, 1, 1));
    expect(session.transactionCount).toBe(0);
  });

  it('blocks a move that would split a merged cell', () => {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, numberValue(1));
    put(sheet, 0, 1, numberValue(2));
    sheet.merges.push({ range: range(0, 0, 0, 1) });
    const session = new DocumentSession(workbook);
    const cellsBefore = cellsOf(sheet);

    const result = dispatch(session, 'edit.dragMove', sheet, {
      source: range(0, 0, 0, 0),
      to: { row: 3, col: 0 },
    });
    expect(result.structure?.state).toBe('blocked');
    expect(result.ui?.message).toBe('Cannot change part of a merged cell.');
    expect(session.transactionCount).toBe(0);
    expect(cellsOf(sheet)).toEqual(cellsBefore);
  });

  it('moves a whole merge with its cells', () => {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, numberValue(1));
    sheet.merges.push({ range: range(0, 0, 1, 1) });
    const session = new DocumentSession(workbook);

    const result = dispatch(session, 'edit.dragMove', sheet, {
      source: range(0, 0, 1, 1),
      to: { row: 4, col: 3 },
    });
    expect(result.structure?.state).toBe('applied');
    expect(sheet.merges[0]?.range).toEqual(range(4, 3, 5, 4));
    dispatch(session, 'edit.undo', sheet, {});
    expect(sheet.merges[0]?.range).toEqual(range(0, 0, 1, 1));
  });

  it('blocks a shared formula that would move', () => {
    const { workbook, sheet } = workbookWithSheet();
    const cell = putFormula(sheet, 0, 0, 'B1+1');
    setCellSnapshot(sheet, 0, 0, {
      ...cell,
      formulaSource: { kind: 'shared-master', sharedIndex: '0', xml: '<f t="shared" si="0"/>' },
    });
    const session = new DocumentSession(workbook);

    const result = dispatch(session, 'edit.dragMove', sheet, {
      source: range(0, 0, 0, 0),
      to: { row: 3, col: 0 },
    });
    expect(result.structure?.state).toBe('blocked');
    expect(result.ui?.message).toContain('Shared and array formulas');
    expect(session.transactionCount).toBe(0);
  });

  it('moves a comment and a hyperlink with their cells and follows the link target', () => {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, numberValue(1));
    put(sheet, 5, 0, numberValue(2));
    sheet.comments.set(cellKey(0, 0), { text: 'note', threaded: false });
    sheet.hyperlinks.set(cellKey(5, 0), {
      target: '',
      location: 'Sheet1!A1',
    });
    const session = new DocumentSession(workbook);

    dispatch(session, 'edit.dragMove', sheet, {
      source: range(0, 0, 0, 0),
      to: { row: 2, col: 2 },
    });
    expect(sheet.comments.has(cellKey(0, 0))).toBe(false);
    expect(sheet.comments.get(cellKey(2, 2))?.text).toBe('note');
    expect(sheet.hyperlinks.get(cellKey(5, 0))?.location).toBe('Sheet1!C3');

    dispatch(session, 'edit.undo', sheet, {});
    expect(sheet.comments.get(cellKey(0, 0))?.text).toBe('note');
    expect(sheet.hyperlinks.get(cellKey(5, 0))?.location).toBe('Sheet1!A1');
  });
});

// ---------------------------------------------------------------------------
// edit.dragInsert
// ---------------------------------------------------------------------------

describe('edit.dragInsert', () => {
  it('inserts and shifts cells down, with reference rewriting and one undo', () => {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, numberValue(1));
    put(sheet, 2, 0, numberValue(3));
    put(sheet, 3, 0, numberValue(4));
    put(sheet, 2, 1, numberValue(30));
    putFormula(sheet, 2, 2, 'A3*2');
    const session = new DocumentSession(workbook);
    const cellsBefore = cellsOf(sheet);

    // A1 drops at A3: the gap at A3 pushes A3 down inside column A only, and
    // the carried A1 lands in the gap. The formula that read the shifted A3
    // follows it to A4.
    const result = dispatch(session, 'edit.dragInsert', sheet, {
      source: range(0, 0, 0, 0),
      to: { row: 2, col: 0 },
      shift: 'down',
    });
    expect(result.structure?.state).toBe('applied');
    expect(session.transactionCount).toBe(1);

    expect(sheet.cells.get(cellKey(0, 0))).toBeUndefined();
    expect(sheet.cells.get(cellKey(2, 0))?.value).toEqual(numberValue(1));
    expect(sheet.cells.get(cellKey(3, 0))?.value).toEqual(numberValue(3));
    expect(sheet.cells.get(cellKey(4, 0))?.value).toEqual(numberValue(4));
    // The band is column A: B3 keeps its place.
    expect(sheet.cells.get(cellKey(2, 1))?.value).toEqual(numberValue(30));
    expect(sheet.cells.get(cellKey(2, 2))?.formula).toBe('A4*2');

    dispatch(session, 'edit.undo', sheet, {});
    expect(cellsOf(sheet)).toEqual(cellsBefore);
  });

  it('inserts and shifts cells right', () => {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, numberValue(1));
    put(sheet, 0, 1, numberValue(2));
    put(sheet, 0, 3, numberValue(4));
    putFormula(sheet, 1, 3, 'D1+1');
    const session = new DocumentSession(workbook);
    const cellsBefore = cellsOf(sheet);

    // A1:B1 drops at D1: the gap at D1:E1 pushes the old D1 right inside row 1
    // only, and the carried cells land in the gap.
    const result = dispatch(session, 'edit.dragInsert', sheet, {
      source: range(0, 0, 0, 1),
      to: { row: 0, col: 3 },
      shift: 'right',
    });
    expect(result.structure?.state).toBe('applied');
    expect(session.transactionCount).toBe(1);

    expect(sheet.cells.get(cellKey(0, 0))).toBeUndefined();
    expect(sheet.cells.get(cellKey(0, 1))).toBeUndefined();
    expect(sheet.cells.get(cellKey(0, 3))?.value).toEqual(numberValue(1));
    expect(sheet.cells.get(cellKey(0, 4))?.value).toEqual(numberValue(2));
    expect(sheet.cells.get(cellKey(0, 5))?.value).toEqual(numberValue(4));
    // Row 2 is outside the band: its formula keeps its text but the value it
    // points at moved from D1 to F1.
    expect(sheet.cells.get(cellKey(1, 3))?.formula).toBe('F1+1');

    dispatch(session, 'edit.undo', sheet, {});
    expect(cellsOf(sheet)).toEqual(cellsBefore);
  });

  it('a drop back on the dragged range inserts nothing', () => {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, numberValue(1));
    put(sheet, 0, 1, numberValue(2));
    const session = new DocumentSession(workbook);
    const cellsBefore = cellsOf(sheet);

    const result = dispatch(session, 'edit.dragInsert', sheet, {
      source: range(0, 0, 0, 1),
      to: { row: 0, col: 1 },
      shift: 'down',
    });
    expect(result.changes).toHaveLength(0);
    expect(session.transactionCount).toBe(0);
    expect(cellsOf(sheet)).toEqual(cellsBefore);
  });

  it('carries a defined name to the gap and shifts the name it does not carry', () => {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, numberValue(1));
    put(sheet, 2, 0, numberValue(3));
    workbook.names.push({ name: 'Carried', formula: 'Sheet1!$A$1', sheetId: null });
    workbook.names.push({ name: 'Stayed', formula: 'Sheet1!$A$3', sheetId: null });
    const session = new DocumentSession(workbook);

    dispatch(session, 'edit.dragInsert', sheet, {
      source: range(0, 0, 0, 0),
      to: { row: 2, col: 0 },
      shift: 'down',
    });
    expect(session.workbook.names[0]?.formula).toBe('Sheet1!$A$3');
    // The name that pointed at the shifted A3 follows it to A4.
    expect(session.workbook.names[1]?.formula).toBe('Sheet1!$A$4');
  });
});
