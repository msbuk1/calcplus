import { readFileSync } from 'node:fs';
import { Buffer } from 'node:buffer';
import { describe, expect, it } from 'vitest';
import { DocumentSession } from '@commands/session';
import { createAutoFilter, evaluateAutoFilter } from '@model/autofilter';
import type { Cell, GridRange, Sheet } from '@model/types';
import { cellKey, cloneCell, setCellSnapshot } from '@model/types';
import { numberValue, textValue } from '@model/value';
import { createEmptyWorkbook, createSheet } from '@model/workbook';
import { openWorkbook } from '@ooxml/open';
import { saveWorkbook } from '@ooxml/save';
import { buildXlsx, directChildNames, elementRanges, unzipPart } from './helpers/build-xlsx';

/**
 * The six structural commands through the registry: each action is ONE
 * transaction (M5 exit criterion 4), undo restores the exact prior state and
 * redo reapplies it. Fixtures F11, F12, CF17 and CF20 run through the command
 * layer here; `m5-structure-transform.test.ts` owns the transform detail.
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

/** Everything a structural edit can touch, deep enough to catch a lost field. */
function sheetState(sheet: Sheet) {
  return {
    cells: [...sheet.cells.entries()]
      .sort(([left], [right]) => left.localeCompare(right))
      .map(([address, cell]) => [address, cloneCell(cell)] as const),
    rows: [...sheet.rowDimensions.entries()].sort((left, right) => left[0] - right[0]),
    columns: [...sheet.columnDimensions.entries()].sort((left, right) => left[0] - right[0]),
    merges: sheet.merges.map((merge) => ({ ...merge.range })),
    comments: [...sheet.comments.entries()].map(([key, record]) => [key, { ...record }] as const),
    hyperlinks: [...sheet.hyperlinks.entries()].map(
      ([key, record]) => [key, { ...record }] as const,
    ),
    validation: sheet.validation.map((record) => ({ ...record, range: { ...record.range } })),
    conditionalFormats: sheet.conditionalFormats.map((record) => ({
      ...record,
      range: { ...record.range },
    })),
    autoFilter: sheet.autoFilter === null ? null : filterState(sheet.autoFilter),
  };
}

function filterState(filter: NonNullable<Sheet['autoFilter']>) {
  return {
    range: { ...filter.range },
    columns: [...filter.columns.entries()]
      .sort((left, right) => left[0] - right[0])
      .map(([offset, column]) => [offset, column] as const),
  };
}

function sourceSnapshot(session: DocumentSession, sheet: Sheet): Uint8Array | null {
  return session.workbook.sourceParts.get(sheet.partPath) ?? null;
}

function sameBytes(left: Uint8Array | null, right: Uint8Array | null): boolean {
  if (left === null || right === null) return left === right;
  return Buffer.from(left).equals(Buffer.from(right));
}

function dispatch(
  session: DocumentSession,
  id: string,
  sheet: Sheet,
  selection: GridRange | null,
  args: Record<string, unknown> = {},
) {
  return session.dispatch(id, args, {
    activeSheetId: sheet.id,
    selection,
    activeCell: { row: selection?.startRow ?? 0, col: selection?.startCol ?? 0 },
  });
}

// ---------------------------------------------------------------------------
// One transaction per command, undo and redo
// ---------------------------------------------------------------------------

describe('structural commands are one transaction each', () => {
  it('inserts the selected number of rows above the first selected row', () => {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, numberValue(1));
    put(sheet, 1, 0, numberValue(2));
    putFormula(sheet, 1, 1, 'A2*10');
    put(sheet, 2, 0, numberValue(3));
    const session = new DocumentSession(workbook);
    const before = sheetState(sheet);
    const sourceBefore = sourceSnapshot(session, sheet);

    // Select two rows; Insert Rows Above inserts two rows above the first.
    const result = dispatch(session, 'structure.insertRows', sheet, range(1, 0, 2, 0));
    expect(result.structure?.state).toBe('applied');
    expect(session.transactionCount).toBe(1);

    expect(sheet.cells.get(cellKey(0, 0))?.value).toEqual(numberValue(1));
    expect(sheet.cells.get(cellKey(1, 0))).toBeUndefined();
    expect(sheet.cells.get(cellKey(2, 0))).toBeUndefined();
    expect(sheet.cells.get(cellKey(3, 0))?.value).toEqual(numberValue(2));
    expect(sheet.cells.get(cellKey(3, 1))?.formula).toBe('A4*10');
    expect(sheet.cells.get(cellKey(4, 0))?.value).toEqual(numberValue(3));
    expect(result.selectionAfter).toEqual(range(1, 0, 2, 0));

    const after = sheetState(sheet);
    session.dispatch('edit.undo', {}, { activeSheetId: sheet.id });
    expect(sheetState(sheet)).toEqual(before);
    expect(sameBytes(sourceSnapshot(session, sheet), sourceBefore)).toBe(true);

    session.dispatch('edit.redo', {}, { activeSheetId: sheet.id });
    expect(sheetState(sheet)).toEqual(after);
  });

  it('deletes the selected rows and drops the content, merges, comments and hyperlinks', () => {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, textValue('header'));
    put(sheet, 1, 0, textValue('gone'));
    put(sheet, 2, 0, textValue('kept'));
    sheet.merges.push({ range: range(1, 0, 1, 1) });
    sheet.comments.set(cellKey(1, 1), { text: 'note', threaded: false });
    sheet.hyperlinks.set(cellKey(2, 0), { target: 'https://example.com' });
    sheet.rowDimensions.set(2, { height: 30 });
    const session = new DocumentSession(workbook);
    const before = sheetState(sheet);

    const result = dispatch(session, 'structure.deleteRows', sheet, range(1, 0, 1, 0));
    expect(result.structure?.state).toBe('applied');
    expect(session.transactionCount).toBe(1);
    expect(sheet.cells.get(cellKey(1, 0))?.value).toEqual(textValue('kept'));
    expect(sheet.merges).toEqual([]);
    expect(sheet.comments.size).toBe(0);
    expect(sheet.hyperlinks.get(cellKey(1, 0))?.target).toBe('https://example.com');
    expect(sheet.rowDimensions.get(1)).toEqual({ height: 30 });

    session.dispatch('edit.undo', {}, { activeSheetId: sheet.id });
    expect(sheetState(sheet)).toEqual(before);
  });

  it('inserts columns and renumbers the column width records', () => {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, textValue('a'));
    put(sheet, 0, 1, textValue('b'));
    sheet.columnDimensions.set(1, { width: 20, hidden: true });
    const session = new DocumentSession(workbook);
    const before = sheetState(sheet);

    dispatch(session, 'structure.insertColumns', sheet, range(0, 1, 0, 1));
    expect(sheet.cells.get(cellKey(0, 2))?.value).toEqual(textValue('b'));
    expect(sheet.columnDimensions.get(2)).toEqual({ width: 20, hidden: true });
    expect(session.transactionCount).toBe(1);

    session.dispatch('edit.undo', {}, { activeSheetId: sheet.id });
    expect(sheetState(sheet)).toEqual(before);
  });

  it('shifts only the selected rows for Insert Cells & Shift Right', () => {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, textValue('a'));
    put(sheet, 0, 1, textValue('b'));
    put(sheet, 4, 0, textValue('outside'));
    const session = new DocumentSession(workbook);
    const before = sheetState(sheet);

    const result = dispatch(session, 'structure.insertCells', sheet, range(0, 0, 0, 0), {
      shift: 'right',
    });
    expect(result.structure?.state).toBe('applied');
    expect(sheet.cells.get(cellKey(0, 0))).toBeUndefined();
    expect(sheet.cells.get(cellKey(0, 1))?.value).toEqual(textValue('a'));
    expect(sheet.cells.get(cellKey(0, 2))?.value).toEqual(textValue('b'));
    expect(sheet.cells.get(cellKey(4, 0))?.value).toEqual(textValue('outside'));
    expect(session.transactionCount).toBe(1);

    session.dispatch('edit.undo', {}, { activeSheetId: sheet.id });
    expect(sheetState(sheet)).toEqual(before);
  });

  it('shifts cells up and rewrites a collapsed reference to #REF! on Delete Cells', () => {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, textValue('a'));
    put(sheet, 1, 0, textValue('b'));
    put(sheet, 2, 0, textValue('c'));
    putFormula(sheet, 0, 2, 'A1+A3');
    const session = new DocumentSession(workbook);

    const result = dispatch(session, 'structure.deleteCells', sheet, range(0, 0, 0, 0), {
      shift: 'up',
    });
    expect(result.structure?.state).toBe('applied');
    expect(sheet.cells.get(cellKey(0, 0))?.value).toEqual(textValue('b'));
    expect(sheet.cells.get(cellKey(1, 0))?.value).toEqual(textValue('c'));
    expect(sheet.cells.get(cellKey(0, 2))?.formula).toBe('#REF!+A2');
    expect(session.transactionCount).toBe(1);
  });

  it('deletes the selected columns and drops the content below', () => {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, textValue('a'));
    put(sheet, 0, 1, textValue('b'));
    put(sheet, 0, 2, textValue('c'));
    const session = new DocumentSession(workbook);
    const before = sheetState(sheet);

    dispatch(session, 'structure.deleteColumns', sheet, range(0, 1, 0, 1));
    expect(sheet.cells.get(cellKey(0, 0))?.value).toEqual(textValue('a'));
    expect(sheet.cells.get(cellKey(0, 1))?.value).toEqual(textValue('c'));
    expect(session.transactionCount).toBe(1);

    session.dispatch('edit.undo', {}, { activeSheetId: sheet.id });
    expect(sheetState(sheet)).toEqual(before);
  });
});

// ---------------------------------------------------------------------------
// Target resolution and blocking
// ---------------------------------------------------------------------------

describe('structural target resolution and blocking', () => {
  it('resolves the range from the command context when args omit it (ticket 29)', () => {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, textValue('a'));
    put(sheet, 1, 0, textValue('b'));
    const session = new DocumentSession(workbook);
    dispatch(session, 'structure.insertRows', sheet, range(1, 0, 1, 0));
    expect(sheet.cells.get(cellKey(2, 0))?.value).toEqual(textValue('b'));
  });

  it('blocks Insert Cells that would change part of a merged cell', () => {
    const { workbook, sheet } = workbookWithSheet();
    sheet.merges.push({ range: range(0, 0, 1, 1) });
    const session = new DocumentSession(workbook);
    const result = dispatch(session, 'structure.insertCells', sheet, range(0, 0, 0, 0), {
      shift: 'right',
    });
    expect(result.structure?.state).toBe('blocked');
    expect(result.ui).toEqual({ type: 'toast', message: 'Cannot change part of a merged cell.' });
    expect(session.transactionCount).toBe(0);
  });

  it('blocks an edit that would move a shared formula group', () => {
    const { workbook, sheet } = workbookWithSheet();
    setCellSnapshot(sheet, 2, 0, {
      value: numberValue(1),
      formula: 'A3',
      formulaDerived: false,
      savedResult: numberValue(1),
      styleId: 0,
      richText: null,
      formulaSource: { kind: 'shared-master', sharedIndex: '0', range: 'A3:A4', xml: '<f>A3</f>' },
    });
    const session = new DocumentSession(workbook);
    const result = dispatch(session, 'structure.insertRows', sheet, range(0, 0, 0, 0));
    expect(result.structure?.state).toBe('blocked');
    expect(result.structure?.message).toContain('Shared and array formulas');
    expect(session.transactionCount).toBe(0);
  });

  it('rejects an unknown Insert Cells direction without a transaction', () => {
    const { workbook, sheet } = workbookWithSheet();
    const session = new DocumentSession(workbook);
    const result = dispatch(session, 'structure.insertCells', sheet, range(0, 0, 0, 0), {
      shift: 'sideways',
    });
    expect(result.structure?.state).toBe('blocked');
    expect(session.transactionCount).toBe(0);
  });
});

// ---------------------------------------------------------------------------
// F11 filter-structural
// ---------------------------------------------------------------------------

describe('F11 filter-structural through the command', () => {
  function filtered() {
    const { workbook, sheet } = workbookWithSheet();
    const labels = ['Name', 'Qty'];
    for (const [col, label] of labels.entries()) put(sheet, 0, col, textValue(label));
    for (let row = 1; row <= 5; row += 1) {
      put(sheet, row, 0, textValue(`r${row}`));
      put(sheet, row, 1, numberValue(row * 10));
    }
    sheet.autoFilter = createAutoFilter(sheet.id, range(0, 0, 5, 1), [
      [
        1,
        {
          kind: 'conditions',
          join: 'and',
          conditions: [{ kind: 'number', operator: '>', value: 20 }],
        },
      ],
    ]);
    return { workbook, sheet };
  }

  it('shifts the filter range and keeps the criteria on their columns', () => {
    const { workbook, sheet } = filtered();
    const session = new DocumentSession(workbook);
    const before = sheetState(sheet);
    dispatch(session, 'structure.insertRows', sheet, range(1, 0, 1, 0));
    expect(sheet.autoFilter?.range).toEqual(range(0, 0, 6, 1));
    expect([...(sheet.autoFilter?.columns.keys() ?? [])]).toEqual([1]);
    // Visibility recomputes against the shifted cells: the inserted blank row
    // fails the Qty > 20 criterion and is filter-hidden.
    expect([...evaluateAutoFilter(sheet, sheet.autoFilter!).filterHidden]).toContain(1);
    session.dispatch('edit.undo', {}, { activeSheetId: sheet.id });
    expect(sheetState(sheet)).toEqual(before);
  });

  it('moves the column offset with its column', () => {
    const { workbook, sheet } = filtered();
    const session = new DocumentSession(workbook);
    dispatch(session, 'structure.insertColumns', sheet, range(0, 0, 0, 0));
    expect(sheet.autoFilter?.range).toEqual(range(0, 1, 5, 2));
    expect([...(sheet.autoFilter?.columns.keys() ?? [])]).toEqual([1]);
  });
});

// ---------------------------------------------------------------------------
// F12 filter-header-delete
// ---------------------------------------------------------------------------

describe('F12 filter-header-delete through the command', () => {
  it('clears the filter and restores the filter-hidden rows in one transaction', () => {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, textValue('Name'));
    put(sheet, 0, 1, textValue('Qty'));
    for (let row = 1; row <= 4; row += 1) {
      put(sheet, row, 0, textValue(`r${row}`));
      put(sheet, row, 1, numberValue(row === 2 || row === 4 ? 5 : 15));
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
    const session = new DocumentSession(workbook);
    const before = sheetState(sheet);

    const result = dispatch(session, 'structure.deleteRows', sheet, range(0, 0, 0, 0));
    expect(result.structure?.filterCleared).toBe(true);
    expect(sheet.autoFilter).toBeNull();
    // Rows 2 and 4 became 1 and 3; their filter-hidden flag is restored.
    expect(sheet.rowDimensions.get(1)?.hidden).toBeUndefined();
    expect(sheet.rowDimensions.get(3)?.hidden).toBeUndefined();
    expect(session.transactionCount).toBe(1);

    session.dispatch('edit.undo', {}, { activeSheetId: sheet.id });
    expect(sheetState(sheet)).toEqual(before);
    expect(sheet.rowDimensions.get(2)?.hidden).toBe(true);
    expect(sheet.rowDimensions.get(4)?.hidden).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// CF17 / CF20
// ---------------------------------------------------------------------------

describe('CF17 and CF20 structural conditional formatting', () => {
  function cfWorkbook() {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, textValue('h'));
    put(sheet, 1, 0, numberValue(1));
    put(sheet, 2, 0, numberValue(2));
    sheet.conditionalFormats.push(
      {
        range: range(0, 0, 2, 0),
        rawXml:
          '<conditionalFormatting sqref="A1:A3"><cfRule type="cellIs" dxfId="1" priority="1" operator="greaterThan"><formula>A2&gt;1</formula></cfRule></conditionalFormatting>',
      },
      {
        range: range(1, 0, 1, 0),
        rawXml: '<cfRule type="top10" dxfId="2" priority="2" rank="10"/>',
      },
    );
    return { workbook, sheet };
  }

  it('transforms sqref and references, removes a collapsed rule and keeps priority order', () => {
    const { workbook, sheet } = cfWorkbook();
    const session = new DocumentSession(workbook);
    dispatch(session, 'structure.insertRows', sheet, range(1, 0, 1, 0));
    expect(sheet.conditionalFormats).toHaveLength(2);
    expect(sheet.conditionalFormats[0]?.range).toEqual(range(0, 0, 3, 0));
    expect(sheet.conditionalFormats[0]?.rawXml).toContain('sqref="A1:A4"');
    expect(sheet.conditionalFormats[0]?.rawXml).toContain('priority="1"');
    expect(sheet.conditionalFormats[0]?.rawXml).toContain('<formula>A3&gt;1</formula>');
    expect(sheet.conditionalFormats[1]?.range).toEqual(range(2, 0, 2, 0));

    // Deleting the anchor row of the second rule removes that rule only.
    dispatch(session, 'structure.deleteRows', sheet, range(2, 0, 2, 0));
    expect(sheet.conditionalFormats).toHaveLength(1);
    expect(sheet.conditionalFormats[0]?.rawXml).toContain('priority="1"');
  });

  it('undo of a structural edit restores the transformed rules exactly (CF20)', () => {
    const { workbook, sheet } = cfWorkbook();
    const session = new DocumentSession(workbook);
    const before = sheetState(sheet);
    dispatch(session, 'structure.insertRows', sheet, range(1, 0, 1, 0));
    const after = sheetState(sheet);
    session.dispatch('edit.undo', {}, { activeSheetId: sheet.id });
    expect(sheetState(sheet)).toEqual(before);
    session.dispatch('edit.redo', {}, { activeSheetId: sheet.id });
    expect(sheetState(sheet)).toEqual(after);
  });
});

// ---------------------------------------------------------------------------
// Patch identity: renumbering and untouched parts
// ---------------------------------------------------------------------------

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const STRUCTURE_WORKSHEET = `${HEADER}
<worksheet xmlns="${NS}">
  <dimension ref="A1:C3"/>
  <cols><col min="1" max="3" width="12" customWidth="1" foo="bar"/></cols>
  <sheetData>
    <row r="1" ht="20" customHeight="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c></row>
    <row r="2"><c r="A2"><v>1</v></c><c r="B2"><f>A2*2</f><v>2</v></c></row>
    <row r="3"><c r="A3"><v>2</v></c></row>
  </sheetData>
  <mergeCells count="1"><mergeCell ref="A2:B2"/></mergeCells>
  <conditionalFormatting sqref="A2:A3"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>A2&gt;0</formula></cfRule></conditionalFormatting>
  <dataValidations count="1"><dataValidation type="list" sqref="A2:A3"><formula1>$C$2:$C$3</formula1></dataValidation></dataValidations>
  <extLst><ext uri="calcplus:unknown"><cp:property xmlns:cp="urn:calcplus:test" name="preserve-me"/></ext></extLst>
</worksheet>`;

describe('structural save renumbers and preserves unmodelled content', () => {
  it('renumbers rows, cells and range-bearing XML, leaving other parts byte-identical', () => {
    const original = buildXlsx(STRUCTURE_WORKSHEET);
    const opened = openWorkbook(original);
    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    const beforeXml = unzipPart(original, 'xl/worksheets/sheet1.xml');

    dispatch(session, 'structure.insertRows', sheet, range(1, 0, 1, 0));
    const result = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
      dirtyDimensions: session.dirtyDimensionSnapshot(),
      dirtyMerges: new Set(session.dirtyMergeSnapshot().keys()),
    });

    const afterXml = unzipPart(result.bytes, 'xl/worksheets/sheet1.xml');
    expect(directChildNames(afterXml)).toEqual(directChildNames(beforeXml));
    // Unmodelled row and column attributes move with their records.
    expect(afterXml).toContain('<col min="1" max="3" width="12" customWidth="1" foo="bar"/>');
    expect(afterXml).toContain('<row r="1" ht="20" customHeight="1">');
    // The inserted row is empty, the old row 2 content is now row 3 and the
    // formula cell moved with its row.
    expect(afterXml).toContain(
      '<row r="3"><c r="A3"><v>1</v></c><c r="B3"><f>A3*2</f><v>2</v></c></row>',
    );
    expect(afterXml).toContain('<row r="4"><c r="A4"><v>2</v></c></row>');
    // Merged range, CF sqref, validation sqref and formula renumber too.
    expect(afterXml).toContain('<mergeCell ref="A3:B3"/>');
    expect(afterXml).toContain('<conditionalFormatting sqref="A3:A4">');
    expect(afterXml).toContain('<formula>A3&gt;0</formula>');
    expect(afterXml).toContain('<dataValidation type="list" sqref="A3:A4">');
    expect(afterXml).toContain('<formula1>$C$3:$C$4</formula1>');
    // Unknown extensions stay byte-for-byte.
    expect(afterXml.slice(afterXml.indexOf('<extLst>'))).toBe(
      beforeXml.slice(beforeXml.indexOf('<extLst>')),
    );

    // Only the worksheet and the workbook (fullCalcOnLoad) differ.
    for (const [path, bytes] of opened.package.entries) {
      if (path === 'xl/worksheets/sheet1.xml' || path === 'xl/workbook.xml') continue;
      expect(Buffer.from(result.package.entries.get(path)!).equals(Buffer.from(bytes))).toBe(true);
    }

    const reopened = openWorkbook(result.bytes);
    const reopenedSheet = reopened.workbook.sheets[0]!;
    expect(reopenedSheet.cells.get('2:0')?.value).toEqual({ type: 'number', value: 1 });
    expect(reopenedSheet.cells.get('2:1')?.formula).toBe('A3*2');
    expect(reopenedSheet.merges[0]?.range).toEqual(range(2, 0, 2, 1));
  });

  it('writes a pure move that has no dirty cell, formula or merge', () => {
    const original = buildXlsx(
      `${HEADER}
<worksheet xmlns="${NS}">
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>one</t></is></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>two</t></is></c></row>
  </sheetData>
</worksheet>`,
    );
    const opened = openWorkbook(original);
    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    dispatch(session, 'structure.insertRows', sheet, range(0, 0, 0, 0));
    expect(session.dirtyCellMap().get(sheet.id)?.size ?? 0).toBe(0);
    const result = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
      dirtyDimensions: session.dirtyDimensionSnapshot(),
      dirtyMerges: new Set(session.dirtyMergeSnapshot().keys()),
    });
    expect(result.noOp).toBe(false);
    const reopened = openWorkbook(result.bytes);
    const reopenedSheet = reopened.workbook.sheets[0]!;
    expect(reopenedSheet.cells.get('1:0')?.value).toEqual({ type: 'text', value: 'one' });
    expect(reopenedSheet.cells.get('2:0')?.value).toEqual({ type: 'text', value: 'two' });
    // The inline-string form survives a pure move.
    expect(unzipPart(result.bytes, 'xl/worksheets/sheet1.xml')).toContain(
      '<c r="A2" t="inlineStr"><is><t>one</t></is></c>',
    );
  });

  it('composes two structural edits before one save', () => {
    const original = buildXlsx(
      `${HEADER}
<worksheet xmlns="${NS}">
  <dimension ref="A1:B2"/>
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row>
    <row r="2"><c r="A2"><v>3</v></c><c r="B2"><v>4</v></c></row>
  </sheetData>
</worksheet>`,
    );
    const opened = openWorkbook(original);
    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    dispatch(session, 'structure.insertRows', sheet, range(0, 0, 0, 0));
    dispatch(session, 'structure.deleteColumns', sheet, range(0, 0, 0, 0));
    expect(session.transactionCount).toBe(2);
    const result = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
      dirtyDimensions: session.dirtyDimensionSnapshot(),
      dirtyMerges: new Set(session.dirtyMergeSnapshot().keys()),
    });
    const reopened = openWorkbook(result.bytes);
    const reopenedSheet = reopened.workbook.sheets[0]!;
    expect(reopenedSheet.cells.get('1:0')?.value).toEqual({ type: 'number', value: 2 });
    expect(reopenedSheet.cells.get('2:0')?.value).toEqual({ type: 'number', value: 4 });
  });

  it('splits a spanning col record on a column insert', () => {
    const original = buildXlsx(
      `${HEADER}
<worksheet xmlns="${NS}">
  <cols><col min="1" max="2" width="10" customWidth="1" foo="bar"/></cols>
  <sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row></sheetData>
</worksheet>`,
    );
    const opened = openWorkbook(original);
    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    dispatch(session, 'structure.insertColumns', sheet, range(0, 1, 0, 1));
    const result = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
      dirtyDimensions: session.dirtyDimensionSnapshot(),
      dirtyMerges: new Set(session.dirtyMergeSnapshot().keys()),
    });
    const xml = unzipPart(result.bytes, 'xl/worksheets/sheet1.xml');
    expect(xml).toContain('<col min="1" max="1" width="10" customWidth="1" foo="bar"/>');
    expect(xml).toContain('<col min="3" max="3" width="10" customWidth="1" foo="bar"/>');
    const reopened = openWorkbook(result.bytes);
    expect(reopened.workbook.sheets[0]!.cells.get('0:2')?.value).toEqual({
      type: 'number',
      value: 2,
    });
  });

  it('saves a cleared header filter and restores only the filter-hidden layer', () => {
    const original = buildXlsx(
      `${HEADER}
<worksheet xmlns="${NS}">
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Qty</t></is></c></row>
    <row r="2" hidden="1"><c r="A2" t="inlineStr"><is><t>r1</t></is></c><c r="B2"><v>15</v></c></row>
    <row r="3" hidden="1"><c r="A3" t="inlineStr"><is><t>r2</t></is></c><c r="B3"><v>5</v></c></row>
    <row r="4" hidden="1"><c r="A4" t="inlineStr"><is><t>r3</t></is></c><c r="B4"><v>15</v></c></row>
  </sheetData>
  <autoFilter ref="A1:B4"><filterColumn colId="1"><customFilters><customFilter operator="greaterThan" val="10"/></customFilters></filterColumn></autoFilter>
</worksheet>`,
    );
    const opened = openWorkbook(original);
    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    // On open the failing row is the filter layer; the other hidden rows are manual.
    expect(sheet.rowDimensions.get(1)?.hidden).toBe(true);
    expect(sheet.rowDimensions.get(2)?.hidden).toBe(true);

    dispatch(session, 'structure.deleteRows', sheet, range(0, 0, 0, 0));
    expect(sheet.autoFilter).toBeNull();
    // The filter-hidden row (now row 1) is visible; the manual rows stay hidden.
    expect(sheet.rowDimensions.get(1)?.hidden).toBeUndefined();
    expect(sheet.rowDimensions.get(0)?.hidden).toBe(true);
    expect(sheet.rowDimensions.get(2)?.hidden).toBe(true);

    const result = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
      dirtyDimensions: session.dirtyDimensionSnapshot(),
      dirtyMerges: new Set(session.dirtyMergeSnapshot().keys()),
    });
    const xml = unzipPart(result.bytes, 'xl/worksheets/sheet1.xml');
    expect(xml).not.toContain('<autoFilter');
    expect(xml).toContain('<row r="1" hidden="1">');
    expect(xml).toContain('<row r="2">');
    expect(xml).toContain('<row r="3" hidden="1">');
    const reopened = openWorkbook(result.bytes);
    expect(reopened.workbook.sheets[0]!.autoFilter).toBeNull();
    expect(reopened.workbook.sheets[0]!.rowDimensions.get(0)?.hidden).toBe(true);
    expect(reopened.workbook.sheets[0]!.rowDimensions.get(1)?.hidden).toBeUndefined();
    expect(reopened.workbook.sheets[0]!.rowDimensions.get(2)?.hidden).toBe(true);
  });

  it('keeps a structurally untouched part byte-identical', () => {
    const original = buildXlsx(STRUCTURE_WORKSHEET);
    const opened = openWorkbook(original);
    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    dispatch(session, 'structure.insertRows', sheet, range(1, 0, 1, 0));
    const result = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
      dirtyDimensions: session.dirtyDimensionSnapshot(),
      dirtyMerges: new Set(session.dirtyMergeSnapshot().keys()),
    });
    expect(result.noOp).toBe(false);
    // A second save with no further edits is a byte-identical no-op.
    const second = saveWorkbook({
      workbook: opened.workbook,
      package: result.package,
      dirtyCells: new Map(),
    });
    expect(second.noOp).toBe(true);
    expect(Buffer.from(second.bytes).equals(Buffer.from(result.bytes))).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// Existing compatibility fixture
// ---------------------------------------------------------------------------

describe('structural save on the compatibility fixture', () => {
  it('renumbers a real worksheet without disturbing its opaque parts', () => {
    const original = new Uint8Array(
      readFileSync(new URL('../fixtures/compat/m1-chart-unknown.xlsx', import.meta.url)),
    );
    const opened = openWorkbook(original);
    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    dispatch(session, 'structure.insertRows', sheet, range(0, 0, 0, 0));
    const result = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
      dirtyDimensions: session.dirtyDimensionSnapshot(),
      dirtyMerges: new Set(session.dirtyMergeSnapshot().keys()),
    });
    const beforeXml = unzipPart(original, 'xl/worksheets/sheet1.xml');
    const afterXml = unzipPart(result.bytes, 'xl/worksheets/sheet1.xml');
    expect(directChildNames(afterXml)).toEqual(directChildNames(beforeXml));
    const afterRows = elementRanges(afterXml, 'row');
    const beforeRows = elementRanges(beforeXml, 'row');
    expect(afterRows.length).toBe(beforeRows.length);
    expect(afterXml).toContain('<legacyDrawing r:id="rId2"/>');
    expect(afterXml).toContain('preserve-me');
    for (const [path, bytes] of opened.package.entries) {
      if (
        path === 'xl/worksheets/sheet1.xml' ||
        path === 'xl/workbook.xml' ||
        path === 'xl/calcChain.xml'
      ) {
        continue;
      }
      expect(Buffer.from(result.package.entries.get(path)!).equals(Buffer.from(bytes)), path).toBe(
        true,
      );
    }
  });
});
