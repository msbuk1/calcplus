import { describe, expect, it } from 'vitest';
import { createCommandRegistry } from '@commands/registry';
import { DocumentSession } from '@commands/session';
import { searchWorkbook } from '@model/find';
import { MAX_COLUMN, MAX_ROW } from '@model/references';
import { appendStyle, createEmptyWorkbook, createSheet, getStyle } from '@model/workbook';
import {
  getCell,
  setCellSnapshot,
  type Cell,
  type GridRange,
  type Sheet,
  type Workbook,
} from '@model/types';
import { numberValue, textValue } from '@model/value';
import { formatValue } from '@view/number-format';

function textCell(value: string, styleId = 0): Cell {
  return {
    value: textValue(value),
    formula: null,
    formulaDerived: false,
    savedResult: null,
    styleId,
    richText: null,
    stale: false,
  };
}

function numberCell(value: number, styleId = 0): Cell {
  return {
    value: numberValue(value),
    formula: null,
    formulaDerived: false,
    savedResult: null,
    styleId,
    richText: null,
    stale: false,
  };
}

function formulaCell(formula: string, value: number, savedResult: number | null = value): Cell {
  return {
    value: numberValue(value),
    formula,
    formulaDerived: false,
    savedResult: savedResult === null ? null : numberValue(savedResult),
    styleId: 0,
    richText: null,
    stale: false,
  };
}

function addNumberFormat(workbook: Workbook): number {
  return appendStyle(workbook, {
    numFmtId: 0,
    fontId: 0,
    fillId: 0,
    borderId: 0,
    xfId: 0,
    rawXml: '',
    numFmtCode: '#,##0.00',
    numFmtKind: 'custom',
  });
}

function workbookPair(): { workbook: Workbook; first: Sheet; second: Sheet } {
  const workbook = createEmptyWorkbook();
  const first = createSheet('Sheet1', 'xl/worksheets/sheet1.xml');
  const second = createSheet('Sheet2', 'xl/worksheets/sheet2.xml');
  workbook.sheets.push(first, second);
  return { workbook, first, second };
}

function searchFixture() {
  const { workbook, first, second } = workbookPair();
  const formatId = addNumberFormat(workbook);
  setCellSnapshot(first, 1, 0, textCell('Apple'));
  setCellSnapshot(first, 0, 1, numberCell(1234.5, formatId));
  setCellSnapshot(first, 2, 0, formulaCell('SUM(A2)', 57));
  setCellSnapshot(first, 4, 0, textCell('[draft] (a+b)'));
  setCellSnapshot(second, 0, 0, textCell('apple pie'));
  return { workbook, first, second, formatId };
}

function replaceFixture() {
  const { workbook, first, second } = workbookPair();
  setCellSnapshot(first, 0, 0, textCell('old value'));
  setCellSnapshot(first, 0, 1, numberCell(1234));
  setCellSnapshot(first, 1, 0, formulaCell('SUM(A1:A2)', 7));
  setCellSnapshot(second, 0, 0, textCell('old two'));
  return { workbook, first, second };
}

function gotoFixture() {
  const { workbook, first, second } = workbookPair();
  workbook.names.push({
    name: 'TaxRange',
    formula: "'Sheet2'!$A$1:$B$2",
    sheetId: null,
  });
  second.names.push({
    name: 'LocalAmount',
    formula: '$C$3',
    sheetId: second.id,
  });
  return { workbook, first, second };
}

describe('M4C find command ids', () => {
  it('registers the four submission commands the dialogs call', () => {
    const registry = createCommandRegistry();
    for (const id of ['find.search', 'find.replaceOne', 'find.replaceAll', 'goto.goto']) {
      expect(registry.require(id).id).toBe(id);
    }
    expect(registry.require('find.search').keybindings).toContain('Ctrl+F');
    expect(registry.require('find.replaceOne').keybindings).toContain('Ctrl+H');
    expect(registry.require('goto.goto').keybindings).toContain('Ctrl+G');
  });
});

describe('M4C find search', () => {
  it('finds display text, raw values and formulas with scope and options', () => {
    const { workbook, first, second } = searchFixture();

    const sheetScope = {
      query: 'apple',
      scope: 'sheet' as const,
      sheetId: first.id,
      inFormulas: true,
      inValues: true,
    };
    expect(searchWorkbook(workbook, sheetScope)).toEqual([
      { sheetId: first.id, row: 1, col: 0, preview: 'Apple' },
    ]);

    expect(searchWorkbook(workbook, { ...sheetScope, scope: 'workbook' })).toEqual([
      { sheetId: first.id, row: 1, col: 0, preview: 'Apple' },
      { sheetId: second.id, row: 0, col: 0, preview: 'apple pie' },
    ]);

    expect(searchWorkbook(workbook, { ...sheetScope, scope: 'workbook', matchCase: true })).toEqual(
      [{ sheetId: second.id, row: 0, col: 0, preview: 'apple pie' }],
    );
  });

  it('searches formatted display text and the raw value for the same cell', () => {
    const { workbook, first } = searchFixture();
    const displayText = (cell: Cell): string =>
      formatValue(cell.value, getStyle(workbook, cell.styleId).numFmtCode, workbook.date1904);
    const base = {
      query: '',
      scope: 'sheet' as const,
      sheetId: first.id,
      displayText,
    };

    expect(searchWorkbook(workbook, { ...base, query: '1,234' })).toEqual([
      { sheetId: first.id, row: 0, col: 1, preview: '1,234.50' },
    ]);
    expect(searchWorkbook(workbook, { ...base, query: '1234' })).toEqual([
      { sheetId: first.id, row: 0, col: 1, preview: '1234.5' },
    ]);
  });

  it('honours whole-cell matching and the formula/value switches', () => {
    const { workbook, first } = searchFixture();
    const base = { query: 'Apple', scope: 'sheet' as const, sheetId: first.id };

    expect(searchWorkbook(workbook, { ...base, wholeCell: true })).toHaveLength(1);
    expect(searchWorkbook(workbook, { ...base, query: 'App', wholeCell: true })).toHaveLength(0);

    expect(
      searchWorkbook(workbook, {
        query: 'SUM',
        scope: 'sheet',
        sheetId: first.id,
        inFormulas: true,
        inValues: false,
      }),
    ).toEqual([{ sheetId: first.id, row: 2, col: 0, preview: 'SUM(A2)' }]);

    expect(
      searchWorkbook(workbook, {
        query: '57',
        scope: 'sheet',
        sheetId: first.id,
        inFormulas: false,
        inValues: true,
      }),
    ).toEqual([{ sheetId: first.id, row: 2, col: 0, preview: '57' }]);
  });

  it('returns document order: sheet order, then row, then column', () => {
    const { workbook, first, second } = workbookPair();
    setCellSnapshot(first, 1, 0, textCell('hit'));
    setCellSnapshot(first, 0, 1, textCell('hit'));
    setCellSnapshot(second, 0, 0, textCell('hit'));

    expect(
      searchWorkbook(workbook, {
        query: 'hit',
        scope: 'workbook',
        inFormulas: true,
        inValues: true,
      }).map((match) => [match.sheetId, match.row, match.col]),
    ).toEqual([
      [first.id, 0, 1],
      [first.id, 1, 0],
      [second.id, 0, 0],
    ]);
  });

  it('treats regex punctuation as literal text without throwing', () => {
    const { workbook, first } = searchFixture();
    const matches = searchWorkbook(workbook, {
      query: '[',
      scope: 'sheet',
      sheetId: first.id,
    });
    expect(matches).toEqual([{ sheetId: first.id, row: 4, col: 0, preview: '[draft] (a+b)' }]);
    expect(searchWorkbook(workbook, { query: '', scope: 'workbook' })).toEqual([]);
  });

  it('dispatches find.search without creating a transaction', () => {
    const { workbook } = searchFixture();
    const session = new DocumentSession(workbook);
    const result = session.dispatch('find.search', {
      query: 'apple',
      scope: 'workbook',
      matchCase: false,
      wholeCell: false,
      inFormulas: true,
      inValues: true,
    });
    expect(result.matches).toHaveLength(2);
    expect(result.find?.matches).toHaveLength(2);
    expect(result.find?.total).toBe(2);
    expect(result.changes).toHaveLength(0);
    expect(session.transactionCount).toBe(0);

    const displayResult = session.dispatch('find.search', {
      query: '1,234',
      scope: 'sheet',
      matchCase: false,
      wholeCell: false,
      inFormulas: true,
      inValues: true,
    });
    expect(displayResult.find?.matches).toEqual([
      {
        sheetId: session.workbook.sheets[0]!.id,
        row: 0,
        col: 1,
        preview: '1,234.50',
      },
    ]);
  });
});

describe('M4C replace', () => {
  it('replaces one matching value and undoes it as one transaction', () => {
    const { workbook, first } = replaceFixture();
    const session = new DocumentSession(workbook);

    const result = session.dispatch(
      'find.replaceOne',
      {
        query: 'old',
        replacement: 'new',
        scope: 'sheet',
        matchCase: false,
        wholeCell: false,
        inFormulas: true,
        inValues: true,
      },
      { activeSheetId: first.id },
    );

    expect(result.replaced).toBe(1);
    expect(result.find?.replaced).toBe(1);
    expect(result.find?.total).toBe(0);
    expect(result.selectionAfter).toEqual({ startRow: 0, startCol: 0, endRow: 0, endCol: 0 });
    expect(getCell(first, 0, 0)?.value).toEqual(textValue('new value'));
    expect(session.transactionCount).toBe(1);

    session.dispatch('edit.undo', undefined);
    expect(getCell(first, 0, 0)?.value).toEqual(textValue('old value'));
    expect(session.transactionCount).toBe(0);
  });

  it('replaces every match across the workbook as one transaction', () => {
    const { workbook, first, second } = replaceFixture();
    const session = new DocumentSession(workbook);

    const result = session.dispatch('find.replaceAll', {
      query: 'old',
      replacement: 'new',
      scope: 'workbook',
      matchCase: false,
      wholeCell: false,
      inFormulas: true,
      inValues: true,
    });

    expect(result.replaced).toBe(2);
    expect(result.find?.replaced).toBe(2);
    expect(result.find?.total).toBe(0);
    expect(result.changes).toHaveLength(2);
    expect(getCell(first, 0, 0)?.value).toEqual(textValue('new value'));
    expect(getCell(second, 0, 0)?.value).toEqual(textValue('new two'));
    expect(session.transactionCount).toBe(1);

    session.dispatch('edit.undo', undefined);
    expect(getCell(first, 0, 0)?.value).toEqual(textValue('old value'));
    expect(getCell(second, 0, 0)?.value).toEqual(textValue('old two'));
  });

  it('replaces formula text, drops the cached result and marks the cell stale', () => {
    const { workbook, first } = replaceFixture();
    const session = new DocumentSession(workbook);

    const result = session.dispatch('find.replaceAll', {
      query: 'A1:A2',
      replacement: 'A1:A3',
      scope: 'workbook',
      matchCase: true,
      wholeCell: false,
      inFormulas: true,
      inValues: false,
    });

    expect(result.replaced).toBe(1);
    expect(getCell(first, 1, 0)?.formula).toBe('SUM(A1:A3)');
    expect(getCell(first, 1, 0)?.savedResult).toBeNull();
    expect(getCell(first, 1, 0)?.stale).toBe(true);
    expect(session.transactionCount).toBe(1);

    const changed = session.applyCalculationResults(
      [
        {
          sheetId: first.id,
          row: 1,
          col: 0,
          value: numberValue(9),
          status: 'value',
          stale: false,
        },
      ],
      { markDirty: true },
    );
    expect(changed).toBe(1);
    expect(getCell(first, 1, 0)?.value).toEqual(numberValue(9));
    expect(getCell(first, 1, 0)?.savedResult).toEqual(numberValue(9));
    expect(getCell(first, 1, 0)?.stale).toBe(false);

    session.dispatch('edit.undo', undefined);
    expect(getCell(first, 1, 0)?.formula).toBe('SUM(A1:A2)');
    expect(getCell(first, 1, 0)?.savedResult).toEqual(numberValue(7));
    expect(getCell(first, 1, 0)?.stale).toBe(false);
  });

  it('replaces numeric values with typed values', () => {
    const { workbook, first } = replaceFixture();
    const session = new DocumentSession(workbook);

    session.dispatch('find.replaceAll', {
      query: '1234',
      replacement: '5678',
      scope: 'sheet',
      inFormulas: false,
      inValues: true,
    });

    expect(getCell(first, 0, 1)?.value).toEqual(numberValue(5678));
  });

  it('honours whole-cell replacement', () => {
    const { workbook, first } = workbookPair();
    setCellSnapshot(first, 0, 0, textCell('cat'));
    setCellSnapshot(first, 1, 0, textCell('concatenate'));
    const session = new DocumentSession(workbook);

    const result = session.dispatch('find.replaceAll', {
      query: 'cat',
      replacement: 'dog',
      scope: 'sheet',
      wholeCell: true,
    });

    expect(result.find?.replaced).toBe(1);
    expect(getCell(first, 0, 0)?.value).toEqual(textValue('dog'));
    expect(getCell(first, 1, 0)?.value).toEqual(textValue('concatenate'));
  });

  it('inserts replacement text literally, even when it looks like a regex pattern', () => {
    const { workbook, first } = workbookPair();
    setCellSnapshot(first, 0, 0, textCell('a[b(c'));
    const session = new DocumentSession(workbook);

    expect(() =>
      session.dispatch('find.replaceAll', {
        query: '[',
        replacement: '$&',
        scope: 'sheet',
        inFormulas: true,
        inValues: true,
      }),
    ).not.toThrow();
    expect(getCell(first, 0, 0)?.value).toEqual(textValue('a$&b(c'));
  });

  it('reports no match without changing the document', () => {
    const { workbook } = replaceFixture();
    const session = new DocumentSession(workbook);

    const result = session.dispatch('find.replaceAll', {
      query: 'absent',
      replacement: 'new',
      scope: 'workbook',
    });
    expect(result.changes).toHaveLength(0);
    expect(result.replaced).toBeUndefined();
    expect(result.ui?.message).toMatch(/No matches/);
    expect(session.transactionCount).toBe(0);
  });

  it('clears a plain cell when the replacement is empty and undoes the clear', () => {
    const { workbook, first } = workbookPair();
    setCellSnapshot(first, 0, 0, textCell('old value'));
    const session = new DocumentSession(workbook);

    const result = session.dispatch('find.replaceAll', {
      query: 'old value',
      replacement: '',
      scope: 'sheet',
      wholeCell: true,
    });

    expect(result.replaced).toBe(1);
    expect(result.find?.replaced).toBe(1);
    expect(getCell(first, 0, 0)).toBeNull();
    expect(session.transactionCount).toBe(1);

    session.dispatch('edit.undo', undefined);
    expect(getCell(first, 0, 0)?.value).toEqual(textValue('old value'));
    expect(session.transactionCount).toBe(0);
  });

  it('clears a formula when the replacement is empty and undoes the clear', () => {
    const { workbook, first } = workbookPair();
    setCellSnapshot(first, 1, 0, formulaCell('SUM(A1:A2)', 7));
    const session = new DocumentSession(workbook);

    const result = session.dispatch('find.replaceAll', {
      query: 'SUM(A1:A2)',
      replacement: '',
      scope: 'sheet',
      wholeCell: true,
      inValues: false,
    });

    expect(result.replaced).toBe(1);
    expect(getCell(first, 1, 0)).toBeNull();
    expect(session.transactionCount).toBe(1);

    session.dispatch('edit.undo', undefined);
    expect(getCell(first, 1, 0)?.formula).toBe('SUM(A1:A2)');
    expect(getCell(first, 1, 0)?.savedResult).toEqual(numberValue(7));
    expect(getCell(first, 1, 0)?.stale).toBe(false);
  });
});

describe('M4C replace read-only cells', () => {
  it('refuses a legacy array follower without changing or dirtying the document', () => {
    const { workbook, first } = workbookPair();
    setCellSnapshot(first, 0, 0, {
      value: numberValue(7),
      formula: 'A2*2',
      formulaDerived: false,
      savedResult: numberValue(7),
      styleId: 0,
      richText: null,
      stale: false,
      formulaSource: {
        kind: 'array',
        range: 'A1:B2',
        xml: '<f t="array" ref="A1:B2">A2*2</f>',
      },
    });
    setCellSnapshot(first, 0, 1, textCell('hit'));
    const session = new DocumentSession(workbook);
    const version = session.version;

    const result = session.dispatch('find.replaceAll', {
      query: 'hit',
      replacement: 'replaced',
      scope: 'sheet',
    });

    expect(result.changes).toHaveLength(0);
    expect(result.replaced).toBeUndefined();
    expect(result.ui?.message).toMatch(/array formulas are read-only/);
    expect(getCell(first, 0, 1)?.value).toEqual(textValue('hit'));
    expect(session.transactionCount).toBe(0);
    expect(session.hasUnsavedChanges).toBe(false);
    expect(session.version).toBe(version);
  });

  it('refuses rich text without flattening it', () => {
    const { workbook, first } = workbookPair();
    setCellSnapshot(first, 0, 0, {
      value: textValue('rich hit'),
      formula: null,
      formulaDerived: false,
      savedResult: null,
      styleId: 0,
      richText: { xml: '<r><t>rich hit</t></r>', plainText: 'rich hit' },
      stale: false,
    });
    const session = new DocumentSession(workbook);

    const result = session.dispatch('find.replaceOne', {
      query: 'hit',
      replacement: 'replaced',
      scope: 'sheet',
    });

    expect(result.changes).toHaveLength(0);
    expect(result.replaced).toBeUndefined();
    expect(result.ui?.message).toMatch(/rich text/);
    expect(getCell(first, 0, 0)?.richText?.plainText).toBe('rich hit');
    expect(session.transactionCount).toBe(0);
    expect(session.hasUnsavedChanges).toBe(false);
  });
});

describe('M4C Go To', () => {
  const sheetOneTarget: GridRange = { startRow: 0, startCol: 0, endRow: 0, endCol: 0 };

  it('selects a cell, a range, whole columns and whole rows', () => {
    const { workbook, first, second } = gotoFixture();
    const session = new DocumentSession(workbook);

    const cell = session.dispatch(
      'goto.goto',
      { target: 'B2' },
      { activeSheetId: first.id, selection: sheetOneTarget },
    );
    expect(cell.selectionAfter).toEqual({ startRow: 1, startCol: 1, endRow: 1, endCol: 1 });
    expect(cell.activeSheetId).toBe(first.id);

    const qualified = session.dispatch(
      'goto.goto',
      { target: "'Sheet2'!A1:B2" },
      { activeSheetId: first.id, selection: sheetOneTarget },
    );
    expect(qualified.selectionAfter).toEqual({
      startRow: 0,
      startCol: 0,
      endRow: 1,
      endCol: 1,
    });
    expect(qualified.activeSheetId).toBe(second.id);

    const columns = session.dispatch('goto.goto', { target: 'A:A' }, { activeSheetId: first.id });
    expect(columns.selectionAfter).toEqual({
      startRow: 0,
      startCol: 0,
      endRow: MAX_ROW - 1,
      endCol: 0,
    });

    const rows = session.dispatch('goto.goto', { target: '1:2' }, { activeSheetId: first.id });
    expect(rows.selectionAfter).toEqual({
      startRow: 0,
      startCol: 0,
      endRow: 1,
      endCol: MAX_COLUMN - 1,
    });
  });

  it('resolves workbook- and sheet-scoped defined names', () => {
    const { workbook, first, second } = gotoFixture();
    const session = new DocumentSession(workbook);

    const workbookName = session.dispatch(
      'goto.goto',
      { target: 'TaxRange' },
      { activeSheetId: first.id },
    );
    expect(workbookName.selectionAfter).toEqual({
      startRow: 0,
      startCol: 0,
      endRow: 1,
      endCol: 1,
    });
    expect(workbookName.activeSheetId).toBe(second.id);

    const sheetName = session.dispatch(
      'goto.goto',
      { target: 'LocalAmount' },
      { activeSheetId: second.id },
    );
    expect(sheetName.selectionAfter).toEqual({
      startRow: 2,
      startCol: 2,
      endRow: 2,
      endCol: 2,
    });
    expect(sheetName.activeSheetId).toBe(second.id);
  });

  it('returns an error and leaves the selection unchanged on invalid input', () => {
    const { workbook, first } = gotoFixture();
    const session = new DocumentSession(workbook);
    const version = session.version;

    for (const target of ['NotAName', 'NoSuchSheet!A1', '', 'A1:B2:C3']) {
      const result = session.dispatch(
        'goto.goto',
        { target },
        { activeSheetId: first.id, selection: sheetOneTarget },
      );
      expect(result.changes).toHaveLength(0);
      expect(result.selectionAfter).toBeUndefined();
      expect(result.ui?.type).toBe('toast');
      expect(result.ui?.message).not.toBe('');
    }

    expect(session.version).toBe(version);
    expect(session.transactionCount).toBe(0);
  });
});
