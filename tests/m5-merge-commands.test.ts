import { describe, expect, it } from 'vitest';
import { createCommandRegistry } from '@commands/registry';
import { DocumentSession } from '@commands/session';
import {
  MERGE_ACROSS_COLUMN_MESSAGE,
  MERGE_ACROSS_SPAN_MESSAGE,
  MERGE_DATA_WARNING,
  MERGE_PARTIAL_MESSAGE,
  SHARED_FORMULA_BLOCK_MESSAGE,
} from '@model/merge';
import { ARRAY_BLOCK_MESSAGE } from '@model/sort';
import type { GridRange, Sheet } from '@model/types';
import { cloneCell, createCell, getCell, setCellSnapshot } from '@model/types';
import type { CellValue } from '@model/value';
import { numberValue, textValue } from '@model/value';
import { appendStyle, createEmptyWorkbook, createSheet } from '@model/workbook';
import { DEFAULT_ALIGNMENT } from '@model/style';
import { openWorkbook } from '@ooxml/open';
import { saveWorkbook } from '@ooxml/save';
import { buildXlsx, unzipPart } from './helpers/build-xlsx';

/**
 * Merge authoring through the command registry (feature matrix "Merge &
 * Center, Merge Across, Merge Cells, Unmerge"; command-system §2 and §4).
 *
 * Every command is one transaction. Undo restores the exact prior merges,
 * values and anchor alignment; a user-created merge then survives save and
 * reload (file-compat merged-ranges row).
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

/** Everything a merge must restore, deep enough that a lost field fails. */
function sheetState(sheet: Sheet) {
  return {
    cells: [...sheet.cells.entries()]
      .sort(([left], [right]) => left.localeCompare(right))
      .map(([address, cell]) => [address, cloneCell(cell)] as const),
    merges: sheet.merges.map((merge) => ({ ...merge.range })),
  };
}

function dispatch(
  session: DocumentSession,
  id: string,
  sheet: Sheet,
  selection: GridRange,
  args: Record<string, unknown> = {},
) {
  return session.dispatch(
    id,
    { sheetId: sheet.id, ...args },
    {
      activeSheetId: sheet.id,
      selection,
      activeCell: { row: selection.startRow, col: selection.startCol },
    },
  );
}

describe('merge.center', () => {
  it('warns before discarding data, centres the anchor, and undoes as one step', () => {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, textValue('keep'));
    put(sheet, 0, 1, textValue('lost'));
    put(sheet, 1, 0, numberValue(2));
    put(sheet, 1, 1, numberValue(3));
    const session = new DocumentSession(workbook);
    const selection = range(0, 0, 1, 1);
    const before = sheetState(sheet);

    // Excel asks before merging cells that hold data outside the anchor.
    const warned = dispatch(session, 'merge.center', sheet, selection);
    expect(warned.changes).toHaveLength(0);
    expect(warned.ui).toEqual({ type: 'dialog', message: MERGE_DATA_WARNING });
    expect(session.transactionCount).toBe(0);
    expect(sheet.merges).toHaveLength(0);
    expect(getCell(sheet, 0, 1)?.value).toEqual(textValue('lost'));

    const result = dispatch(session, 'merge.center', sheet, selection, { confirm: true });
    expect(result.ui).toBeUndefined();
    expect(session.transactionCount).toBe(1);
    expect(sheet.merges).toEqual([{ range: selection }]);
    expect(getCell(sheet, 0, 0)?.value).toEqual(textValue('keep'));
    expect(getCell(sheet, 0, 1)).toBeNull();
    expect(getCell(sheet, 1, 0)).toBeNull();
    expect(getCell(sheet, 1, 1)).toBeNull();
    const anchor = getCell(sheet, 0, 0)!;
    expect(workbook.styles[anchor.styleId]?.alignment.horizontal).toBe('center');
    const after = sheetState(sheet);

    session.dispatch('edit.undo', {}, { activeSheetId: sheet.id });
    expect(sheetState(sheet)).toEqual(before);
    expect(workbook.styles[getCell(sheet, 0, 0)?.styleId ?? 0]?.alignment.horizontal).toBe(
      'general',
    );

    session.dispatch('edit.redo', {}, { activeSheetId: sheet.id });
    expect(sheetState(sheet)).toEqual(after);
  });
});

describe('merge.cells', () => {
  it('keeps the anchor alignment and discards covered content after confirmation', () => {
    const { workbook, sheet } = workbookWithSheet();
    const left = appendStyle(workbook, {
      numFmtId: 0,
      fontId: 0,
      fillId: 0,
      borderId: 0,
      xfId: 0,
      rawXml: '',
      alignment: { ...DEFAULT_ALIGNMENT, horizontal: 'left' },
    });
    put(sheet, 0, 0, textValue('keep'), left);
    put(sheet, 0, 1, textValue('lost'));
    const session = new DocumentSession(workbook);
    const selection = range(0, 0, 0, 1);
    const before = sheetState(sheet);

    const warned = dispatch(session, 'merge.cells', sheet, selection);
    expect(warned.ui).toEqual({ type: 'dialog', message: MERGE_DATA_WARNING });
    expect(session.transactionCount).toBe(0);

    dispatch(session, 'merge.cells', sheet, selection, { confirm: true });
    expect(session.transactionCount).toBe(1);
    expect(sheet.merges).toEqual([{ range: selection }]);
    expect(getCell(sheet, 0, 0)?.value).toEqual(textValue('keep'));
    expect(getCell(sheet, 0, 1)).toBeNull();
    // Merge Cells does not centre: the anchor keeps its authored alignment.
    expect(getCell(sheet, 0, 0)?.styleId).toBe(left);
    expect(workbook.styles[left]?.alignment.horizontal).toBe('left');

    session.dispatch('edit.undo', {}, { activeSheetId: sheet.id });
    expect(sheetState(sheet)).toEqual(before);
  });
});

describe('merge.across', () => {
  it('merges one cell per row and keeps each row’s leftmost content', () => {
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 0, 0, textValue('top'));
    put(sheet, 0, 1, textValue('lost-1'));
    put(sheet, 0, 2, textValue('lost-2'));
    put(sheet, 1, 0, textValue('bottom'));
    put(sheet, 1, 1, numberValue(1));
    put(sheet, 1, 2, numberValue(2));
    const session = new DocumentSession(workbook);
    const selection = range(0, 0, 1, 2);
    const before = sheetState(sheet);

    dispatch(session, 'merge.across', sheet, selection, { confirm: true });
    expect(session.transactionCount).toBe(1);
    expect(sheet.merges).toEqual([{ range: range(0, 0, 0, 2) }, { range: range(1, 0, 1, 2) }]);
    expect(getCell(sheet, 0, 0)?.value).toEqual(textValue('top'));
    expect(getCell(sheet, 1, 0)?.value).toEqual(textValue('bottom'));
    expect(getCell(sheet, 0, 1)).toBeNull();
    expect(getCell(sheet, 0, 2)).toBeNull();
    expect(getCell(sheet, 1, 1)).toBeNull();
    expect(getCell(sheet, 1, 2)).toBeNull();

    session.dispatch('edit.undo', {}, { activeSheetId: sheet.id });
    expect(sheetState(sheet)).toEqual(before);
  });
});

describe('merge.unmerge', () => {
  it('splits every intersecting merge and leaves content in the anchor cells', () => {
    const { workbook, sheet } = workbookWithSheet();
    sheet.merges.push(
      { range: range(0, 0, 0, 2) },
      { range: range(0, 3, 1, 4) },
      { range: range(0, 6, 0, 7) },
    );
    put(sheet, 0, 0, textValue('left'));
    put(sheet, 0, 3, textValue('right'));
    put(sheet, 0, 6, textValue('untouched'));
    const session = new DocumentSession(workbook);
    const before = sheetState(sheet);

    // B1:E1 overlaps the first merge at an edge and fully covers the second.
    dispatch(session, 'merge.unmerge', sheet, range(0, 1, 0, 4));
    expect(session.transactionCount).toBe(1);
    expect(sheet.merges).toEqual([{ range: range(0, 6, 0, 7) }]);
    expect(getCell(sheet, 0, 0)?.value).toEqual(textValue('left'));
    expect(getCell(sheet, 0, 3)?.value).toEqual(textValue('right'));

    session.dispatch('edit.undo', {}, { activeSheetId: sheet.id });
    expect(sheetState(sheet)).toEqual(before);
  });

  it('is enabled only when the selection contains a merged cell', () => {
    const { workbook, sheet } = workbookWithSheet();
    sheet.merges.push({ range: range(0, 0, 0, 1) });
    const registry = createCommandRegistry();
    const unmerge = registry.require('merge.unmerge');
    const context = (selection: GridRange | null) => ({
      workbook,
      activeSheetId: sheet.id,
      selection,
      activeCell: { row: 0, col: 0 },
    });

    expect(unmerge.isEnabled(context(range(2, 2, 2, 2)))).toBe(false);
    expect(unmerge.isEnabled(context(null))).toBe(false);
    // A single cell inside the merged range is enough, as in Excel.
    expect(unmerge.isEnabled(context(range(0, 1, 0, 1)))).toBe(true);
  });
});

describe('merge enablement and blocking messages', () => {
  it('enables the three merge commands only for a multi-cell selection', () => {
    const { workbook, sheet } = workbookWithSheet();
    const registry = createCommandRegistry();
    const context = (selection: GridRange | null) => ({
      workbook,
      activeSheetId: sheet.id,
      selection,
      activeCell: { row: 0, col: 0 },
    });

    for (const id of ['merge.center', 'merge.across', 'merge.cells']) {
      const command = registry.require(id);
      expect(command.isEnabled(context(null))).toBe(false);
      expect(command.isEnabled(context(range(0, 0, 0, 0)))).toBe(false);
      expect(command.isEnabled(context(range(0, 0, 0, 1)))).toBe(true);
    }
  });

  it('blocks a selection that would split an existing merged cell', () => {
    const { workbook, sheet } = workbookWithSheet();
    sheet.merges.push({ range: range(0, 1, 0, 2) });
    const session = new DocumentSession(workbook);

    const result = dispatch(session, 'merge.center', sheet, range(0, 0, 0, 1));
    expect(result.changes).toHaveLength(0);
    expect(result.ui).toEqual({ type: 'toast', message: MERGE_PARTIAL_MESSAGE });
    expect(session.transactionCount).toBe(0);
    expect(sheet.merges).toEqual([{ range: range(0, 1, 0, 2) }]);
  });

  it('blocks a merge over a legacy array formula with the sort message', () => {
    const { workbook, sheet } = workbookWithSheet();
    setCellSnapshot(sheet, 0, 0, {
      value: numberValue(0),
      formula: 'SUM(A1:B1)',
      formulaDerived: false,
      savedResult: numberValue(0),
      styleId: 0,
      richText: null,
      formulaSource: {
        kind: 'array',
        range: 'A1:B1',
        xml: '<f t="array" ref="A1:B1">SUM(A1:B1)</f>',
      },
    });
    const session = new DocumentSession(workbook);

    const result = dispatch(session, 'merge.center', sheet, range(0, 0, 0, 1));
    expect(result.changes).toHaveLength(0);
    expect(result.ui).toEqual({ type: 'toast', message: ARRAY_BLOCK_MESSAGE });
    expect(session.transactionCount).toBe(0);
  });

  it('blocks a merge over a shared formula group before any cell is cleared', () => {
    const { workbook, sheet } = workbookWithSheet();
    setCellSnapshot(sheet, 0, 0, {
      value: numberValue(2),
      formula: 'B1*2',
      formulaDerived: false,
      savedResult: numberValue(2),
      styleId: 0,
      richText: null,
      formulaSource: {
        kind: 'shared-master',
        sharedIndex: '0',
        range: 'A1:B1',
        xml: '<f t="shared" ref="A1:B1" si="0">B1*2</f>',
      },
    });
    setCellSnapshot(sheet, 0, 1, {
      value: numberValue(1),
      formula: 'C1*2',
      formulaDerived: true,
      savedResult: numberValue(1),
      styleId: 0,
      richText: null,
      formulaSource: {
        kind: 'shared-follower',
        sharedIndex: '0',
        xml: '<f t="shared" si="0"/>',
      },
    });
    const session = new DocumentSession(workbook);

    const result = dispatch(session, 'merge.center', sheet, range(0, 0, 0, 1));
    expect(result.changes).toHaveLength(0);
    expect(result.ui).toEqual({ type: 'toast', message: SHARED_FORMULA_BLOCK_MESSAGE });
    expect(session.transactionCount).toBe(0);
    expect(sheet.cells.get('0:1')?.formulaSource?.kind).toBe('shared-follower');
  });

  it('blocks Merge Across on a multi-row merge and on a one-column selection', () => {
    const { workbook, sheet } = workbookWithSheet();
    sheet.merges.push({ range: range(0, 0, 1, 1) });
    const session = new DocumentSession(workbook);

    const span = dispatch(session, 'merge.across', sheet, range(0, 0, 1, 1));
    expect(span.ui).toEqual({ type: 'toast', message: MERGE_ACROSS_SPAN_MESSAGE });
    expect(session.transactionCount).toBe(0);

    const column = dispatch(session, 'merge.across', sheet, range(0, 4, 1, 4));
    expect(column.ui).toEqual({ type: 'toast', message: MERGE_ACROSS_COLUMN_MESSAGE });
    expect(session.transactionCount).toBe(0);
  });
});

describe('a user-created merge round-trips through .xlsx', () => {
  const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="${NS}">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
</styleSheet>`;
  const WORKSHEET = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="${NS}">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>keep</t></is></c>
      <c r="B1" t="inlineStr"><is><t>lost</t></is></c>
    </row>
  </sheetData>
</worksheet>`;

  it('writes the merge, its anchor value and its centred alignment', () => {
    const opened = openWorkbook(buildXlsx(WORKSHEET, { stylesXml: STYLES_XML }));
    const sheet = opened.workbook.sheets[0]!;
    const session = new DocumentSession(opened.workbook);

    session.dispatch(
      'merge.center',
      { sheetId: sheet.id, confirm: true },
      {
        activeSheetId: sheet.id,
        selection: range(0, 0, 0, 1),
        activeCell: { row: 0, col: 0 },
      },
    );
    const saved = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
      dirtyDimensions: session.dirtyDimensionSnapshot(),
      dirtyMerges: new Set(session.dirtyMergeSnapshot().keys()),
    });
    expect(saved.noOp).toBe(false);
    expect(unzipPart(saved.bytes, 'xl/worksheets/sheet1.xml')).toContain(
      '<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>',
    );

    const reopened = openWorkbook(saved.bytes);
    const reloaded = reopened.workbook.sheets[0]!;
    expect(reloaded.merges).toEqual([{ range: range(0, 0, 0, 1) }]);
    expect(getCell(reloaded, 0, 0)?.value).toEqual(textValue('keep'));
    expect(getCell(reloaded, 0, 1)).toBeNull();
    const anchor = getCell(reloaded, 0, 0)!;
    expect(reopened.workbook.styles[anchor.styleId]?.alignment.horizontal).toBe('center');
  });
});
