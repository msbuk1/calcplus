import { describe, expect, it } from 'vitest';
import { DocumentSession } from '@commands/session';
import { createCommandRegistry } from '@commands/registry';
import { createSheet } from '@model/workbook';
import { encodeXml } from '@ooxml/xml';
import { planPaneSave } from '@ooxml/pane-xml';
import { openWorkbook } from '@ooxml/open';
import { saveWorkbook } from '@ooxml/save';
import { buildXlsx, unzipPart } from './helpers/build-xlsx';

/**
 * Freeze panes authoring (feature matrix §6): the command matches Excel 365's
 * "rows above and columns left of the active cell" rule, a repeat press at the
 * split point unfreezes, an existing split moves with a new active cell, and
 * one press is one undo transaction. The split round-trips in the sheet view.
 */

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

function worksheet(pane = ''): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="${NS}">
  <sheetViews><sheetView workbookViewId="0">${pane}</sheetView></sheetViews>
  <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
</worksheet>`;
}

function open(worksheetXml = worksheet()) {
  return openWorkbook(buildXlsx(worksheetXml));
}

function freezeSheet() {
  const opened = open();
  const session = new DocumentSession(opened.workbook);
  const sheet = opened.workbook.sheets[0]!;
  return { opened, session, sheet };
}

describe('view.freeze', () => {
  it('freezes the rows above and the columns left of the active cell', () => {
    const { session, sheet } = freezeSheet();
    const result = session.dispatch(
      'view.freeze',
      {},
      { activeSheetId: sheet.id, activeCell: { row: 2, col: 1 } },
    );

    expect(result.changes).toHaveLength(1);
    expect(result.changes[0]?.type).toBe('sheet.freeze');
    expect(sheet.view.freezeRows).toBe(2);
    expect(sheet.view.freezeCols).toBe(1);
    expect(result.ui?.message).toBe('Froze 2 rows and 1 column.');
  });

  it('freezes only rows or only columns when one edge is the first cell', () => {
    const rowOnly = freezeSheet();
    rowOnly.session.dispatch(
      'view.freeze',
      {},
      { activeSheetId: rowOnly.sheet.id, activeCell: { row: 2, col: 0 } },
    );
    expect(rowOnly.sheet.view.freezeRows).toBe(2);
    expect(rowOnly.sheet.view.freezeCols).toBe(0);

    const columnOnly = freezeSheet();
    columnOnly.session.dispatch(
      'view.freeze',
      {},
      { activeSheetId: columnOnly.sheet.id, activeCell: { row: 0, col: 1 } },
    );
    expect(columnOnly.sheet.view.freezeRows).toBe(0);
    expect(columnOnly.sheet.view.freezeCols).toBe(1);
  });

  it('unfreezes on a second press at the same split point', () => {
    const { session, sheet } = freezeSheet();
    const options = { activeSheetId: sheet.id, activeCell: { row: 2, col: 1 } };
    session.dispatch('view.freeze', {}, options);
    const result = session.dispatch('view.freeze', {}, options);
    expect(result.ui?.message).toBe('Unfroze panes.');
    expect(sheet.view.freezeRows).toBe(0);
    expect(sheet.view.freezeCols).toBe(0);
  });

  it('moves an existing split to a new active cell', () => {
    const { session, sheet } = freezeSheet();
    session.dispatch(
      'view.freeze',
      {},
      { activeSheetId: sheet.id, activeCell: { row: 1, col: 1 } },
    );
    session.dispatch(
      'view.freeze',
      {},
      { activeSheetId: sheet.id, activeCell: { row: 3, col: 2 } },
    );
    expect(sheet.view.freezeRows).toBe(3);
    expect(sheet.view.freezeCols).toBe(2);
  });

  it('undo and redo restore the exact split as one transaction', () => {
    const { session, sheet } = freezeSheet();
    session.dispatch(
      'view.freeze',
      {},
      { activeSheetId: sheet.id, activeCell: { row: 2, col: 1 } },
    );

    session.dispatch('edit.undo', undefined);
    expect(sheet.view.freezeRows).toBe(0);
    expect(sheet.view.freezeCols).toBe(0);

    session.dispatch('edit.redo', undefined);
    expect(sheet.view.freezeRows).toBe(2);
    expect(sheet.view.freezeCols).toBe(1);
  });

  it('is a no-op at A1 when nothing is frozen', () => {
    const { session, sheet } = freezeSheet();
    const result = session.dispatch(
      'view.freeze',
      {},
      { activeSheetId: sheet.id, activeCell: { row: 0, col: 0 } },
    );
    expect(result.changes).toEqual([]);
    expect(result.ui?.message).toContain('below the first row');
  });

  it('registers the command so the ribbon control resolves', () => {
    expect(createCommandRegistry().require('view.freeze').id).toBe('view.freeze');
  });
});

describe('freeze panes round-trip', () => {
  it('writes the frozen pane in the sheet view and reads it back', () => {
    const { opened, session, sheet } = freezeSheet();
    session.dispatch(
      'view.freeze',
      {},
      { activeSheetId: sheet.id, activeCell: { row: 2, col: 1 } },
    );

    const saved = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
    });
    const part = unzipPart(saved.bytes, 'xl/worksheets/sheet1.xml');
    expect(part).toContain(
      '<pane xSplit="1" ySplit="2" topLeftCell="B3" activePane="bottomRight" state="frozen"/>',
    );

    const reopened = openWorkbook(saved.bytes);
    expect(reopened.workbook.sheets[0]!.view.freezeRows).toBe(2);
    expect(reopened.workbook.sheets[0]!.view.freezeCols).toBe(1);
  });

  it('removes the pane when the split is cleared', () => {
    const { opened, session, sheet } = freezeSheet();
    session.dispatch(
      'view.freeze',
      {},
      { activeSheetId: sheet.id, activeCell: { row: 2, col: 1 } },
    );
    session.dispatch(
      'view.freeze',
      {},
      { activeSheetId: sheet.id, activeCell: { row: 2, col: 1 } },
    );
    const saved = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
    });
    const part = unzipPart(saved.bytes, 'xl/worksheets/sheet1.xml');
    expect(part).not.toContain('<pane');
    expect(openWorkbook(saved.bytes).workbook.sheets[0]!.view.freezeRows).toBe(0);
  });

  it('leaves an untouched freeze byte-identical', () => {
    const original = buildXlsx(
      worksheet(
        '<pane xSplit="1" ySplit="2" topLeftCell="B3" activePane="bottomRight" state="frozen"/>',
      ),
    );
    const opened = openWorkbook(original);
    const saved = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: new Map(),
    });
    expect(saved.noOp).toBe(true);
    expect(Buffer.compare(Buffer.from(saved.bytes), Buffer.from(original))).toBe(0);
  });
});

describe('the pane writer', () => {
  it('inserts a sheetViews block when the part has none', () => {
    const sheet = createSheet('Sheet1', 'xl/worksheets/sheet1.xml');
    sheet.view.freezeRows = 1;
    sheet.view.freezeCols = 1;
    const source = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="${NS}">
  <dimension ref="A1"/>
  <sheetData/>
</worksheet>`;
    const plan = planPaneSave(sheet, encodeXml(source));
    const written = new TextDecoder().decode(plan.bytes);
    expect(written).toContain('<sheetViews><sheetView workbookViewId="0">');
    expect(written).toContain('topLeftCell="B2" activePane="bottomRight" state="frozen"');
    expect(written.indexOf('<sheetViews>')).toBeLessThan(written.indexOf('<sheetData'));
  });

  it('keeps a namespace prefix and replaces the existing pane', () => {
    const sheet = createSheet('Sheet1', 'xl/worksheets/sheet1.xml');
    sheet.view.freezeRows = 1;
    sheet.view.freezeCols = 0;
    const source = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:worksheet xmlns:x="${NS}">
  <x:sheetViews><x:sheetView workbookViewId="0"><x:pane ySplit="2" topLeftCell="A3" activePane="bottomLeft" state="frozen"/></x:sheetView></x:sheetViews>
  <x:sheetData/>
</x:worksheet>`;
    const plan = planPaneSave(sheet, encodeXml(source));
    const written = new TextDecoder().decode(plan.bytes);
    expect(written).toContain(
      '<x:pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>',
    );
    expect(written).not.toContain('topLeftCell="A3"');
    expect(written).not.toContain('xSplit');
  });
});
