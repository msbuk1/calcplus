import { Buffer } from 'node:buffer';
import { describe, expect, it } from 'vitest';
import { DocumentSession } from '@commands/session';
import { createCommandRegistry } from '@commands/registry';
import { changesRequireRebuild } from '@engine/calculation-controller';
import { createSheet, nextSheetName, sheetHasContent, validateSheetName } from '@model/workbook';
import {
  planSheetRename,
  rewriteReferencesToRefError,
  rewriteSheetNameInFormula,
} from '@model/sheet-references';
import { getCell } from '@model/types';
import { openWorkbook } from '@ooxml/open';
import { saveWorkbook } from '@ooxml/save';
import { restoreWorkbook, snapshotWorkbook } from '@ooxml/snapshot';
import {
  buildMultiSheetXlsx,
  unzipPart,
  unzipParts,
  type MultiSheetSpec,
} from './helpers/build-xlsx';

const SPREADSHEET_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

function worksheet(body: string, extras = ''): string {
  return `${XML_HEADER}
<worksheet xmlns="${SPREADSHEET_NS}">
  ${body}
  ${extras}
</worksheet>`;
}

const RELATIONSHIPS_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

function attachedWorksheet(body: string): string {
  return `${XML_HEADER}
<worksheet xmlns="${SPREADSHEET_NS}" xmlns:r="${RELATIONSHIPS_NS}">
  ${body}
</worksheet>`;
}

function sheet(name: string, cells: string, options: Partial<MultiSheetSpec> = {}): MultiSheetSpec {
  return {
    name,
    worksheetXml: worksheet(cells),
    ...options,
  };
}

function byteEqual(left: Uint8Array, right: Uint8Array): boolean {
  return Buffer.compare(Buffer.from(left), Buffer.from(right)) === 0;
}

function baseWorkbook(): Uint8Array {
  return buildMultiSheetXlsx(
    [
      sheet('Sheet1', '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>'),
      sheet(
        'Sheet2',
        '<sheetData><row r="1"><c r="A1"><f>Sheet1!A1</f><v>1</v></c></row></sheetData>',
      ),
    ],
    {
      workbookExtras:
        '<definedNames><definedName name="DataRange">Sheet1!$A$1</definedName></definedNames>',
    },
  );
}

describe('sheet model helpers', () => {
  it('picks the next Excel default name and validates Excel name rules', () => {
    const opened = openWorkbook(baseWorkbook());
    expect(nextSheetName(opened.workbook)).toBe('Sheet3');
    expect(validateSheetName(opened.workbook, 'Sheet3')).toBeNull();
    expect(validateSheetName(opened.workbook, 'Sheet2')).toMatch(/already exists/);
    expect(validateSheetName(opened.workbook, 'Bad:Name')).toMatch(/cannot contain/);
    expect(validateSheetName(opened.workbook, '')).toMatch(/cannot be blank/);
    expect(validateSheetName(opened.workbook, 'x'.repeat(32))).toMatch(/31 characters/);
    expect(validateSheetName(opened.workbook, 'Sheet2', opened.workbook.sheets[1]!.id)).toBeNull();
  });

  it('treats cells and formatting as content, and a blank sheet as empty', () => {
    const opened = openWorkbook(baseWorkbook());
    expect(sheetHasContent(opened.workbook.sheets[0]!)).toBe(true);
    const empty = createSheet('Blank', 'xl/worksheets/sheet9.xml');
    expect(sheetHasContent(empty)).toBe(false);
  });

  it('treats a drawing-only sheet as content', () => {
    const opened = openWorkbook(
      buildMultiSheetXlsx([{ name: 'Sheet1', worksheetXml: attachedWorksheet('<sheetData/>') }], {
        extraParts: {
          'xl/drawings/drawing1.xml':
            '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>',
          'xl/worksheets/_rels/sheet1.xml.rels': `${XML_HEADER}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${RELATIONSHIPS_NS}/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`,
        },
      }),
    );
    expect(opened.workbook.sheets[0]!.cells.size).toBe(0);
    expect(sheetHasContent(opened.workbook.sheets[0]!)).toBe(true);
  });
});

describe('sheet insert', () => {
  it('inserts a blank sheet in tab order and undoes and redoes with the same id', () => {
    const opened = openWorkbook(baseWorkbook());
    const session = new DocumentSession(opened.workbook);
    const initialVersion = session.version;
    const result = session.dispatch('sheet.insert', {});
    expect(result.changes).toHaveLength(1);
    expect(result.activeSheetId).toBe(session.workbook.sheets[2]!.id);
    const insertedId = session.workbook.sheets[2]!.id;
    expect(session.workbook.sheets[2]!.name).toBe('Sheet3');
    expect(session.workbook.sheets[2]!.view).toEqual({
      freezeRows: 0,
      freezeCols: 0,
      showGridLines: true,
      showHeadings: true,
      zoom: 100,
    });
    expect(session.version).toBeGreaterThan(initialVersion);
    expect(session.hasUnsavedChanges).toBe(true);

    const undo = session.dispatch('edit.undo', undefined);
    expect(session.workbook.sheets).toHaveLength(2);
    expect(undo.activeSheetId).toBe(session.workbook.sheets[1]!.id);

    const redo = session.dispatch('edit.redo', undefined);
    expect(session.workbook.sheets[2]!.id).toBe(insertedId);
    expect(redo.activeSheetId).toBe(insertedId);
  });

  it('inserts at an explicit index and rejects a duplicate name', () => {
    const opened = openWorkbook(baseWorkbook());
    const session = new DocumentSession(opened.workbook);
    session.dispatch('sheet.insert', { name: 'Middle', index: 1 });
    expect(session.workbook.sheets.map((entry) => entry.name)).toEqual([
      'Sheet1',
      'Middle',
      'Sheet2',
    ]);

    const duplicate = session.dispatch('sheet.insert', { name: 'Sheet1' });
    expect(duplicate.changes).toHaveLength(0);
    expect(duplicate.ui?.message).toMatch(/already exists/);
    expect(session.transactionCount).toBe(1);
  });
});

describe('sheet delete', () => {
  it('warns first, deletes terminally and clears the undo and redo history', () => {
    const opened = openWorkbook(baseWorkbook());
    const session = new DocumentSession(opened.workbook);
    const sheetId = session.workbook.sheets[0]!.id;
    const sheet2 = session.workbook.sheets[1]!;
    session.dispatch('cell.set', {
      sheetId: sheet2.id,
      row: 1,
      col: 0,
      value: { type: 'text', value: 'prior' },
    });
    // Leave both a redo entry and an undo entry in place so the delete has a
    // non-empty history to clear.
    session.dispatch('sheet.insert', {});
    session.dispatch('edit.undo', undefined);
    expect(session.transactionCount).toBe(1);
    expect(session.canRedo).toBe(true);

    const prompt = session.dispatch('sheet.delete', { sheetId });
    expect(prompt.changes).toHaveLength(0);
    expect(prompt.ui?.type).toBe('dialog');
    expect(prompt.ui?.message).toMatch(/cannot be undone/);
    expect(prompt.ui?.message).toMatch(/may be lost/);
    expect(session.transactionCount).toBe(1);

    const result = session.dispatch('sheet.delete', { sheetId, confirm: true });
    expect(result.changes[0]!.type).toBe('sheet.delete');
    expect(session.workbook.sheets.map((entry) => entry.name)).toEqual(['Sheet2']);
    expect(result.activeSheetId).toBe(session.workbook.sheets[0]!.id);
    // Excel clears the whole history when a sheet is deleted; the earlier cell
    // edit is no longer undoable because the deleted sheet is gone for good.
    expect(session.transactionCount).toBe(0);
    expect(session.canUndo).toBe(false);
    expect(session.canRedo).toBe(false);

    const undo = session.dispatch('edit.undo', undefined);
    expect(undo.changes).toHaveLength(0);
    expect(session.workbook.sheets.map((entry) => entry.name)).toEqual(['Sheet2']);
    expect(getCell(session.workbook.sheets[0]!, 1, 0)?.value).toEqual({
      type: 'text',
      value: 'prior',
    });
  });

  it('refuses to delete the only visible sheet in the workbook', () => {
    const opened = openWorkbook(baseWorkbook());
    const session = new DocumentSession(opened.workbook);
    session.dispatch('sheet.delete', {
      sheetId: session.workbook.sheets[1]!.id,
      confirm: true,
    });
    const result = session.dispatch('sheet.delete', {
      sheetId: session.workbook.sheets[0]!.id,
      confirm: true,
    });
    expect(result.changes).toHaveLength(0);
    expect(result.ui?.type).toBe('toast');
    expect(session.workbook.sheets).toHaveLength(1);
  });
});

describe('sheet rename', () => {
  it('rewrites sheet references in formulas, preserves literals, and undo restores exact text', () => {
    const opened = openWorkbook(baseWorkbook());
    const session = new DocumentSession(opened.workbook);
    const sheet1 = session.workbook.sheets[0]!.id;
    const sheet2 = session.workbook.sheets[1]!;
    const beforeFormula = getCell(sheet2, 0, 0)!.formula;

    const result = session.dispatch('sheet.rename', { sheetId: sheet1, name: 'Data' });
    expect(result.changes[0]!.type).toBe('sheet.rename');
    expect(session.workbook.sheets[0]!.name).toBe('Data');
    expect(getCell(sheet2, 0, 0)!.formula).toBe('Data!A1');
    expect(session.workbook.names[0]!.formula).toBe('Data!$A$1');

    session.dispatch('edit.undo', undefined);
    expect(session.workbook.sheets[0]!.name).toBe('Sheet1');
    expect(getCell(sheet2, 0, 0)!.formula).toBe(beforeFormula);
    expect(session.workbook.names[0]!.formula).toBe('Sheet1!$A$1');
  });

  it('formats quoted and unquoted prefixes for the new name', () => {
    expect(rewriteSheetNameInFormula('Sheet1!A1', 'Sheet1', 'Data')).toBe('Data!A1');
    expect(rewriteSheetNameInFormula("'Sheet1'!A1", 'Sheet1', 'Data')).toBe("'Data'!A1");
    expect(rewriteSheetNameInFormula('Sheet1!A1', 'Sheet1', 'My Data')).toBe("'My Data'!A1");
    expect(rewriteSheetNameInFormula('"Sheet1!A1"', 'Sheet1', 'Data')).toBe('"Sheet1!A1"');
    expect(rewriteSheetNameInFormula('[Book1]Sheet1!A1', 'Sheet1', 'Data')).toBe(
      '[Book1]Sheet1!A1',
    );
  });

  it('blocks a rename that would rewrite a read-only shared formula group', () => {
    const bytes = buildMultiSheetXlsx([
      sheet('Sheet1', '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>'),
      sheet(
        'Sheet2',
        '<sheetData><row r="1"><c r="A1"><f t="shared" si="0" ref="A1">Sheet1!A1</f><v>1</v></c></row></sheetData>',
      ),
    ]);
    const session = new DocumentSession(openWorkbook(bytes).workbook);
    const result = session.dispatch('sheet.rename', {
      sheetId: session.workbook.sheets[0]!.id,
      name: 'Data',
    });
    expect(result.changes).toHaveLength(0);
    expect(result.ui?.message).toMatch(/read-only formula group/);
    expect(session.workbook.sheets[0]!.name).toBe('Sheet1');
    expect(session.transactionCount).toBe(0);
  });

  it('blocks a rename that would leave a sheet-scoped name reference stale', () => {
    const bytes = buildMultiSheetXlsx(
      [
        sheet('Sheet1', '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>'),
        sheet(
          'Sheet2',
          '<sheetData><row r="1"><c r="A1"><f>Sheet1!LocalName</f><v>1</v></c></row></sheetData>',
        ),
      ],
      {
        workbookExtras:
          '<definedNames><definedName name="LocalName" localSheetId="0">Sheet1!$A$1</definedName></definedNames>',
      },
    );
    const session = new DocumentSession(openWorkbook(bytes).workbook);
    const sheetId = session.workbook.sheets[0]!.id;
    const rename = session.dispatch('sheet.rename', { sheetId, name: 'Data' });
    expect(rename.changes).toHaveLength(0);
    expect(rename.ui?.message).toMatch(/sheet-scoped defined name reference/);
    const remove = session.dispatch('sheet.delete', { sheetId, confirm: true });
    expect(remove.changes).toHaveLength(0);
    expect(remove.ui?.message).toMatch(/sheet-scoped defined name reference/);
    expect(session.transactionCount).toBe(0);
  });

  it('rejects duplicate and invalid names without creating a transaction', () => {
    const session = new DocumentSession(openWorkbook(baseWorkbook()).workbook);
    const sheetId = session.workbook.sheets[0]!.id;
    for (const name of ['Sheet2', 'bad/name', '']) {
      const result = session.dispatch('sheet.rename', { sheetId, name });
      expect(result.changes).toHaveLength(0);
      expect(result.ui?.type).toBe('toast');
    }
    expect(session.transactionCount).toBe(0);
  });
});

describe('delete references', () => {
  it('turns references to a deleted sheet into #REF! permanently', () => {
    const bytes = buildMultiSheetXlsx([
      sheet('Sheet1', '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>'),
      sheet(
        'Sheet2',
        '<sheetData><row r="1"><c r="A1"><f>Sheet1!A1</f><v>1</v></c></row><row r="2"><c r="A2"><f>SUM(Sheet1!A1:A2)</f><v>1</v></c></row></sheetData>',
      ),
    ]);
    const session = new DocumentSession(openWorkbook(bytes).workbook);
    const sheet1 = session.workbook.sheets[0]!.id;
    const sheet2 = session.workbook.sheets[1]!;
    session.dispatch('sheet.delete', { sheetId: sheet1, confirm: true });
    expect(getCell(sheet2, 0, 0)!.formula).toBe('#REF!');
    expect(getCell(sheet2, 1, 0)!.formula).toBe('SUM(#REF!)');
    const undo = session.dispatch('edit.undo', undefined);
    expect(undo.changes).toHaveLength(0);
    expect(getCell(sheet2, 0, 0)!.formula).toBe('#REF!');
    expect(getCell(sheet2, 1, 0)!.formula).toBe('SUM(#REF!)');
    expect(session.workbook.sheets).toHaveLength(1);
    expect(rewriteReferencesToRefError('A1+Sheet1!B2', 'Sheet1')).toBe('A1+#REF!');
  });
});

describe('sheet save patches', () => {
  it('writes a new worksheet part, relationship and content type, and reopens it', () => {
    const original = buildMultiSheetXlsx(
      [
        sheet(
          'Sheet1',
          '<sheetPr><tabColor rgb="FFFF0000"/></sheetPr><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>',
        ),
        sheet('Sheet2', '<sheetData><row r="1"><c r="A1"><v>2</v></c></row></sheetData>', {
          state: 'hidden',
        }),
      ],
      { extraParts: { 'xl/charts/chart1.xml': 'chart-bytes' } },
    );
    const opened = openWorkbook(original);
    const session = new DocumentSession(opened.workbook);
    session.dispatch('sheet.insert', {});
    const inserted = session.workbook.sheets[2]!;
    session.dispatch('cell.set', {
      sheetId: inserted.id,
      row: 0,
      col: 0,
      value: { type: 'text', value: 'hello' },
    });
    session.dispatch('cell.set', {
      sheetId: inserted.id,
      row: 1,
      col: 1,
      value: { type: 'number', value: 5 },
      formula: 'A1',
    });

    const result = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
    });
    expect(result.noOp).toBe(false);
    expect(result.package.entries.has('xl/worksheets/sheet3.xml')).toBe(true);
    expect(
      byteEqual(
        result.package.entries.get('xl/worksheets/sheet1.xml')!,
        originalPart(original, 'xl/worksheets/sheet1.xml'),
      ),
    ).toBe(true);
    expect(
      byteEqual(
        result.package.entries.get('xl/worksheets/sheet2.xml')!,
        originalPart(original, 'xl/worksheets/sheet2.xml'),
      ),
    ).toBe(true);
    expect(
      byteEqual(
        result.package.entries.get('xl/charts/chart1.xml')!,
        originalPart(original, 'xl/charts/chart1.xml'),
      ),
    ).toBe(true);
    expect(unzipPart(result.bytes, 'xl/workbook.xml')).toContain(
      '<sheet name="Sheet3" sheetId="3" r:id="rId3"/>',
    );
    expect(unzipPart(result.bytes, 'xl/workbook.xml')).toContain('state="hidden"');
    expect(unzipPart(result.bytes, 'xl/_rels/workbook.xml.rels')).toContain(
      'Target="worksheets/sheet3.xml"',
    );
    expect(unzipPart(result.bytes, '[Content_Types].xml')).toContain(
      'PartName="/xl/worksheets/sheet3.xml"',
    );

    const reopened = openWorkbook(result.bytes);
    expect(reopened.workbook.sheets.map((entry) => entry.name)).toEqual([
      'Sheet1',
      'Sheet2',
      'Sheet3',
    ]);
    expect(getCell(reopened.workbook.sheets[2]!, 0, 0)?.value).toEqual({
      type: 'text',
      value: 'hello',
    });
    expect(getCell(reopened.workbook.sheets[2]!, 1, 1)?.formula).toBe('A1');
    expect(reopened.workbook.sheets[2]!.view.zoom).toBe(100);
    expect(
      Buffer.from(reopened.package.entries.get('xl/worksheets/sheet1.xml')!).toString('utf8'),
    ).toContain('tabColor');
  });

  it('renames in the workbook part only and rewrites referencing formulas on save', () => {
    const original = buildMultiSheetXlsx(
      [
        sheet('Sheet1', '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>', {
          worksheetXml: worksheet('<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>'),
        }),
        sheet(
          'Sheet2',
          '<sheetData><row r="1"><c r="A1"><f>Sheet1!A1</f><v>1</v></c></row></sheetData>',
        ),
      ],
      {
        workbookExtras:
          '<definedNames><definedName name="DataRange">Sheet1!$A$1</definedName></definedNames>',
      },
    );
    const opened = openWorkbook(original);
    const session = new DocumentSession(opened.workbook);
    const sheet1 = session.workbook.sheets[0]!.id;
    session.dispatch('sheet.rename', { sheetId: sheet1, name: 'Data' });

    const result = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
    });
    const workbookXml = unzipPart(result.bytes, 'xl/workbook.xml');
    expect(workbookXml).toContain('<sheet name="Data" sheetId="1" r:id="rId1"/>');
    expect(workbookXml).toContain('<definedName name="DataRange">Data!$A$1</definedName>');
    expect(unzipPart(result.bytes, 'xl/worksheets/sheet1.xml')).toBe(
      unzipPart(original, 'xl/worksheets/sheet1.xml'),
    );
    expect(unzipPart(result.bytes, 'xl/worksheets/sheet2.xml')).toContain('<f>Data!A1</f>');

    const reopened = openWorkbook(result.bytes);
    expect(reopened.workbook.sheets[0]!.name).toBe('Data');
    expect(reopened.workbook.sheets[1]!.cells.get('0:0')?.formula).toBe('Data!A1');
    expect(reopened.workbook.names[0]!.formula).toBe('Data!$A$1');
  });

  it('removes the worksheet part and its package references on delete', () => {
    const original = buildMultiSheetXlsx(
      [
        sheet('Sheet1', '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>'),
        sheet('Sheet2', '<sheetData><row r="1"><c r="A1"><v>2</v></c></row></sheetData>'),
        sheet('Sheet3', '<sheetData><row r="1"><c r="A1"><v>3</v></c></row></sheetData>'),
      ],
      { extraParts: { 'xl/charts/chart1.xml': 'chart-bytes' } },
    );
    const opened = openWorkbook(original);
    const session = new DocumentSession(opened.workbook);
    session.dispatch('sheet.delete', {
      sheetId: session.workbook.sheets[1]!.id,
      confirm: true,
    });
    const result = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
    });
    expect(result.package.entries.has('xl/worksheets/sheet2.xml')).toBe(false);
    expect(unzipPart(result.bytes, 'xl/workbook.xml')).not.toContain('name="Sheet2"');
    expect(unzipPart(result.bytes, 'xl/_rels/workbook.xml.rels')).not.toContain('sheet2.xml');
    expect(unzipPart(result.bytes, '[Content_Types].xml')).not.toContain(
      '/xl/worksheets/sheet2.xml',
    );
    expect(
      byteEqual(
        result.package.entries.get('xl/charts/chart1.xml')!,
        originalPart(original, 'xl/charts/chart1.xml'),
      ),
    ).toBe(true);
    expect(unzipPart(result.bytes, 'xl/worksheets/sheet1.xml')).toBe(
      unzipPart(original, 'xl/worksheets/sheet1.xml'),
    );
    expect(unzipPart(result.bytes, 'xl/worksheets/sheet3.xml')).toBe(
      unzipPart(original, 'xl/worksheets/sheet3.xml'),
    );
    const reopened = openWorkbook(result.bytes);
    expect(reopened.workbook.sheets.map((entry) => entry.name)).toEqual(['Sheet1', 'Sheet3']);
  });

  it('deletes a sheet permanently: save removes its parts and names, and undo cannot restore it', () => {
    const original = buildMultiSheetXlsx(
      [
        {
          name: 'Sheet1',
          state: 'hidden',
          worksheetXml: attachedWorksheet(`
  <sheetPr><tabColor rgb="FFFF0000"/></sheetPr>
  <dimension ref="A1"/>
  <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
  <drawing r:id="rId1"/>
  <legacyDrawing r:id="rId2"/>
  <extLst><ext uri="urn:calcplus:test"><calcplus:keep xmlns:calcplus="urn:calcplus">preserve-me</calcplus:keep></ext></extLst>`),
        },
        sheet('Sheet2', '<sheetData><row r="1"><c r="A1"><v>2</v></c></row></sheetData>'),
      ],
      {
        workbookExtras:
          '<definedNames><definedName name="LocalName" localSheetId="0" hidden="1">Sheet1!$A$1</definedName></definedNames>',
        extraParts: {
          'xl/drawings/drawing1.xml':
            '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>',
          'xl/drawings/_rels/drawing1.xml.rels': `${XML_HEADER}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${RELATIONSHIPS_NS}/image" Target="../media/image1.png"/>
</Relationships>`,
          'xl/drawings/vmlDrawing1.vml': 'vml-bytes',
          'xl/media/image1.png': 'image-bytes',
          'xl/comments1.xml': `${XML_HEADER}
<comments xmlns="${SPREADSHEET_NS}"><authors><author>Test</author></authors><commentList/></comments>`,
          'xl/worksheets/_rels/sheet1.xml.rels': `${XML_HEADER}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${RELATIONSHIPS_NS}/drawing" Target="../drawings/drawing1.xml"/>
  <Relationship Id="rId2" Type="${RELATIONSHIPS_NS}/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>
  <Relationship Id="rId3" Type="${RELATIONSHIPS_NS}/comments" Target="../comments1.xml"/>
</Relationships>`,
        },
      },
    );
    const opened = openWorkbook(original);
    const session = new DocumentSession(opened.workbook);
    const sheet1 = session.workbook.sheets[0]!.id;
    session.dispatch('cell.set', {
      sheetId: sheet1,
      row: 2,
      col: 0,
      value: { type: 'number', value: 9 },
    });
    session.dispatch('sheet.delete', {
      sheetId: sheet1,
      confirm: true,
    });
    expect(session.canUndo).toBe(false);
    expect(session.hasUnsavedChanges).toBe(true);
    const deleted = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
    });
    const removedPaths = [
      'xl/worksheets/sheet1.xml',
      'xl/worksheets/_rels/sheet1.xml.rels',
      'xl/drawings/drawing1.xml',
      'xl/drawings/_rels/drawing1.xml.rels',
      'xl/drawings/vmlDrawing1.vml',
      'xl/media/image1.png',
      'xl/comments1.xml',
    ];
    for (const path of removedPaths) {
      expect(deleted.package.entries.has(path)).toBe(false);
      // The terminal delete prunes the model too; no skeleton waits for undo.
      expect(opened.workbook.sourceParts.has(path)).toBe(false);
      expect(opened.workbook.opaqueParts.has(path)).toBe(false);
    }
    expect(unzipPart(deleted.bytes, 'xl/workbook.xml')).not.toContain('LocalName');
    expect(unzipPart(deleted.bytes, 'xl/workbook.xml')).not.toContain('Sheet1');

    session.markSaved(session.dirtyCellSnapshot(), session.dirtyDimensionSnapshot());
    expect(session.hasUnsavedChanges).toBe(false);

    // There is no history left to undo, and the deleted sheet never returns.
    const undo = session.dispatch('edit.undo', undefined);
    expect(undo.changes).toHaveLength(0);
    expect(session.workbook.sheets.map((entry) => entry.name)).toEqual(['Sheet2']);
    const second = saveWorkbook({
      workbook: opened.workbook,
      package: deleted.package,
      dirtyCells: session.dirtyCellMap(),
    });
    expect(second.noOp).toBe(true);
    expect(byteEqual(second.bytes, deleted.bytes)).toBe(true);
    for (const path of removedPaths) {
      expect(second.package.entries.has(path)).toBe(false);
    }

    const reopened = openWorkbook(second.bytes);
    expect(reopened.workbook.sheets.map((entry) => entry.name)).toEqual(['Sheet2']);
    expect(reopened.workbook.names).toHaveLength(0);
    expect(reopened.workbook.sheets[0]!.names).toHaveLength(0);
  });

  it('preserves imported tab order and inserts at the requested position', () => {
    const original = buildMultiSheetXlsx([
      sheet('Zeta', '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>'),
      sheet('Alpha', '<sheetData><row r="1"><c r="A1"><v>2</v></c></row></sheetData>'),
      sheet('Middle', '<sheetData><row r="1"><c r="A1"><v>3</v></c></row></sheetData>'),
    ]);
    const opened = openWorkbook(original);
    expect(opened.workbook.sheets.map((entry) => entry.name)).toEqual(['Zeta', 'Alpha', 'Middle']);
    const session = new DocumentSession(opened.workbook);
    session.dispatch('sheet.insert', { name: 'Inserted', index: 1 });
    const result = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
    });
    const reopened = openWorkbook(result.bytes);
    expect(reopened.workbook.sheets.map((entry) => entry.name)).toEqual([
      'Zeta',
      'Inserted',
      'Alpha',
      'Middle',
    ]);
    expect(unzipPart(result.bytes, 'xl/worksheets/sheet1.xml')).toBe(
      unzipPart(original, 'xl/worksheets/sheet1.xml'),
    );
    expect(unzipPart(result.bytes, 'xl/worksheets/sheet3.xml')).toBe(
      unzipPart(original, 'xl/worksheets/sheet3.xml'),
    );
  });

  it('shifts localSheetId on insert and delete and removes a deleted sheet’s names', () => {
    const original = buildMultiSheetXlsx(
      [
        sheet('Sheet1', '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>'),
        sheet('Sheet2', '<sheetData><row r="1"><c r="A1"><v>2</v></c></row></sheetData>'),
        sheet('Sheet3', '<sheetData><row r="1"><c r="A1"><v>3</v></c></row></sheetData>'),
      ],
      {
        workbookExtras:
          '<definedNames><definedName name="OnSecond" localSheetId="1">Sheet2!$A$1</definedName><definedName name="OnThird" localSheetId="2">Sheet3!$A$1</definedName></definedNames>',
      },
    );
    const opened = openWorkbook(original);
    const session = new DocumentSession(opened.workbook);
    session.dispatch('sheet.insert', { name: 'First', index: 0 });
    const inserted = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
    });
    const insertedWorkbookXml = unzipPart(inserted.bytes, 'xl/workbook.xml');
    expect(insertedWorkbookXml).toContain(
      '<definedName name="OnSecond" localSheetId="2">Sheet2!$A$1</definedName>',
    );
    expect(insertedWorkbookXml).toContain(
      '<definedName name="OnThird" localSheetId="3">Sheet3!$A$1</definedName>',
    );
    const afterInsert = openWorkbook(inserted.bytes);
    expect(afterInsert.workbook.sheets.map((entry) => entry.name)).toEqual([
      'First',
      'Sheet1',
      'Sheet2',
      'Sheet3',
    ]);
    expect(afterInsert.workbook.sheets[2]!.names.map((name) => name.name)).toEqual(['OnSecond']);

    const deleteSession = new DocumentSession(afterInsert.workbook);
    deleteSession.dispatch('sheet.delete', {
      sheetId: afterInsert.workbook.sheets[2]!.id,
      confirm: true,
    });
    const deleted = saveWorkbook({
      workbook: afterInsert.workbook,
      package: afterInsert.package,
      dirtyCells: deleteSession.dirtyCellMap(),
    });
    const deletedWorkbookXml = unzipPart(deleted.bytes, 'xl/workbook.xml');
    expect(deletedWorkbookXml).not.toContain('OnSecond');
    expect(deletedWorkbookXml).toContain(
      '<definedName name="OnThird" localSheetId="2">Sheet3!$A$1</definedName>',
    );
    const afterDelete = openWorkbook(deleted.bytes);
    expect(afterDelete.workbook.sheets.map((entry) => entry.name)).toEqual([
      'First',
      'Sheet1',
      'Sheet3',
    ]);
    expect(afterDelete.workbook.sheets[2]!.names.map((name) => name.name)).toEqual(['OnThird']);
  });

  it('rewrites structural formulas and names after a snapshot restore', () => {
    const opened = openWorkbook(baseWorkbook());
    const session = new DocumentSession(opened.workbook);
    session.dispatch('sheet.rename', {
      sheetId: session.workbook.sheets[0]!.id,
      name: 'Data',
    });
    const snapshot = snapshotWorkbook(opened.workbook, opened.package);
    const restored = restoreWorkbook(snapshot);
    const result = saveWorkbook({
      workbook: restored.workbook,
      package: restored.package,
      dirtyCells: new Map(),
    });
    expect(result.noOp).toBe(false);
    expect(unzipPart(result.bytes, 'xl/workbook.xml')).toContain('<sheet name="Data"');
    expect(unzipPart(result.bytes, 'xl/workbook.xml')).toContain(
      '<definedName name="DataRange">Data!$A$1</definedName>',
    );
    expect(unzipPart(result.bytes, 'xl/worksheets/sheet2.xml')).toContain('<f>Data!A1</f>');
    const reopened = openWorkbook(result.bytes);
    expect(reopened.workbook.sheets[1]!.cells.get('0:0')?.formula).toBe('Data!A1');
  });

  it('returns the original bytes when an insert is undone before saving', () => {
    const original = baseWorkbook();
    const opened = openWorkbook(original);
    const session = new DocumentSession(opened.workbook);
    session.dispatch('sheet.insert', {});
    session.dispatch('edit.undo', undefined);
    const result = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
    });
    expect(result.noOp).toBe(true);
    expect(byteEqual(result.bytes, original)).toBe(true);
  });

  it('stays dirty when a sheet command runs after the save snapshot', () => {
    const opened = openWorkbook(baseWorkbook());
    const session = new DocumentSession(opened.workbook);
    // A save captures the version before its file-write await. A rename that
    // lands during the await is not in the saved bytes, so the workbook must
    // stay dirty and the close prompt must still warn.
    const savedStructuralVersion = session.structuralVersion;
    session.dispatch('sheet.rename', {
      sheetId: session.workbook.sheets[0]!.id,
      name: 'During write',
    });
    session.markSaved(
      session.dirtyCellSnapshot(),
      session.dirtyDimensionSnapshot(),
      savedStructuralVersion,
    );
    expect(session.hasUnsavedChanges).toBe(true);

    session.markSaved(
      session.dirtyCellSnapshot(),
      session.dirtyDimensionSnapshot(),
      session.structuralVersion,
    );
    expect(session.hasUnsavedChanges).toBe(false);
  });
});

describe('structural rebuild routing', () => {
  it('rebuilds the engine for sheet and defined-name changes but not cell edits', () => {
    expect(
      changesRequireRebuild([
        { type: 'sheet.rename', sheetId: 'sheet:x', before: 'A', after: 'B' },
      ]),
    ).toBe(true);
    expect(
      changesRequireRebuild([
        {
          type: 'name.set',
          scope: 'workbook',
          sheetId: null,
          name: 'N',
          before: 'Sheet1!A1',
          after: 'Data!A1',
        },
      ]),
    ).toBe(true);
    expect(
      changesRequireRebuild([
        {
          type: 'cell.set',
          sheetId: 'sheet:x',
          range: { startRow: 0, startCol: 0, endRow: 0, endCol: 0 },
          before: [[null]],
          after: [[null]],
        },
      ]),
    ).toBe(false);
  });
});

describe('rename planning', () => {
  it('blocks a 3D reference rather than half-rewriting it', () => {
    const session = new DocumentSession(openWorkbook(baseWorkbook()).workbook);
    session.workbook.sheets[1]!.cells.set('2:0', {
      value: { type: 'blank' },
      formula: 'SUM(Sheet1:Sheet2!A1)',
      formulaDerived: false,
      savedResult: null,
      styleId: 0,
      richText: null,
      stale: false,
    });
    const plan = planSheetRename(session.workbook, session.workbook.sheets[0]!.id, 'Data');
    expect(plan.blockedReason).toMatch(/3D reference/);
  });

  it('keeps the command registry stable', () => {
    const registry = createCommandRegistry();
    for (const id of ['sheet.insert', 'sheet.delete', 'sheet.rename']) {
      expect(registry.require(id).id).toBe(id);
    }
    expect(new Set(registry.list().map((command) => command.id)).size).toBe(registry.list().length);
  });
});

function originalPart(bytes: Uint8Array, name: string): Uint8Array {
  return unzipParts(bytes).get(name)!;
}
