import { Buffer } from 'node:buffer';
import { describe, expect, it } from 'vitest';
import { DocumentSession } from '@commands/session';
import { editorHasChanges } from '@grid/editor';
import { confirmReplaceDocument } from '@io/confirm';
import { numberValue, textValue } from '@model/value';
import { setCellSnapshot, type Cell } from '@model/types';
import { openBlankWorkbook, createBlankWorkbookBytes } from '@ooxml/blank';
import { WorkbookSaveError } from '@ooxml/errors';
import { openWorkbook } from '@ooxml/open';
import { saveWorkbook } from '@ooxml/save';
import { validatePatchedPart } from '@ooxml/xml';
import { buildXlsx, unzipPart } from './helpers/build-xlsx';

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const ws = (rows: string): string =>
  `${XML_HEADER}<worksheet xmlns="${NS}"><sheetData>${rows}</sheetData></worksheet>`;

function saveOpened(opened: ReturnType<typeof openWorkbook>, session: DocumentSession) {
  return saveWorkbook({
    workbook: opened.workbook,
    package: opened.package,
    dirtyCells: session.dirtyCellMap(),
  });
}

function formulaCellXml(xml: string, ref: string): string | undefined {
  return new RegExp(`<c r="${ref}"[^>]*>[\\s\\S]*?</c>`).exec(xml)?.[0];
}

describe('P0 worksheet cell ordering', () => {
  it('inserts a new cell between existing cells in column order', () => {
    const opened = openWorkbook(
      buildXlsx(ws('<row r="1"><c r="A1"><v>1</v></c><c r="C1"><v>3</v></c></row>')),
    );
    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    session.dispatch('cell.set', { sheetId: sheet.id, row: 0, col: 1, value: numberValue(2) });
    const result = saveOpened(opened, session);
    const xml = unzipPart(result.bytes, 'xl/worksheets/sheet1.xml');
    expect([...xml.matchAll(/<c r="([A-Z]+1)"/g)].map((match) => match[1])).toEqual([
      'A1',
      'B1',
      'C1',
    ]);
    expect(result.report.modeByPart.get('xl/worksheets/sheet1.xml')).toBe('position');
    const reopened = openWorkbook(result.bytes);
    expect(reopened.workbook.sheets[0]!.cells.get('0:1')?.value).toEqual({
      type: 'number',
      value: 2,
    });
  });

  it('inserts a new first cell before an existing cell', () => {
    const opened = openWorkbook(buildXlsx(ws('<row r="1"><c r="B1"><v>2</v></c></row>')));
    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    session.dispatch('cell.set', { sheetId: sheet.id, row: 0, col: 0, value: numberValue(1) });
    const result = saveOpened(opened, session);
    const xml = unzipPart(result.bytes, 'xl/worksheets/sheet1.xml');
    expect([...xml.matchAll(/<c r="([A-Z]+1)"/g)].map((match) => match[1])).toEqual(['A1', 'B1']);
    expect(result.report.modeByPart.get('xl/worksheets/sheet1.xml')).toBe('position');
  });

  it('inserts several new cells around an existing cell in column order', () => {
    const opened = openWorkbook(buildXlsx(ws('<row r="1"><c r="C1"><v>3</v></c></row>')));
    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    session.dispatch('cell.set', { sheetId: sheet.id, row: 0, col: 0, value: numberValue(1) });
    session.dispatch('cell.set', { sheetId: sheet.id, row: 0, col: 1, value: numberValue(2) });
    session.dispatch('cell.set', { sheetId: sheet.id, row: 0, col: 3, value: numberValue(4) });
    const result = saveOpened(opened, session);
    const xml = unzipPart(result.bytes, 'xl/worksheets/sheet1.xml');
    expect([...xml.matchAll(/<c r="([A-Z]+1)"/g)].map((match) => match[1])).toEqual([
      'A1',
      'B1',
      'C1',
      'D1',
    ]);
  });
});

describe('P0 shared and array formulas', () => {
  const sharedRows =
    '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="shared" ref="B1:B2" si="0">A1*2</f><v>2</v></c></row>' +
    '<row r="2"><c r="A2"><v>5</v></c><c r="B2"><f t="shared" si="0"/><v>10</v></c></row>';
  const textSharedRows =
    '<row r="1"><c r="A1"><v>1</v></c><c r="B1" t="str"><f t="shared" ref="B1:B2" si="0">A1&amp;"!"</f><v>x</v></c></row>' +
    '<row r="2"><c r="A2"><v>5</v></c><c r="B2" t="str"><f t="shared" si="0"/><v>Hello</v></c></row>';

  it('safely refuses a shared-formula master edit instead of breaking the group', () => {
    const opened = openWorkbook(buildXlsx(ws(sharedRows)));
    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    const result = session.dispatch('cell.set', {
      sheetId: sheet.id,
      row: 0,
      col: 1,
      value: numberValue(7),
    });
    expect(result.changes).toHaveLength(0);
    expect(result.ui?.message).toMatch(/shared formulas are read-only/i);
    expect(session.hasUnsavedChanges).toBe(false);

    const saved = saveOpened(opened, session);
    expect(saved.noOp).toBe(true);
    const xml = unzipPart(saved.bytes, 'xl/worksheets/sheet1.xml');
    expect(xml).toContain('<f t="shared" ref="B1:B2" si="0">A1*2</f>');
    expect(formulaCellXml(xml, 'B2')).toContain('<f t="shared" si="0"/>');
  });

  it('refuses a direct model edit of a shared master at save time', () => {
    const opened = openWorkbook(buildXlsx(ws(sharedRows)));
    const sheet = opened.workbook.sheets[0]!;
    setCellSnapshot(sheet, 0, 1, {
      value: numberValue(7),
      formula: null,
      formulaDerived: false,
      savedResult: null,
      styleId: 0,
      richText: null,
    });
    expect(() =>
      saveWorkbook({
        workbook: opened.workbook,
        package: opened.package,
        dirtyCells: new Map([[sheet.id, new Set(['0:1'])]]),
      }),
    ).toThrow(/shared formula.*read-only|read-only.*shared formula/i);
  });

  it('treats legacy array formulas as a user-visible no-op', () => {
    const rows =
      '<row r="1"><c r="A1"><f t="array" ref="A1:A2">B1:B2*2</f><v>2</v></c><c r="B1"><v>1</v></c></row>' +
      '<row r="2"><c r="A2"><v>4</v></c><c r="B2"><v>2</v></c></row>';
    const opened = openWorkbook(buildXlsx(ws(rows)));
    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    const result = session.dispatch('cell.set', {
      sheetId: sheet.id,
      row: 0,
      col: 0,
      value: numberValue(9),
    });
    expect(result.changes).toHaveLength(0);
    expect(result.ui?.message).toMatch(/array formulas are read-only/i);
    expect(session.hasUnsavedChanges).toBe(false);
    const saved = saveOpened(opened, session);
    const xml = unzipPart(saved.bytes, 'xl/worksheets/sheet1.xml');
    expect(xml).toContain('<f t="array" ref="A1:A2">B1:B2*2</f>');

    const follower = session.dispatch('cell.set', {
      sheetId: sheet.id,
      row: 1,
      col: 0,
      value: numberValue(4),
    });
    expect(follower.changes).toHaveLength(0);
    expect(follower.ui?.message).toMatch(/array formulas are read-only/i);

    setCellSnapshot(sheet, 1, 0, {
      value: numberValue(9),
      formula: null,
      formulaDerived: false,
      savedResult: null,
      styleId: 0,
      richText: null,
    });
    expect(() =>
      saveWorkbook({
        workbook: opened.workbook,
        package: opened.package,
        dirtyCells: new Map([[sheet.id, new Set(['1:0'])]]),
      }),
    ).toThrow(/array formulas are read-only/i);
  });

  it('rejects a shared-formula follower edit before it can dirty the document', () => {
    const original = buildXlsx(ws(sharedRows));
    const opened = openWorkbook(original);
    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    const result = session.dispatch('cell.set', {
      sheetId: sheet.id,
      row: 1,
      col: 1,
      value: numberValue(11),
    });
    expect(result.changes).toHaveLength(0);
    expect(result.ui?.message).toMatch(/shared formulas are read-only/i);
    expect(session.hasUnsavedChanges).toBe(false);
    expect(session.canUndo).toBe(false);

    session.undo();
    const saved = saveOpened(opened, session);
    expect(saved.noOp).toBe(true);
    expect(Buffer.compare(Buffer.from(saved.bytes), Buffer.from(original))).toBe(0);
    const xml = unzipPart(saved.bytes, 'xl/worksheets/sheet1.xml');
    expect(formulaCellXml(xml, 'B2')).toContain('<f t="shared" si="0"/>');
    expect(formulaCellXml(xml, 'B2')).toContain('<v>10</v>');
    expect(formulaCellXml(xml, 'B1')).toContain('ref="B1:B2"');
  });

  it('rejects a text-cached follower edit and preserves t="str" through save', () => {
    const original = buildXlsx(ws(textSharedRows), { sharedStrings: ['Existing'] });
    const opened = openWorkbook(original);
    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    const result = session.dispatch('cell.set', {
      sheetId: sheet.id,
      row: 1,
      col: 1,
      value: textValue('Changed'),
    });
    expect(result.changes).toHaveLength(0);
    expect(result.ui?.message).toMatch(/shared formulas are read-only/i);
    expect(session.hasUnsavedChanges).toBe(false);
    expect(session.canUndo).toBe(false);

    session.undo();
    const identity = saveOpened(opened, session);
    expect(identity.noOp).toBe(true);
    expect(Buffer.compare(Buffer.from(identity.bytes), Buffer.from(original))).toBe(0);

    // An unrelated edit forces a real patched save, not just the no-op path.
    session.dispatch('cell.set', { sheetId: sheet.id, row: 0, col: 0, value: numberValue(2) });
    const saved = saveOpened(opened, session);
    expect(saved.noOp).toBe(false);
    const xml = unzipPart(saved.bytes, 'xl/worksheets/sheet1.xml');
    expect(formulaCellXml(xml, 'B2')).toBe(
      '<c r="B2" t="str"><f t="shared" si="0"/><v>Hello</v></c>',
    );
    expect(formulaCellXml(xml, 'B1')).toContain('ref="B1:B2"');
    expect(unzipPart(saved.bytes, 'xl/sharedStrings.xml')).toBe(
      unzipPart(original, 'xl/sharedStrings.xml'),
    );

    const reopened = openWorkbook(saved.bytes);
    const follower = reopened.workbook.sheets[0]!.cells.get('1:1')!;
    expect(follower.value).toEqual({ type: 'text', value: 'Hello' });
    expect(follower.savedResult).toEqual({ type: 'text', value: 'Hello' });
    expect(follower.formulaDerived).toBe(true);
    expect(follower.formulaSource?.kind).toBe('shared-follower');
  });

  it('safely refuses a direct model edit of a text-cached follower at save time', () => {
    const opened = openWorkbook(buildXlsx(ws(textSharedRows)));
    const sheet = opened.workbook.sheets[0]!;
    setCellSnapshot(sheet, 1, 1, {
      value: textValue('Changed'),
      formula: null,
      formulaDerived: false,
      savedResult: null,
      styleId: 0,
      richText: null,
    });
    expect(() =>
      saveWorkbook({
        workbook: opened.workbook,
        package: opened.package,
        dirtyCells: new Map([[sheet.id, new Set(['1:1'])]]),
      }),
    ).toThrow(/shared formulas are read-only/i);
  });

  it('refuses a bypass that changes only a follower cached text result', () => {
    const opened = openWorkbook(buildXlsx(ws(textSharedRows)));
    const sheet = opened.workbook.sheets[0]!;
    const follower = sheet.cells.get('1:1')!;
    setCellSnapshot(sheet, 1, 1, {
      ...follower,
      value: textValue('Changed'),
      savedResult: textValue('Changed'),
    });
    expect(() =>
      saveWorkbook({
        workbook: opened.workbook,
        package: opened.package,
        dirtyCells: new Map([[sheet.id, new Set(['1:1'])]]),
      }),
    ).toThrow(/shared formulas are read-only/i);
  });
});

describe('P1 rich text preservation', () => {
  const richRows = '<row r="1"><c r="A1" t="s"><v>0</v></c></row>';
  const richSharedStrings =
    `${XML_HEADER}<sst xmlns="${NS}" count="1" uniqueCount="1">` +
    '<si><r><rPr><b/></rPr><t>Bold</t></r><r><t xml:space="preserve"> plain</t></r></si></sst>';
  const richWorkbook = (): Uint8Array =>
    buildXlsx(ws(richRows), {
      sharedStrings: ['placeholder'],
      extraParts: { 'xl/sharedStrings.xml': richSharedStrings },
    });

  it('blocks an M1 editor edit that would flatten rich text', () => {
    const original = richWorkbook();
    const opened = openWorkbook(original);
    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    expect(sheet.cells.get('0:0')?.richText).not.toBeNull();

    const result = session.dispatch('cell.set', {
      sheetId: sheet.id,
      row: 0,
      col: 0,
      value: textValue('Bold plain'),
    });
    expect(result.changes).toHaveLength(0);
    expect(result.ui?.message).toMatch(/rich text/i);
    expect(result.ui?.message).toMatch(/read-only|flatten/i);
    expect(session.hasUnsavedChanges).toBe(false);
    expect(session.canUndo).toBe(false);

    session.undo();
    const saved = saveOpened(opened, session);
    expect(saved.noOp).toBe(true);
    expect(Buffer.compare(Buffer.from(saved.bytes), Buffer.from(original))).toBe(0);
  });

  it('preserves the original rich runs through edit, undo and an unrelated save', () => {
    const original = richWorkbook();
    const opened = openWorkbook(original);
    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    session.dispatch('cell.set', {
      sheetId: sheet.id,
      row: 0,
      col: 0,
      value: textValue('Bold plain'),
    });
    session.undo();
    session.dispatch('cell.set', { sheetId: sheet.id, row: 0, col: 1, value: numberValue(7) });

    const saved = saveOpened(opened, session);
    const beforeXml = unzipPart(original, 'xl/worksheets/sheet1.xml');
    const afterXml = unzipPart(saved.bytes, 'xl/worksheets/sheet1.xml');
    expect(formulaCellXml(afterXml, 'A1')).toBe(formulaCellXml(beforeXml, 'A1'));
    expect(unzipPart(saved.bytes, 'xl/sharedStrings.xml')).toBe(richSharedStrings);
    expect(unzipPart(saved.bytes, 'xl/sharedStrings.xml')).toContain('<rPr><b/></rPr>');

    const reopened = openWorkbook(saved.bytes);
    const cell = reopened.workbook.sheets[0]!.cells.get('0:0')!;
    expect(cell.value).toEqual({ type: 'text', value: 'Bold plain' });
    expect(cell.richText?.xml).toContain('<rPr><b/></rPr>');
  });

  it('refuses a direct model edit that would flatten rich runs at save time', () => {
    const opened = openWorkbook(richWorkbook());
    const sheet = opened.workbook.sheets[0]!;
    const original = sheet.cells.get('0:0')!;
    setCellSnapshot(sheet, 0, 0, { ...original, richText: null });
    expect(() =>
      saveWorkbook({
        workbook: opened.workbook,
        package: opened.package,
        dirtyCells: new Map([[sheet.id, new Set(['0:0'])]]),
      }),
    ).toThrow(/rich text.*cannot be rewritten|rich text.*read-only/i);
  });
});

describe('P0 XML validity', () => {
  it('rejects control characters in a typed edit with a user-visible no-op', () => {
    const opened = openWorkbook(buildXlsx(ws('<row r="1"><c r="A1"><v>1</v></c></row>')));
    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    const result = session.dispatch('cell.set', {
      sheetId: sheet.id,
      row: 0,
      col: 0,
      value: textValue('a\u0001b'),
    });
    expect(result.changes).toHaveLength(0);
    expect(result.ui?.message).toMatch(/XML cannot store/i);
    expect(session.hasUnsavedChanges).toBe(false);
  });

  it('refuses to write a control character injected into the model', () => {
    const opened = openWorkbook(buildXlsx(ws('<row r="1"><c r="A1"><v>1</v></c></row>')));
    const sheet = opened.workbook.sheets[0]!;
    const original = sheet.cells.get('0:0')!;
    const poisoned: Cell = { ...original, value: textValue('a\u0001b') };
    setCellSnapshot(sheet, 0, 0, poisoned);
    expect(() =>
      saveWorkbook({
        workbook: opened.workbook,
        package: opened.package,
        dirtyCells: new Map([[sheet.id, new Set(['0:0'])]]),
      }),
    ).toThrow(WorkbookSaveError);
    expect(() =>
      saveWorkbook({
        workbook: opened.workbook,
        package: opened.package,
        dirtyCells: new Map([[sheet.id, new Set(['0:0'])]]),
      }),
    ).toThrow(/XML cannot store/i);
  });

  it('validates a patched part before it can be written', () => {
    expect(() =>
      validatePatchedPart(
        new TextEncoder().encode('<worksheet><sheetData></worksheet>'),
        'xl/worksheets/sheet1.xml',
      ),
    ).toThrow(WorkbookSaveError);
  });

  it('rejects a control character before damaging the shared-strings table', () => {
    const opened = openWorkbook(
      buildXlsx(ws('<row r="1"><c r="A1" t="s"><v>0</v></c></row>'), { sharedStrings: ['Name'] }),
    );
    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    const result = session.dispatch('cell.set', {
      sheetId: sheet.id,
      row: 0,
      col: 0,
      value: textValue('a\u0002b'),
    });
    expect(result.changes).toHaveLength(0);
    expect(saveOpened(opened, session).noOp).toBe(true);
  });

  it('refuses a find replace that would introduce a control character', () => {
    const opened = openWorkbook(
      buildXlsx(
        ws('<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><f>SUM(A2:A3)</f><v>0</v></c></row>'),
        { sharedStrings: ['SUM'] },
      ),
    );
    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    const version = session.version;

    const result = session.dispatch('find.replaceAll', {
      query: 'SUM',
      replacement: 'x\u0001y',
      scope: 'sheet',
    });

    expect(result.changes).toHaveLength(0);
    expect(result.replaced).toBeUndefined();
    expect(result.ui?.message).toMatch(/XML cannot store/i);
    expect(sheet.cells.get('0:0')?.value).toEqual(textValue('SUM'));
    expect(sheet.cells.get('0:1')?.formula).toBe('SUM(A2:A3)');
    expect(session.transactionCount).toBe(0);
    expect(session.hasUnsavedChanges).toBe(false);
    expect(session.version).toBe(version);
    expect(() => saveOpened(opened, session)).not.toThrow();
    expect(saveOpened(opened, session).noOp).toBe(true);
  });
});

describe('P1 calcChain, recalc flag and stale cached results', () => {
  function formulaFixture(options: { calcChain: boolean }) {
    return buildXlsx(
      ws('<row r="1"><c r="A1"><v>1</v></c><c r="B1"><f>A1*2</f><v>2</v></c></row>'),
      {
        ...(options.calcChain
          ? {
              extraParts: {
                'xl/calcChain.xml': `${XML_HEADER}<calcChain xmlns="${NS}"><c r="B1" i="1"/></calcChain>`,
              },
              workbookRelsExtras: `<Relationship Id="rId9" Type="${REL_NS}/calcChain" Target="calcChain.xml"/>`,
              contentTypeOverrides: [
                [
                  'xl/calcChain.xml',
                  'application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml',
                ] as const,
              ],
            }
          : {}),
      },
    );
  }

  it('deletes calcChain and sets fullCalcOnLoad when a formula changes', () => {
    const opened = openWorkbook(formulaFixture({ calcChain: true }));
    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    session.dispatch('cell.set', { sheetId: sheet.id, row: 0, col: 1, value: numberValue(7) });
    const result = saveOpened(opened, session);
    expect(result.package.entries.has('xl/calcChain.xml')).toBe(false);
    expect(result.report.removedParts).toContain('xl/calcChain.xml');
    expect(unzipPart(result.bytes, 'xl/_rels/workbook.xml.rels')).not.toContain('/calcChain');
    expect(unzipPart(result.bytes, '[Content_Types].xml')).not.toContain('/xl/calcChain.xml');
    expect(unzipPart(result.bytes, 'xl/workbook.xml')).toContain('fullCalcOnLoad="1"');
  });

  it('keeps a dependent cached result, marks it stale, and keeps calcChain when only inputs change', () => {
    const opened = openWorkbook(formulaFixture({ calcChain: true }));
    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    session.dispatch('cell.set', { sheetId: sheet.id, row: 0, col: 0, value: numberValue(5) });
    const result = saveOpened(opened, session);
    const xml = unzipPart(result.bytes, 'xl/worksheets/sheet1.xml');
    expect(xml).toContain('<c r="B1"><f>A1*2</f><v>2</v></c>');
    expect(unzipPart(result.bytes, 'xl/workbook.xml')).toContain('fullCalcOnLoad="1"');
    expect(result.package.entries.has('xl/calcChain.xml')).toBe(true);
    const reopened = openWorkbook(result.bytes);
    const b1 = reopened.workbook.sheets[0]!.cells.get('0:1')!;
    expect(b1.formula).toBe('A1*2');
    expect(b1.savedResult).toEqual({ type: 'number', value: 2 });
  });

  it('does not add a recalculation flag when the workbook has no formulas', () => {
    const opened = openWorkbook(buildXlsx(ws('<row r="1"><c r="A1"><v>1</v></c></row>')));
    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    session.dispatch('cell.set', { sheetId: sheet.id, row: 0, col: 0, value: numberValue(2) });
    const result = saveOpened(opened, session);
    expect(unzipPart(result.bytes, 'xl/workbook.xml')).not.toContain('fullCalcOnLoad="1"');
  });
});

describe('P1 new blank workbook', () => {
  it('creates one sheet with default settings through the open path', () => {
    const opened = openBlankWorkbook();
    expect(opened.workbook.sheets).toHaveLength(1);
    const sheet = opened.workbook.sheets[0]!;
    expect(sheet.name).toBe('Sheet1');
    expect(sheet.cells.size).toBe(0);
    expect(sheet.view).toEqual({
      freezeRows: 0,
      freezeCols: 0,
      showGridLines: true,
      showHeadings: true,
      zoom: 100,
    });
    expect(opened.workbook.date1904).toBe(false);
    expect(opened.compatibility.notes).toEqual([]);
  });

  it('no-op saves the blank template byte-for-byte and saves edits on Save As', () => {
    const bytes = createBlankWorkbookBytes();
    const opened = openWorkbook(bytes);
    const identity = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: new Map(),
    });
    expect(identity.noOp).toBe(true);
    expect(Buffer.compare(Buffer.from(identity.bytes), Buffer.from(bytes))).toBe(0);

    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    session.dispatch('cell.set', { sheetId: sheet.id, row: 0, col: 0, value: textValue('Hello') });
    const saved = saveOpened(opened, session);
    const reopened = openWorkbook(saved.bytes);
    expect(reopened.workbook.sheets[0]!.cells.get('0:0')?.value).toEqual({
      type: 'text',
      value: 'Hello',
    });
  });
});

describe('P1 dirty state, save race and prompts', () => {
  it('keeps edits made while a save write is in flight dirty', () => {
    const opened = openBlankWorkbook();
    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    session.dispatch('cell.set', { sheetId: sheet.id, row: 0, col: 0, value: numberValue(1) });
    const snapshot = session.dirtyCellSnapshot();

    // Simulate edits that land while the save is awaiting the file write.
    session.dispatch('cell.set', { sheetId: sheet.id, row: 0, col: 0, value: numberValue(2) });
    session.dispatch('cell.set', { sheetId: sheet.id, row: 0, col: 1, value: numberValue(3) });
    session.markSaved(snapshot);

    expect(session.dirtyCellMap().get(sheet.id)?.has('0:0')).toBe(true);
    expect(session.dirtyCellMap().get(sheet.id)?.has('0:1')).toBe(true);
  });

  it('clears cells that did not change during the save', () => {
    const opened = openBlankWorkbook();
    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    session.dispatch('cell.set', { sheetId: sheet.id, row: 0, col: 0, value: numberValue(1) });
    const snapshot = session.dirtyCellSnapshot();
    session.markSaved(snapshot);
    expect(session.hasUnsavedChanges).toBe(false);
  });

  it('prompts before replacing a document with unsaved edits', () => {
    const asked: string[] = [];
    expect(confirmReplaceDocument(false, (message) => (asked.push(message), true))).toBe(true);
    expect(asked).toHaveLength(0);
    expect(confirmReplaceDocument(true, (message) => (asked.push(message), false))).toBe(false);
    expect(asked[0]).toMatch(/unsaved changes/i);
  });

  it('does not commit an unchanged F2 edit', () => {
    expect(editorHasChanges({ text: '85', initialText: '85', commitIfChanged: true })).toBe(false);
    expect(editorHasChanges({ text: '86', initialText: '85', commitIfChanged: true })).toBe(true);
    expect(editorHasChanges({ text: '4', initialText: '4', commitIfChanged: false })).toBe(true);
  });
});
