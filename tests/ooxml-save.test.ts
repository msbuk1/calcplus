import { readFileSync } from 'node:fs';
import { Buffer } from 'node:buffer';
import { describe, expect, it } from 'vitest';
import fc from 'fast-check';
import { DocumentSession } from '@commands/session';
import { openWorkbook } from '@ooxml/open';
import { saveWorkbook } from '@ooxml/save';
import { WorkbookSaveError } from '@ooxml/errors';
import { createPartInventory } from '@ooxml/inventory';
import { MAX_TREE_PATCH_BYTES } from '@ooxml/xml';
import { errorValue, numberValue, textValue } from '@model/value';
import { findSheet } from '@model/workbook';
import { buildXlsx, directChildNames, elementRanges, unzipPart } from './helpers/build-xlsx';

const fixture = (name: string): Uint8Array =>
  new Uint8Array(readFileSync(new URL(`../fixtures/compat/${name}`, import.meta.url)));

function openedFixture(name = 'm1-chart-unknown.xlsx') {
  return openWorkbook(fixture(name));
}

function editCell(
  session: DocumentSession,
  sheetId: string,
  row: number,
  col: number,
  value: Parameters<DocumentSession['dispatch']>[1],
): void {
  session.dispatch('cell.set', { sheetId, row, col, value });
}

describe('identity save', () => {
  it('returns the original bytes byte-for-byte when nothing changed', () => {
    const original = fixture('m1-chart-unknown.xlsx');
    const opened = openWorkbook(original);
    const result = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: new Map(),
    });
    expect(result.noOp).toBe(true);
    expect(Buffer.compare(Buffer.from(result.bytes), Buffer.from(original))).toBe(0);
  });

  it('does not touch source bytes during a no-op save', () => {
    const opened = openedFixture();
    const worksheetBefore = opened.workbook.sourceParts.get('xl/worksheets/sheet1.xml')!;
    saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: new Map(),
    });
    expect(opened.workbook.sourceParts.get('xl/worksheets/sheet1.xml')).toBe(worksheetBefore);
  });
});

describe('dirty worksheet patch', () => {
  it('patches only the dirty cell, preserving every unmodelled child and its order', async () => {
    const original = fixture('m1-chart-unknown.xlsx');
    const opened = openWorkbook(original);
    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    editCell(session, sheet.id, 1, 1, numberValue(43));

    const result = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
    });

    const beforeXml = unzipPart(original, 'xl/worksheets/sheet1.xml');
    const afterXml = unzipPart(result.bytes, 'xl/worksheets/sheet1.xml');
    expect(directChildNames(afterXml)).toEqual(directChildNames(beforeXml));
    expect(directChildNames(afterXml)).toEqual([
      'sheetPr',
      'dimension',
      'sheetViews',
      'sheetFormatPr',
      'cols',
      'sheetData',
      'mergeCells',
      'phoneticPr',
      'drawing',
      'legacyDrawing',
      'extLst',
    ]);
    expect(afterXml).toContain('<cp:property name="calcplus.unknown"');
    expect(afterXml).toContain('preserve-me');
    expect(afterXml).toContain('<c r="B2" x14:foo="bar"><v>43</v></c>');
    expect(afterXml).not.toContain('<c r="B2"><v>42.5</v></c>');
    expect(afterXml).toContain('<legacyDrawing r:id="rId2"/>');
    expect(afterXml.slice(afterXml.indexOf('<extLst>'))).toBe(
      beforeXml.slice(beforeXml.indexOf('<extLst>')),
    );

    const beforeRows = elementRanges(beforeXml, 'row');
    const afterRows = elementRanges(afterXml, 'row');
    expect(beforeRows).toHaveLength(4);
    expect(afterRows).toHaveLength(4);
    for (const index of [0, 2, 3]) {
      const beforeRow = beforeRows[index]!;
      const afterRow = afterRows[index]!;
      expect(afterXml.slice(afterRow.start, afterRow.end)).toBe(
        beforeXml.slice(beforeRow.start, beforeRow.end),
      );
    }

    expect(afterXml).toContain('<c r="B3"><f>B2*2</f><v>85</v></c>');
    expect(afterXml).toContain('<mergeCell ref="A1:B1"/>');

    const originalParts = opened.package.entries;
    const savedParts = result.package.entries;
    expect(savedParts.size).toBe(originalParts.size);
    for (const [name, bytes] of originalParts) {
      if (name === 'xl/worksheets/sheet1.xml' || name === 'xl/workbook.xml') continue;
      expect(Buffer.compare(Buffer.from(savedParts.get(name)!), Buffer.from(bytes))).toBe(0);
    }
    // The edit can stale the B3 formula's cached result, so the workbook asks
    // Excel to recalculate while B3's original <v> is preserved.
    expect(unzipPart(result.bytes, 'xl/workbook.xml')).toContain('fullCalcOnLoad="1"');

    const beforeInventory = await createPartInventory(originalParts);
    const afterInventory = await createPartInventory(savedParts);
    const beforeByName = new Map(beforeInventory.map((entry) => [entry.name, entry]));
    for (const entry of afterInventory) {
      if (entry.name === 'xl/worksheets/sheet1.xml' || entry.name === 'xl/workbook.xml') {
        expect(entry.sha256).not.toBe(beforeByName.get(entry.name)!.sha256);
      } else {
        expect(entry).toEqual(beforeByName.get(entry.name));
      }
    }

    const reopened = openWorkbook(result.bytes);
    expect(reopened.workbook.sheets[0]!.cells.get('1:1')?.value).toEqual({
      type: 'number',
      value: 43,
    });
    expect(reopened.workbook.sheets[0]!.cells.get('2:1')?.formula).toBe('B2*2');
    expect(reopened.workbook.opaqueParts.has('xl/charts/chart1.xml')).toBe(true);
    expect(
      Buffer.compare(
        Buffer.from(reopened.workbook.opaqueParts.get('xl/charts/chart1.xml')!.bytes),
        Buffer.from(opened.workbook.opaqueParts.get('xl/charts/chart1.xml')!.bytes),
      ),
    ).toBe(0);
  });

  it('appends a new shared string without renumbering existing indices', () => {
    const original = fixture('m1-chart-unknown.xlsx');
    const opened = openWorkbook(original);
    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    editCell(session, sheet.id, 1, 0, textValue('Gamma'));

    const result = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
    });
    const sharedStrings = unzipPart(result.bytes, 'xl/sharedStrings.xml');
    expect(sharedStrings).toContain('<si><t>Name</t></si>');
    expect(sharedStrings).toContain('<si><t>Alpha</t></si>');
    expect(sharedStrings).toContain('<si><t xml:space="preserve">Gamma</t></si>');
    expect(sharedStrings).toMatch(/count="5"/);
    expect(sharedStrings).toMatch(/uniqueCount="5"/);
    expect(unzipPart(result.bytes, 'xl/worksheets/sheet1.xml')).toContain(
      '<c r="A2" t="s"><v>4</v></c>',
    );

    const reopened = openWorkbook(result.bytes);
    expect(reopened.workbook.sheets[0]!.cells.get('1:0')?.value).toEqual({
      type: 'text',
      value: 'Gamma',
    });
    expect(reopened.package.sharedStrings?.entries.slice(0, 4).map((entry) => entry.text)).toEqual([
      'Name',
      'Value',
      'Alpha',
      'Beta',
    ]);
  });

  it('reuses an existing shared string without changing the table', () => {
    const original = fixture('m1-chart-unknown.xlsx');
    const opened = openWorkbook(original);
    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    editCell(session, sheet.id, 1, 0, textValue('Alpha'));

    const result = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
    });
    expect(unzipPart(result.bytes, 'xl/sharedStrings.xml')).toBe(
      unzipPart(original, 'xl/sharedStrings.xml'),
    );
    expect(unzipPart(result.bytes, 'xl/worksheets/sheet1.xml')).toContain(
      '<c r="A2" t="s"><v>2</v></c>',
    );
  });

  it('preserves a formula cell byte-for-byte when another cell changes', () => {
    const original = fixture('m1-chart-unknown.xlsx');
    const opened = openWorkbook(original);
    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    editCell(session, sheet.id, 1, 0, textValue('Changed'));
    const result = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
    });
    expect(unzipPart(result.bytes, 'xl/worksheets/sheet1.xml')).toContain(
      '<c r="B3"><f>B2*2</f><v>85</v></c>',
    );
  });

  it('preserves macro and external-link parts byte-for-byte after an unrelated edit', () => {
    const worksheet =
      '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>';
    const original = buildXlsx(worksheet, {
      extraParts: {
        'xl/vbaProject.bin': 'fake macro bytes',
        'xl/externalLinks/externalLink1.xml':
          '<?xml version="1.0"?><externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>',
      },
    });
    const opened = openWorkbook(original);
    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    editCell(session, sheet.id, 0, 0, numberValue(2));
    const result = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
    });
    for (const path of ['xl/vbaProject.bin', 'xl/externalLinks/externalLink1.xml']) {
      expect(
        Buffer.compare(
          Buffer.from(result.package.entries.get(path)!),
          Buffer.from(opened.package.entries.get(path)!),
        ),
      ).toBe(0);
    }
  });

  it('expands a self-closing row once when several cells in it are edited', () => {
    const worksheet =
      '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="2"/></sheetData></worksheet>';
    const opened = openWorkbook(buildXlsx(worksheet));
    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    editCell(session, sheet.id, 1, 0, numberValue(2));
    editCell(session, sheet.id, 1, 2, numberValue(3));
    const result = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
    });
    const savedXml = unzipPart(result.bytes, 'xl/worksheets/sheet1.xml');
    expect(savedXml).toContain('<row r="2"><c r="A2"><v>2</v></c><c r="C2"><v>3</v></c></row>');
    const reopened = openWorkbook(result.bytes);
    expect(reopened.workbook.sheets[0]!.cells.get('1:0')?.value).toEqual({
      type: 'number',
      value: 2,
    });
    expect(reopened.workbook.sheets[0]!.cells.get('1:2')?.value).toEqual({
      type: 'number',
      value: 3,
    });
  });

  it('tree fallback updates a row whose r attribute is missing', () => {
    const worksheet =
      '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c></row></sheetData></worksheet>';
    const opened = openWorkbook(buildXlsx(worksheet));
    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    editCell(session, sheet.id, 0, 0, numberValue(9));
    const result = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
    });
    expect(result.report.modeByPart.get('xl/worksheets/sheet1.xml')).toBe('tree');
    const reopened = openWorkbook(result.bytes);
    expect(reopened.workbook.sheets[0]!.cells.get('0:0')?.value).toEqual({
      type: 'number',
      value: 9,
    });
    expect(reopened.workbook.sheets[0]!.cells.get('1:0')?.value).toEqual({
      type: 'number',
      value: 2,
    });
  });

  it('expands self-closing sheetData for the first edit', () => {
    const worksheet =
      '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>';
    const opened = openWorkbook(buildXlsx(worksheet));
    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    editCell(session, sheet.id, 0, 0, textValue('first'));
    const result = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
    });
    expect(result.report.modeByPart.get('xl/worksheets/sheet1.xml')).toBe('position');
    const savedXml = unzipPart(result.bytes, 'xl/worksheets/sheet1.xml');
    expect(savedXml).toContain('<row r="1"><c r="A1" t="inlineStr"><is>');
    expect(openWorkbook(result.bytes).workbook.sheets[0]!.cells.get('0:0')?.value).toEqual({
      type: 'text',
      value: 'first',
    });
  });

  it('writes a typed #CIRC! literal as text, never as an error value', () => {
    const worksheet =
      '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>';
    const opened = openWorkbook(buildXlsx(worksheet));
    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    editCell(session, sheet.id, 0, 0, errorValue('#CIRC!'));
    const result = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
    });
    // The authored literal survives as text, not as an error value or a cached
    // <v>; no value is omitted, so the workbook needs no recalculation flag.
    const savedXml = unzipPart(result.bytes, 'xl/worksheets/sheet1.xml');
    expect(savedXml).not.toContain('t="e"');
    expect(savedXml).not.toContain('<v>#CIRC!</v>');
    expect(savedXml).toContain(
      '<c r="A1" t="inlineStr"><is><t xml:space="preserve">#CIRC!</t></is></c>',
    );
    expect(openWorkbook(result.bytes).workbook.sheets[0]!.cells.get('0:0')?.value).toEqual({
      type: 'text',
      value: '#CIRC!',
    });
  });

  it('falls back to a tree rewrite for an unsafe worksheet shape and keeps unknown elements', () => {
    const worksheet = `<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c></row></sheetData><extLst><ext uri="{UNKNOWN}" xmlns:cp="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"><cp:properties><cp:property name="keep" pid="2">tree-fallback</cp:property></cp:properties></ext></extLst></worksheet>`;
    const original = buildXlsx(worksheet);
    const opened = openWorkbook(original);
    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    editCell(session, sheet.id, 1, 0, numberValue(43));
    const result = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
    });
    expect(result.report.modeByPart.get('xl/worksheets/sheet1.xml')).toBe('tree');
    const savedXml = unzipPart(result.bytes, 'xl/worksheets/sheet1.xml');
    expect(directChildNames(savedXml)).toEqual(['sheetData', 'extLst']);
    expect(savedXml).toContain('tree-fallback');
    expect(savedXml).toContain('<c r="A2"><v>43</v></c>');
    expect(openWorkbook(result.bytes).workbook.sheets[0]!.cells.get('1:0')?.value).toEqual({
      type: 'number',
      value: 43,
    });
  });

  it('refuses to patch an unsafe part at or above the tree-fallback threshold', () => {
    expect(MAX_TREE_PATCH_BYTES).toBe(16 * 1024 * 1024);
    const worksheet = `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c></row></sheetData></worksheet>`;
    const opened = openWorkbook(buildXlsx(worksheet));
    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    editCell(session, sheet.id, 1, 0, numberValue(3));
    expect(() =>
      saveWorkbook({
        workbook: opened.workbook,
        package: opened.package,
        dirtyCells: session.dirtyCellMap(),
        options: { worksheet: { maxTreeFallbackBytes: 1 } },
      }),
    ).toThrow(WorkbookSaveError);
    expect(session.hasUnsavedChanges).toBe(true);
  });
});

describe('formula results the engine evaluated without a command', () => {
  const worksheet =
    '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>21</v></c><c r="B1"><f>A1*2</f></c><c r="C1"><f>A1*2</f><v>42</v></c></row><row r="2"><c r="A2"><v>2</v></c></row></sheetData></worksheet>';

  function openedWithFormula() {
    const original = buildXlsx(worksheet);
    const opened = openWorkbook(original);
    const session = new DocumentSession(opened.workbook);
    return { original, opened, session, sheet: opened.workbook.sheets[0]! };
  }

  it('writes the evaluated <v> for a formula the file did not cache', () => {
    const { original, opened, session, sheet } = openedWithFormula();
    // The load phase applies results without a command, as on open.
    session.applyCalculationResults(
      [
        {
          sheetId: sheet.id,
          row: 0,
          col: 1,
          value: numberValue(42),
          status: 'value',
          stale: false,
        },
      ],
      { markDirty: false },
    );
    const result = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
    });
    expect(result.noOp).toBe(false);
    const savedXml = unzipPart(result.bytes, 'xl/worksheets/sheet1.xml');
    expect(savedXml).toContain('<c r="B1"><f>A1*2</f><v>42</v></c>');
    // C1 already agreed with the engine, so its authored bytes are untouched.
    expect(savedXml).toContain('<c r="C1"><f>A1*2</f><v>42</v></c>');
    expect(unzipPart(result.bytes, 'xl/workbook.xml')).toContain('fullCalcOnLoad="1"');
    expect(result.package.entries.get('xl/worksheets/sheet1.xml')).not.toBe(
      opened.package.entries.get('xl/worksheets/sheet1.xml'),
    );
    expect(original.byteLength).toBeGreaterThan(0);
  });

  it('leaves the file byte-identical when the evaluated result equals the cache', () => {
    const { original, opened, session, sheet } = openedWithFormula();
    session.applyCalculationResults(
      [
        {
          sheetId: sheet.id,
          row: 0,
          col: 2,
          value: numberValue(42),
          status: 'value',
          stale: false,
        },
      ],
      { markDirty: false },
    );
    const result = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
    });
    expect(result.noOp).toBe(true);
    expect(Buffer.compare(Buffer.from(result.bytes), Buffer.from(original))).toBe(0);
  });

  it('keeps the original <v> when the evaluated result is stale', () => {
    const { opened, session, sheet } = openedWithFormula();
    session.applyCalculationResults(
      [{ sheetId: sheet.id, row: 0, col: 2, value: numberValue(7), status: 'value', stale: true }],
      { markDirty: false },
    );
    expect(opened.workbook.sheets[0]!.cells.get('0:2')?.stale).toBe(true);
    // Force a write elsewhere so the cell is not saved through the no-op path.
    editCell(session, sheet.id, 1, 0, numberValue(3));
    const result = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
    });
    const savedXml = unzipPart(result.bytes, 'xl/worksheets/sheet1.xml');
    expect(savedXml).toContain('<c r="C1"><f>A1*2</f><v>42</v></c>');
    expect(savedXml).not.toContain('<v>7</v>');
    expect(unzipPart(result.bytes, 'xl/workbook.xml')).toContain('fullCalcOnLoad="1"');
  });
});

describe('used-range dimension covers every written cell', () => {
  const DIM_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const sheet = (dimension: string): Uint8Array =>
    buildXlsx(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<worksheet xmlns="${DIM_NS}"><dimension ref="${dimension}"/><sheetData>` +
        '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row>' +
        '<row r="2"><c r="A2"><v>3</v></c><c r="B2"><v>4</v></c></row>' +
        '</sheetData></worksheet>',
    );

  it('expands the declared range when a patch writes a cell outside it', () => {
    const opened = openWorkbook(sheet('A1:B2'));
    const session = new DocumentSession(opened.workbook);
    const model = opened.workbook.sheets[0]!;
    editCell(session, model.id, 4, 13, numberValue(9));

    const result = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
    });
    const xml = unzipPart(result.bytes, 'xl/worksheets/sheet1.xml');
    expect(xml).toContain('<c r="N5"');
    // A cell outside `<dimension ref>` is a spec violation and makes Excel
    // uncomfortable; the used range grows to cover the written cell.
    expect(xml).toContain('<dimension ref="A1:N5"/>');
  });

  it('leaves the declared range alone when every patched cell is inside it', () => {
    const opened = openWorkbook(sheet('A1:B2'));
    const session = new DocumentSession(opened.workbook);
    const model = opened.workbook.sheets[0]!;
    editCell(session, model.id, 1, 1, numberValue(9));

    const result = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
    });
    const xml = unzipPart(result.bytes, 'xl/worksheets/sheet1.xml');
    expect(xml).toContain('<c r="B2"><v>9</v></c>');
    expect(xml).toContain('<dimension ref="A1:B2"/>');
  });
});

describe('property: edited saves are readable and no-op saves are byte-identical', () => {
  it('round-trips random single-cell edits', () => {
    const original = fixture('m1-chart-unknown.xlsx');
    fc.assert(
      fc.property(
        fc.integer({ min: 1, max: 3 }),
        fc.integer({ min: 0, max: 2 }),
        fc.oneof(fc.integer({ min: -1_000_000, max: 1_000_000 }), fc.string({ maxLength: 40 })),
        (row, col, value) => {
          const opened = openWorkbook(original);
          const session = new DocumentSession(opened.workbook);
          const sheet = findSheet(opened.workbook, opened.initialSheetId)!;
          editCell(
            session,
            sheet.id,
            row,
            col,
            typeof value === 'number' ? numberValue(value) : textValue(value),
          );
          const result = saveWorkbook({
            workbook: opened.workbook,
            package: opened.package,
            dirtyCells: session.dirtyCellMap(),
          });
          const reopened = openWorkbook(result.bytes);
          const cell = reopened.workbook.sheets[0]!.cells.get(`${row}:${col}`);
          if (typeof value === 'number') {
            expect(cell?.value).toEqual({ type: 'number', value });
          } else {
            expect(cell?.value).toEqual({ type: 'text', value });
          }
        },
      ),
      { numRuns: 20 },
    );
  });
});
