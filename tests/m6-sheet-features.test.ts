import { describe, expect, it } from 'vitest';
import { DocumentSession } from '@commands/session';
import { openWorkbook, type OpenedWorkbook } from '@ooxml/open';
import { saveWorkbook, adoptSavedPackage, type SaveWorkbookResult } from '@ooxml/save';
import { printAreaOf } from '@ooxml/print-setup-xml';
import type { Sheet } from '@model/types';
import { DEFAULT_ALIGNMENT, DEFAULT_BORDER, DEFAULT_FILL, DEFAULT_FONT } from '@model/style';
import { numberValue, textValue } from '@model/value';
import { createSheetProjection } from '@view/projection';
import { sheetVisibilityById } from '@view/sheet-visibility';
import type { ProjectedStyle } from '@view/types';
import { cfPaintOps } from '@grid/cf-paint';
import { CanvasRenderFrame, COMMENT_INDICATOR_COLOR } from '@grid/canvas-renderer';
import {
  assertByteIdentical,
  assertDirtyPartPreserved,
  assertOnlyPartsChanged,
  assertRelationshipsResolve,
  assertSemanticDiff,
  fixtureBytes,
  openFixture,
  partText,
  saveNoOp,
} from './helpers/roundtrip';

/**
 * M6 lane 3 — sheet and content features (file-compat §9 F12-F14, F16, F17,
 * F23). Fixtures come from scripts/fixtures/groups/sheet-features.mjs; the
 * identity, dirty-part and semantic measurements are the independent reader in
 * tests/helpers/roundtrip.ts.
 */

const WORKBOOK_PART = 'xl/workbook.xml';
const SHEET1 = 'xl/worksheets/sheet1.xml';

function sheetOf(opened: OpenedWorkbook): Sheet {
  return opened.workbook.sheets[0]!;
}

function saveSession(opened: OpenedWorkbook, session: DocumentSession): SaveWorkbookResult {
  return saveWorkbook({
    workbook: opened.workbook,
    package: opened.package,
    dirtyCells: session.dirtyCellMap(),
    dirtyDimensions: session.dirtyDimensionSnapshot(),
    dirtyMerges: session.dirtyMergeSheets(),
  });
}

function editCell(
  session: DocumentSession,
  sheetId: string,
  row: number,
  col: number,
  value: ReturnType<typeof textValue>,
): void {
  session.dispatch('cell.set', { sheetId, row, col, value });
}

// ---------------------------------------------------------------------------
// F12 threaded-comments.xlsx
// ---------------------------------------------------------------------------

describe('F12 threaded-comments.xlsx — threaded comments and persons preserved; indicator renders', () => {
  const original = fixtureBytes('threaded-comments.xlsx');

  it('reads the legacy note, the thread, the reply and the persons', () => {
    const opened = openFixture('threaded-comments.xlsx');
    const sheet = sheetOf(opened);
    expect(sheet.name).toBe('Discussions');

    // D2 note, D3 thread root, D4 reply, E2 thread on an empty cell.
    expect([...sheet.comments.keys()].sort()).toEqual(['1:3', '1:4', '2:3', '3:3']);
    expect(sheet.comments.get('1:3')).toEqual({
      text: 'Legacy note: confirm the budget.',
      author: 'Ana Silva',
      threaded: false,
    });
    expect(sheet.comments.get('2:3')).toEqual({
      text: 'Threaded: two candidates shortlisted.',
      author: 'Ana Silva',
      threaded: true,
    });
    expect(sheet.comments.get('3:3')).toEqual({
      text: 'Reply: the launch date is blocked.',
      author: 'Bo Chen',
      threaded: true,
    });
    expect(sheet.comments.get('1:4')).toEqual({
      text: 'Follow-up on an empty cell.',
      author: 'Bo Chen',
      threaded: true,
    });
    // A comment may be anchored to a cell the grid has no value for.
    expect(sheet.cells.has('1:4')).toBe(false);

    // The comment parts and the persons part are opaque to the model.
    for (const part of [
      'xl/comments1.xml',
      'xl/threadedComments/threadedComment1.xml',
      'xl/persons/person.xml',
    ]) {
      expect(opened.workbook.opaqueParts.has(part)).toBe(true);
    }
  });

  it('keeps every comment part byte-identical through a no-op save and an unrelated edit', () => {
    const opened = openFixture('threaded-comments.xlsx');
    const noOp = saveNoOp(opened);
    expect(noOp.noOp).toBe(true);
    assertByteIdentical(original, noOp.bytes, 'threaded-comments no-op save');

    const session = new DocumentSession(opened.workbook);
    editCell(session, sheetOf(opened).id, 0, 4, textValue('edited'));
    const saved = saveSession(opened, session);
    assertRelationshipsResolve(saved.bytes);
    assertOnlyPartsChanged(original, saved.bytes, [SHEET1]);
    for (const part of [
      'xl/comments1.xml',
      'xl/threadedComments/threadedComment1.xml',
      'xl/persons/person.xml',
      'xl/drawings/vmlDrawing1.vml',
    ]) {
      expect(partText(saved.bytes, part)).toBe(partText(original, part));
    }
    const beforeXml = partText(original, SHEET1);
    const afterXml = partText(saved.bytes, SHEET1);
    assertDirtyPartPreserved(beforeXml, afterXml, {
      changedCells: ['E1'],
      changedRows: [1],
    });
    expect(afterXml).toContain('<legacyDrawing r:id="rId2"/>');
    // The model still holds the comments after the save.
    expect(openWorkbook(saved.bytes).workbook.sheets[0]!.comments.size).toBe(4);
  });

  it('exposes the indicator cells to the projection and paints one marker per comment', () => {
    const opened = openFixture('threaded-comments.xlsx');
    const projection = createSheetProjection(opened.workbook, sheetOf(opened).id, { version: 1 });
    expect([...(projection.commentCells ?? [])].sort()).toEqual(['1:3', '1:4', '2:3', '3:3']);

    const canvas = recordingCanvas();
    const frame = new CanvasRenderFrame(canvas.element);
    frame.resize(320, 160, 1);
    frame.render({
      projection,
      selection: {
        anchor: { row: 0, col: 0 },
        focus: { row: 0, col: 0 },
        extent: { kind: 'cells' },
      },
      scrollX: 0,
      scrollY: 0,
      width: 320,
      height: 160,
      dpr: 1,
    });
    const markers = canvas.fills.filter((fill) => fill.color === COMMENT_INDICATOR_COLOR);
    expect(markers).toHaveLength(4);
    for (const marker of markers) {
      // A top-right corner triangle: three points, closing back to the edge.
      expect(marker.points).toHaveLength(3);
      const [[leftX], [topX], [bottomX]] = marker.points as [
        [number, number],
        [number, number],
        [number, number],
      ];
      expect(topX - leftX).toBeGreaterThan(0);
      expect(bottomX).toBe(topX);
    }
  });

  it('moves the comment positions with a row insert and keeps the persons part', () => {
    const opened = openFixture('threaded-comments.xlsx');
    const session = new DocumentSession(opened.workbook);
    const sheet = sheetOf(opened);
    session.dispatch('structure.insertRows', {
      sheetId: sheet.id,
      range: { startRow: 0, startCol: 0, endRow: 0, endCol: 4 },
    });
    expect([...sheet.comments.keys()].sort()).toEqual(['2:3', '2:4', '3:3', '4:3']);

    const saved = saveSession(opened, session);
    assertRelationshipsResolve(saved.bytes);
    const comments = partText(saved.bytes, 'xl/comments1.xml');
    expect(comments).toContain('<comment ref="D3" authorId="0">');
    const threaded = partText(saved.bytes, 'xl/threadedComments/threadedComment1.xml');
    for (const ref of ['D4', 'E3', 'D5']) expect(threaded).toContain(`ref="${ref}"`);
    expect(partText(saved.bytes, 'xl/persons/person.xml')).toBe(
      partText(original, 'xl/persons/person.xml'),
    );
    // The legacy note's VML anchor follows the row too.
    expect(partText(saved.bytes, 'xl/drawings/vmlDrawing1.vml')).toContain('<x:Row>2</x:Row>');
  });
});

// ---------------------------------------------------------------------------
// F13 validation-cf.xlsx
// ---------------------------------------------------------------------------

describe('F13 validation-cf.xlsx — validation and every supported CF rule render and round-trip', () => {
  const original = fixtureBytes('validation-cf.xlsx');

  it('parses every supported validation type and both hyperlink kinds', () => {
    const opened = openFixture('validation-cf.xlsx');
    const sheet = sheetOf(opened);

    expect(sheet.validation).toHaveLength(10);
    const types = sheet.validation.map((record) => /type="([^"]*)"/.exec(record.rawXml)?.[1] ?? '');
    expect(new Set(types)).toEqual(
      new Set(['whole', 'decimal', 'list', 'date', 'time', 'textLength', 'custom', 'none']),
    );
    // The union `sqref="A8 B8"` record covers the bounding rectangle.
    const union = sheet.validation.find((record) => record.rawXml.includes('notBetween'));
    expect(union?.range).toEqual({
      startRow: 7,
      startCol: 0,
      endRow: 7,
      endCol: 1,
    });
    // The list and custom formulas stay authored text.
    expect(sheet.validation.some((record) => record.rawXml.includes('"Alpha,Beta,Gamma"'))).toBe(
      true,
    );
    expect(sheet.validation.some((record) => record.rawXml.includes('AND(ISNUMBER(J2)'))).toBe(
      true,
    );

    expect(sheet.hyperlinks.get('7:0')).toEqual({
      target: 'https://example.invalid/report',
      tooltip: 'External report',
    });
    expect(sheet.hyperlinks.get('7:1')).toEqual({
      target: '',
      location: 'Rules!A1',
    });
  });

  it('round-trips the validation, CF and hyperlink elements byte-for-byte', () => {
    const opened = openFixture('validation-cf.xlsx');
    const noOp = saveNoOp(opened);
    expect(noOp.noOp).toBe(true);
    assertByteIdentical(original, noOp.bytes, 'validation-cf no-op save');

    const session = new DocumentSession(opened.workbook);
    editCell(session, sheetOf(opened).id, 0, 10, textValue('note'));
    const saved = saveSession(opened, session);
    assertRelationshipsResolve(saved.bytes);
    assertOnlyPartsChanged(original, saved.bytes, [SHEET1]);
    const beforeXml = partText(original, SHEET1);
    const afterXml = partText(saved.bytes, SHEET1);
    assertDirtyPartPreserved(beforeXml, afterXml, {
      changedCells: ['K1'],
      changedRows: [1],
    });
    expect(afterXml).toContain(
      beforeXml.slice(
        beforeXml.indexOf('<conditionalFormatting'),
        beforeXml.indexOf('<dataValidations'),
      ),
    );
    expect(afterXml).toContain(
      beforeXml.slice(beforeXml.indexOf('<dataValidations'), beforeXml.indexOf('<hyperlinks>')),
    );
    expect(afterXml).toContain(
      beforeXml.slice(beforeXml.indexOf('<hyperlinks>'), beforeXml.indexOf('<pageMargins')),
    );
    assertSemanticDiff(original, saved.bytes, ['sheetParts.Rules.cells.length']);
  });

  it('evaluates every supported rule type and produces paint operations for them', () => {
    const opened = openFixture('validation-cf.xlsx');
    const sheet = sheetOf(opened);
    const rules = opened.conditionalFormatRules.get(sheet.id) ?? [];
    const kinds = rules.map((rule) => rule.kind);
    for (const kind of [
      'cellIs',
      'text',
      'timePeriod',
      'top10',
      'duplicate',
      'unique',
      'colorScale',
      'dataBar',
      'iconSet',
    ]) {
      expect(kinds).toContain(kind);
    }
    // An expression rule is preserved and reported, never painted.
    expect(kinds).toContain('unsupported');
    expect(rules.find((rule) => rule.kind === 'unsupported')).toMatchObject({
      reason: 'formula',
    });

    const projection = createSheetProjection(opened.workbook, sheet.id, {
      version: 1,
      conditionalFormats: {
        rules,
        today: { year: 2026, month: 11, day: 19 },
      },
    });
    const formatting = projection.conditionalFormatting;
    expect(formatting).not.toBeNull();

    // B3 (row 2, col 1) is 150: the cellIs rule uses dxf 0, a red font.
    const cellIs = formatting!.evaluate(2, 1);
    expect(cellIs.ruleIds).toContain(rules.find((rule) => rule.kind === 'cellIs')!.id);
    expect(cellIs.overlay.font?.color).toBe('#9C0006');

    // A3 is "normal": the containsText rule is false; A4 is "URGENT issue".
    expect(formatting!.evaluate(2, 0).ruleIds).not.toContain(
      rules.find((rule) => rule.kind === 'text')!.id,
    );
    expect(formatting!.evaluate(3, 0).ruleIds).toContain(
      rules.find((rule) => rule.kind === 'text')!.id,
    );

    // F3 is today (2026-11-19): the timePeriod rule is true.
    expect(formatting!.evaluate(2, 5).ruleIds).toContain(
      rules.find((rule) => rule.kind === 'timePeriod')!.id,
    );

    // G4 (90) sits on the colour scale; H4 (80) carries a data bar; I4 (80)
    // carries an icon.
    expect(formatting!.evaluate(3, 6).colorScale).not.toBeNull();
    expect(formatting!.evaluate(3, 7).dataBar?.kind).toBe('dataBar');
    expect(formatting!.evaluate(3, 8).iconSet?.kind).toBe('iconSet');

    // The duplicate rule formats D2 (7) with dxf 2, a green fill.
    const style: ProjectedStyle = {
      font: DEFAULT_FONT,
      fill: DEFAULT_FILL,
      border: DEFAULT_BORDER,
      alignment: DEFAULT_ALIGNMENT,
      numberFormat: null,
      numberFormatColor: null,
    };
    const duplicate = formatting!.evaluate(1, 3);
    expect(duplicate.ruleIds).toContain(rules.find((rule) => rule.kind === 'duplicate')!.id);
    const ops = cfPaintOps({
      evaluation: duplicate,
      rect: { x: 0, y: 0, width: 64, height: 20 },
      zoom: 1,
      style,
      displayText: '7',
    });
    expect(ops.some((op) => op.kind === 'fill')).toBe(true);
  });

  it('transforms validation, CF and hyperlink references through a structural edit', () => {
    const opened = openFixture('validation-cf.xlsx');
    const sheet = sheetOf(opened);
    const session = new DocumentSession(opened.workbook);
    const result = session.dispatch('structure.insertRows', {
      sheetId: sheet.id,
      range: { startRow: 0, startCol: 0, endRow: 0, endCol: 9 },
    });
    expect(result.changes.length).toBeGreaterThan(0);

    // Model side: the first validation moves from B2:B6 to B3:B7 and the
    // hyperlinks move from row 8/9 to row 9/10.
    expect(sheet.validation[0]?.range).toEqual({
      startRow: 2,
      startCol: 1,
      endRow: 6,
      endCol: 1,
    });
    expect(sheet.hyperlinks.has('8:0')).toBe(true);
    expect(sheet.hyperlinks.has('8:1')).toBe(true);
    expect(sheet.conditionalFormats[0]?.range.startRow).toBe(2);

    const saved = saveSession(opened, session);
    assertRelationshipsResolve(saved.bytes);
    const afterXml = partText(saved.bytes, SHEET1);
    expect(afterXml).toContain('<dataValidation type="whole" operator="between"');
    expect(afterXml).toContain('sqref="B3:B7"');
    expect(afterXml).toContain('<conditionalFormatting sqref="B3:B7">');
    expect(afterXml).toContain('<hyperlink ref="A9" r:id="rId1" tooltip="External report"/>');
    // A same-workbook `location` is authored text, not a range anchor; it is
    // preserved as written and its transform is a follow-up (see lane report).
    expect(afterXml).toContain('<hyperlink ref="B9" location="Rules!A1" display="Rules sheet"/>');
  });
});

// ---------------------------------------------------------------------------
// F14 protection-hidden.xlsx
// ---------------------------------------------------------------------------

describe('F14 protection-hidden.xlsx — protection and hidden state preserved, not enforced', () => {
  const original = fixtureBytes('protection-hidden.xlsx');

  it('reads workbook and sheet protection, sheet visibility and hidden dimensions', () => {
    const opened = openFixture('protection-hidden.xlsx');
    const workbook = opened.workbook;

    expect(workbook.protection?.rawXml).toContain('workbookPassword="CC1A"');
    expect(workbook.protection?.rawXml).toContain('lockStructure="1"');
    const board = workbook.sheets[0]!;
    expect(board.name).toBe('Board');
    expect(board.protection?.rawXml).toContain('sheet="1"');
    expect(board.protection?.rawXml).toContain('password="CC1A"');

    expect(board.rowDimensions.get(1)?.hidden).toBe(true);
    expect(board.columnDimensions.get(3)?.hidden).toBe(true);

    const visibility = sheetVisibilityById(workbook);
    expect(visibility.get(workbook.sheets[0]!.id)).toBe('visible');
    expect(visibility.get(workbook.sheets[1]!.id)).toBe('hidden');
    expect(visibility.get(workbook.sheets[2]!.id)).toBe('veryHidden');
  });

  it('is byte-identical on a no-op save', () => {
    const opened = openFixture('protection-hidden.xlsx');
    const saved = saveNoOp(opened);
    expect(saved.noOp).toBe(true);
    assertByteIdentical(original, saved.bytes, 'protection-hidden no-op save');
    assertRelationshipsResolve(saved.bytes);
  });

  it('lets an edit through a protected sheet and keeps every protection and visibility byte', () => {
    const opened = openFixture('protection-hidden.xlsx');
    const session = new DocumentSession(opened.workbook);
    const board = sheetOf(opened);
    // Protection is deliberately not enforced (feature matrix §11): the edit
    // lands even though the sheet is protected.
    editCell(session, board.id, 2, 0, numberValue(30));
    const saved = saveSession(opened, session);
    assertRelationshipsResolve(saved.bytes);
    assertOnlyPartsChanged(original, saved.bytes, [SHEET1]);

    // The workbook part (workbookProtection + sheet states) is untouched.
    expect(partText(saved.bytes, WORKBOOK_PART)).toBe(partText(original, WORKBOOK_PART));
    const beforeXml = partText(original, SHEET1);
    const afterXml = partText(saved.bytes, SHEET1);
    assertDirtyPartPreserved(beforeXml, afterXml, {
      changedCells: ['A3'],
      changedRows: [3],
    });
    expect(afterXml).toContain(
      '<sheetProtection sheet="1" objects="1" scenarios="1" formatCells="0" formatColumns="1" password="CC1A"/>',
    );
    expect(afterXml).toContain('<row r="2" hidden="1">');
    expect(afterXml).toContain('<col min="4" max="4" width="9" hidden="1" customWidth="1"/>');

    const reopened = openWorkbook(saved.bytes);
    expect(reopened.workbook.protection?.rawXml).toContain('lockStructure="1"');
    const visibility = sheetVisibilityById(reopened.workbook);
    expect(visibility.get(reopened.workbook.sheets[1]!.id)).toBe('hidden');
    expect(visibility.get(reopened.workbook.sheets[2]!.id)).toBe('veryHidden');
    expect(reopened.workbook.sheets[0]!.cells.get('2:0')?.value).toEqual(numberValue(30));
  });

  it('keeps workbook protection and every sheet state through a sheet-list rewrite', () => {
    const opened = openFixture('protection-hidden.xlsx');
    const session = new DocumentSession(opened.workbook);
    const board = sheetOf(opened);
    session.dispatch('sheet.rename', { sheetId: board.id, name: 'Board 2026' });
    const saved = saveSession(opened, session);
    assertRelationshipsResolve(saved.bytes);
    assertOnlyPartsChanged(original, saved.bytes, [WORKBOOK_PART]);

    const workbookXml = partText(saved.bytes, WORKBOOK_PART);
    expect(workbookXml).toContain(
      '<workbookProtection workbookPassword="CC1A" lockStructure="1" lockWindows="1"/>',
    );
    expect(workbookXml).toContain('<sheet name="Board 2026" sheetId="1" r:id="rId1"/>');
    expect(workbookXml).toContain('<sheet name="Members" sheetId="2" state="hidden" r:id="rId2"/>');
    expect(workbookXml).toContain(
      '<sheet name="Archive" sheetId="3" state="veryHidden" r:id="rId3"/>',
    );
  });
});

// ---------------------------------------------------------------------------
// F16 print-setup.xlsx
// ---------------------------------------------------------------------------

describe('F16 print-setup.xlsx — print area, orientation, fit-to-width and margins', () => {
  const original = fixtureBytes('print-setup.xlsx');

  it('parses the model values and the sheet-scoped print area name', () => {
    const opened = openFixture('print-setup.xlsx');
    const sheet = sheetOf(opened);
    expect(sheet.printSetup).toMatchObject({
      orientation: 'landscape',
      fitToWidth: 1,
      fitToHeight: 0,
      paperSize: 9,
      margins: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
    });
    expect(printAreaOf(sheet)).toEqual([{ startRow: 0, startCol: 0, endRow: 19, endCol: 3 }]);
    expect(sheet.names.map((name) => name.name)).toEqual(['_xlnm.Print_Area']);
  });

  it('is byte-identical on a no-op save and keeps the elements through a cell edit', () => {
    const opened = openFixture('print-setup.xlsx');
    const noOp = saveNoOp(opened);
    expect(noOp.noOp).toBe(true);
    assertByteIdentical(original, noOp.bytes, 'print-setup no-op save');

    const session = new DocumentSession(opened.workbook);
    editCell(session, sheetOf(opened).id, 1, 0, textValue('N'));
    const saved = saveSession(opened, session);
    assertRelationshipsResolve(saved.bytes);
    assertOnlyPartsChanged(original, saved.bytes, [SHEET1]);
    const beforeXml = partText(original, SHEET1);
    const afterXml = partText(saved.bytes, SHEET1);
    assertDirtyPartPreserved(beforeXml, afterXml, {
      changedCells: ['A2'],
      changedRows: [2],
    });
    for (const element of ['<printOptions', '<pageMargins', '<pageSetup', '<sheetPr>']) {
      expect(afterXml).toContain(
        beforeXml.slice(
          beforeXml.indexOf(element),
          beforeXml.indexOf('/>', beforeXml.indexOf(element)) + 2,
        ),
      );
    }
    // The print area name survives in the workbook part.
    expect(partText(saved.bytes, WORKBOOK_PART)).toContain(
      `<definedName name="_xlnm.Print_Area" localSheetId="0">'Report'!$A$1:$D$20</definedName>`,
    );
  });

  it('writes a changed orientation, fit-to-width and margins without touching the rest', () => {
    const opened = openFixture('print-setup.xlsx');
    const sheet = sheetOf(opened);
    const setup = sheet.printSetup!;
    sheet.printSetup = {
      ...setup,
      orientation: 'portrait',
      fitToWidth: 2,
      margins: { ...setup.margins!, left: 0.5 },
    };
    const saved = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: new Map(),
    });
    assertRelationshipsResolve(saved.bytes);
    expect(saved.noOp).toBe(false);
    assertOnlyPartsChanged(original, saved.bytes, [SHEET1]);

    const afterXml = partText(saved.bytes, SHEET1);
    expect(afterXml).toContain(
      '<pageSetup paperSize="9" orientation="portrait" fitToWidth="2" fitToHeight="0" scale="85" firstPageNumber="1" useFirstPageNumber="1" cellComments="none"/>',
    );
    expect(afterXml).toContain(
      '<pageMargins left="0.5" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>',
    );
    expect(afterXml).toContain('<printOptions horizontalCentered="1" verticalCentered="0"/>');
    expect(afterXml).toContain('<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>');

    const reopened = openWorkbook(saved.bytes);
    expect(reopened.workbook.sheets[0]!.printSetup).toMatchObject({
      orientation: 'portrait',
      fitToWidth: 2,
      fitToHeight: 0,
      paperSize: 9,
      margins: { left: 0.5, right: 0.7 },
    });
    // Honouring these values in print preview and PDF is M7 work.

    // A second save with the same model does not rewrite the part again.
    adoptSavedPackage(opened.workbook, saved.package);
    const second = saveWorkbook({
      workbook: opened.workbook,
      package: saved.package,
      dirtyCells: new Map(),
    });
    assertByteIdentical(saved.bytes, second.bytes, 'second print-setup save');
  });

  it('transforms the print area name through a row insert', () => {
    const opened = openFixture('print-setup.xlsx');
    const sheet = sheetOf(opened);
    const session = new DocumentSession(opened.workbook);
    session.dispatch('structure.insertRows', {
      sheetId: sheet.id,
      range: { startRow: 0, startCol: 0, endRow: 0, endCol: 3 },
    });
    expect(sheet.names[0]?.formula).toBe("'Report'!$A$2:$D$21");
    expect(printAreaOf(sheet)).toEqual([{ startRow: 1, startCol: 0, endRow: 20, endCol: 3 }]);

    const saved = saveSession(opened, session);
    expect(partText(saved.bytes, WORKBOOK_PART)).toContain(
      `<definedName name="_xlnm.Print_Area" localSheetId="0">'Report'!$A$2:$D$21</definedName>`,
    );
    expect(openWorkbook(saved.bytes).workbook.sheets[0]!.names[0]?.formula).toBe(
      "'Report'!$A$2:$D$21",
    );
  });
});

// ---------------------------------------------------------------------------
// F17 rich-text.xlsx
// ---------------------------------------------------------------------------

describe('F17 rich-text.xlsx — untouched runs survive; editing warns and flattens', () => {
  const original = fixtureBytes('rich-text.xlsx');

  it('reads both the shared-string and inline runs as opaque payloads', () => {
    const opened = openFixture('rich-text.xlsx');
    const sheet = sheetOf(opened);
    const a1 = sheet.cells.get('0:0')!;
    expect(a1.value).toEqual({ type: 'text', value: 'Bold lead plain tail' });
    expect(a1.richText?.plainText).toBe('Bold lead plain tail');
    expect(a1.richText?.xml).toContain('<rPr><b/>');
    const a2 = sheet.cells.get('1:0')!;
    expect(a2.richText?.plainText).toBe('Rich inline');
    expect(a2.richText?.xml).toContain('<r>');
    // A plain cell has no payload.
    expect(sheet.cells.get('0:1')!.richText).toBeNull();
  });

  it('survives a no-op save byte-identically, runs included', () => {
    const opened = openFixture('rich-text.xlsx');
    const noOp = saveNoOp(opened);
    expect(noOp.noOp).toBe(true);
    assertByteIdentical(original, noOp.bytes, 'rich-text no-op save');
    expect(partText(noOp.bytes, 'xl/sharedStrings.xml')).toContain('<rPr><b/>');
  });

  it('warns and flattens in one transaction, then writes plain text', () => {
    const opened = openFixture('rich-text.xlsx');
    const session = new DocumentSession(opened.workbook);
    const sheet = sheetOf(opened);

    const result = session.dispatch('cell.set', {
      sheetId: sheet.id,
      row: 0,
      col: 0,
      value: textValue('Bold lead plain tail'),
    });
    expect(result.changes).toHaveLength(1);
    expect(result.ui).toMatchObject({ type: 'toast' });
    expect(result.ui?.message).toMatch(/rich text/i);
    expect(result.ui?.message).toMatch(/plain text/i);
    expect(session.canUndo).toBe(true);
    expect(sheet.cells.get('0:0')!.richText).toBeNull();
    expect(sheet.cells.get('0:0')!.value).toEqual({
      type: 'text',
      value: 'Bold lead plain tail',
    });
    // The inline rich cell is untouched.
    expect(sheet.cells.get('1:0')!.richText).not.toBeNull();

    const saved = saveSession(opened, session);
    assertRelationshipsResolve(saved.bytes);
    expect(saved.noOp).toBe(false);
    assertOnlyPartsChanged(original, saved.bytes, [SHEET1, 'xl/sharedStrings.xml']);
    const beforeXml = partText(original, SHEET1);
    const afterXml = partText(saved.bytes, SHEET1);
    expect(afterXml).toContain('<c r="A1" t="s"><v>2</v></c>');
    assertDirtyPartPreserved(beforeXml, afterXml, { changedCells: ['A1'], changedRows: [1] });
    // The inline rich cell keeps its runs byte-for-byte.
    expect(afterXml).toContain(
      '<c r="A2" t="inlineStr"><is><r><rPr><b/><sz val="11"/><rFont val="Calibri"/></rPr><t>Rich</t></r><r><t xml:space="preserve"> inline</t></r></is></c>',
    );
    // The shared-string table is append-only: the rich entry survives and the
    // flattened text is appended after it.
    const strings = partText(saved.bytes, 'xl/sharedStrings.xml');
    expect(strings).toContain('<rPr><b/>');
    expect(strings).toContain('<si><t xml:space="preserve">Bold lead plain tail</t></si>');
    expect(strings).toContain('uniqueCount="3"');

    const reopened = openWorkbook(saved.bytes);
    expect(reopened.workbook.sheets[0]!.cells.get('0:0')!.richText).toBeNull();
    expect(reopened.workbook.sheets[0]!.cells.get('1:0')!.richText).not.toBeNull();
  });

  it('lets Clear contents drop the runs deliberately, keeping the cell style', () => {
    const opened = openFixture('rich-text.xlsx');
    const session = new DocumentSession(opened.workbook);
    const sheet = sheetOf(opened);
    session.dispatch('edit.clearContents', {
      sheetId: sheet.id,
      range: { startRow: 1, startCol: 2, endRow: 1, endCol: 2 },
    });
    const cleared = sheet.cells.get('1:2')!;
    expect(cleared.value).toEqual({ type: 'blank' });
    expect(cleared.richText).toBeNull();
    expect(cleared.styleId).toBe(1);
    expect(cleared.richTextFlattened).toBe(true);

    const saved = saveSession(opened, session);
    assertRelationshipsResolve(saved.bytes);
    const afterXml = partText(saved.bytes, SHEET1);
    expect(afterXml).toContain('<c r="C2" s="1"/>');
    expect(openWorkbook(saved.bytes).workbook.sheets[0]!.cells.get('1:2')?.richText).toBeNull();
  });

  it('restores the exact prior state on undo, and the save keeps the runs', () => {
    const opened = openFixture('rich-text.xlsx');
    const session = new DocumentSession(opened.workbook);
    const sheet = sheetOf(opened);

    session.dispatch('cell.set', {
      sheetId: sheet.id,
      row: 0,
      col: 0,
      value: textValue('plain now'),
    });
    session.undo();

    const restored = sheet.cells.get('0:0')!;
    expect(restored.richText?.xml).toContain('<rPr><b/>');
    expect(restored.value).toEqual({ type: 'text', value: 'Bold lead plain tail' });

    const saved = saveSession(opened, session);
    assertRelationshipsResolve(saved.bytes);
    expect(partText(saved.bytes, SHEET1)).toBe(partText(original, SHEET1));
    expect(partText(saved.bytes, 'xl/sharedStrings.xml')).toBe(
      partText(original, 'xl/sharedStrings.xml'),
    );
    assertOnlyPartsChanged(original, saved.bytes, []);
  });
});

// ---------------------------------------------------------------------------
// F23 names-defined.xlsx
// ---------------------------------------------------------------------------

describe('F23 names-defined.xlsx — workbook and sheet names, quoted references, transforms', () => {
  const original = fixtureBytes('names-defined.xlsx');

  it('reads workbook and sheet scope with the quoted sheet references authored', () => {
    const opened = openFixture('names-defined.xlsx');
    const workbook = opened.workbook;
    expect(workbook.sheets.map((sheet) => sheet.name)).toEqual([
      'Sales Data',
      'Summary',
      "O'Brien's Log",
    ]);
    expect(workbook.names).toEqual([
      { name: 'TotalSales', formula: "'Sales Data'!$B$2:$B$5", sheetId: null },
      { name: 'QuotedApostrophe', formula: "'O''Brien''s Log'!$A$1:$A$2", sheetId: null },
    ]);
    expect(workbook.sheets[0]!.names).toEqual([
      { name: 'TaxRate', formula: '0.2', sheetId: workbook.sheets[0]!.id },
      { name: 'LocalFigure', formula: '$B$2', sheetId: workbook.sheets[0]!.id },
    ]);
    expect(workbook.sheets[1]!.names).toEqual([
      { name: 'BoldTitle', formula: "'Summary'!$A$1", sheetId: workbook.sheets[1]!.id },
    ]);
  });

  it('is byte-identical on a no-op save', () => {
    const opened = openFixture('names-defined.xlsx');
    const saved = saveNoOp(opened);
    expect(saved.noOp).toBe(true);
    assertByteIdentical(original, saved.bytes, 'names-defined no-op save');
  });

  it('transforms workbook and sheet names on a row insert and keeps the others byte-identical', () => {
    const opened = openFixture('names-defined.xlsx');
    const workbook = opened.workbook;
    const sales = workbook.sheets[0]!;
    const session = new DocumentSession(workbook);
    session.dispatch('structure.insertRows', {
      sheetId: sales.id,
      range: { startRow: 1, startCol: 0, endRow: 1, endCol: 3 },
    });

    expect(workbook.names.find((name) => name.name === 'TotalSales')?.formula).toBe(
      "'Sales Data'!$B$3:$B$6",
    );
    expect(sales.names.find((name) => name.name === 'LocalFigure')?.formula).toBe('$B$3');
    // A constant name and a name on another sheet do not move.
    expect(sales.names.find((name) => name.name === 'TaxRate')?.formula).toBe('0.2');
    expect(workbook.sheets[1]!.names[0]?.formula).toBe("'Summary'!$A$1");
    expect(workbook.names.find((name) => name.name === 'QuotedApostrophe')?.formula).toBe(
      "'O''Brien''s Log'!$A$1:$A$2",
    );
    // The cross-sheet formula in Summary follows the same edit.
    expect(workbook.sheets[1]!.cells.get('1:1')?.formula).toBe("SUM('Sales Data'!$B$3:$B$6)");

    const saved = saveSession(opened, session);
    assertRelationshipsResolve(saved.bytes);
    assertOnlyPartsChanged(original, saved.bytes, [
      WORKBOOK_PART,
      SHEET1,
      'xl/worksheets/sheet2.xml',
    ]);
    const workbookXml = partText(saved.bytes, WORKBOOK_PART);
    expect(workbookXml).toContain(
      `<definedName name="TotalSales">'Sales Data'!$B$3:$B$6</definedName>`,
    );
    expect(workbookXml).toContain(
      `<definedName name="LocalFigure" localSheetId="0">$B$3</definedName>`,
    );
    expect(workbookXml).toContain(
      `<definedName name="TaxRate" localSheetId="0" hidden="1">0.2</definedName>`,
    );
    expect(workbookXml).toContain(
      `<definedName name="BoldTitle" localSheetId="1">'Summary'!$A$1</definedName>`,
    );
    expect(workbookXml).toContain(
      `<definedName name="QuotedApostrophe">'O''Brien''s Log'!$A$1:$A$2</definedName>`,
    );
    const summaryXml = partText(saved.bytes, 'xl/worksheets/sheet2.xml');
    expect(summaryXml).toContain('<f>SUM(&apos;Sales Data&apos;!$B$3:$B$6)</f>');

    const reopened = openWorkbook(saved.bytes);
    expect(reopened.workbook.names.find((name) => name.name === 'TotalSales')?.formula).toBe(
      "'Sales Data'!$B$3:$B$6",
    );
  });

  it('collapses a name to #REF! when its only cell is deleted', () => {
    const opened = openFixture('names-defined.xlsx');
    const workbook = opened.workbook;
    const sales = workbook.sheets[0]!;
    const session = new DocumentSession(workbook);
    session.dispatch('structure.deleteRows', {
      sheetId: sales.id,
      range: { startRow: 1, startCol: 0, endRow: 1, endCol: 3 },
    });
    expect(sales.names.find((name) => name.name === 'LocalFigure')?.formula).toBe('#REF!');
    expect(workbook.names.find((name) => name.name === 'TotalSales')?.formula).toBe(
      "'Sales Data'!$B$2:$B$4",
    );

    const saved = saveSession(opened, session);
    expect(partText(saved.bytes, WORKBOOK_PART)).toContain(
      `<definedName name="LocalFigure" localSheetId="0">#REF!</definedName>`,
    );
    expect(
      openWorkbook(saved.bytes).workbook.sheets[0]!.names.find(
        (name) => name.name === 'LocalFigure',
      )?.formula,
    ).toBe('#REF!');
  });
});

// ---------------------------------------------------------------------------
// A recording canvas, for the comment-indicator paint proof.
// ---------------------------------------------------------------------------

interface RecordedFill {
  readonly color: string;
  readonly points: readonly (readonly [number, number])[];
}

function recordingCanvas(): {
  element: HTMLCanvasElement;
  fills: RecordedFill[];
} {
  const fills: RecordedFill[] = [];
  let path: [number, number][] = [];
  const noop = (): void => {};
  const context = {
    fillStyle: '',
    strokeStyle: '',
    lineWidth: 1,
    font: '',
    textAlign: '',
    textBaseline: '',
    globalAlpha: 1,
    setTransform: noop,
    clearRect: noop,
    fillRect: noop,
    strokeRect: noop,
    save: noop,
    restore: noop,
    clip: noop,
    scale: noop,
    translate: noop,
    setLineDash: noop,
    arc: noop,
    ellipse: noop,
    quadraticCurveTo: noop,
    bezierCurveTo: noop,
    fillText: noop,
    stroke: noop,
    strokeText: noop,
    beginPath(): void {
      path = [];
    },
    rect: noop,
    moveTo(x: number, y: number): void {
      path.push([x, y]);
    },
    lineTo(x: number, y: number): void {
      path.push([x, y]);
    },
    closePath: noop,
    fill(): void {
      fills.push({ color: String(context.fillStyle), points: [...path] });
    },
    measureText(text: string): { width: number } {
      return { width: text.length * 7 };
    },
  };
  const element = {
    width: 0,
    height: 0,
    style: {},
    getContext: (): typeof context => context,
  };
  return { element: element as unknown as HTMLCanvasElement, fills };
}
