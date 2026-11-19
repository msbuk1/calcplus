import { describe, expect, it } from 'vitest';
import { unzipSync, zipSync } from 'fflate';
import { DocumentSession } from '@commands/session';
import { HyperFormulaEngine } from '@engine/hyperformula-engine';
import { prepareFormula } from '@engine/formula-text';
import type { EngineBatchResult, EngineCellInput, EngineWorkbookInput } from '@engine/port';
import { WorkbookSaveError } from '@ooxml/errors';
import { openWorkbook, type OpenedWorkbook } from '@ooxml/open';
import { adoptSavedPackage, saveWorkbook, type SaveWorkbookResult } from '@ooxml/save';
import { getCell, parseCellKey } from '@model/types';
import type { Cell } from '@model/types';
import { BLANK, errorValue, numberValue, textValue, type CellValue } from '@model/value';
import { formatValue } from '@view/number-format';
import {
  assertByteIdentical,
  assertDirtyPartPreserved,
  assertOnlyPartsChanged,
  assertRelationshipsResolve,
  assertSemanticDiff,
  bytesEqual,
  changedPartNames,
  diffSemantic,
  fixtureBytes,
  openFixture,
  partText,
  saveNoOp,
  semanticManifest,
} from './helpers/roundtrip';

/**
 * M6 lane 2 — write-side fidelity core (file-compat §9 F1-F8, F15).
 *
 * Fixtures come from scripts/fixtures/groups/roundtrip.mjs (F1 from
 * generate-fixtures.mjs). The identity, dirty-part and semantic-equality
 * measurements live in tests/helpers/roundtrip.ts and never use the workbook
 * model or the OOXML part parsers.
 */

// ---------------------------------------------------------------------------
// Open-phase calculation.
//
// CalculationController owns a Web Worker, so a Node test cannot construct it.
// These two helpers reproduce only the controller's load-phase glue for the
// fixtures below (which declare no defined names): the same engine inputs
// `snapshotWorkbook` builds and the same `applyCalculationResults` call it makes
// with `markDirty: false`. The engine, the screening and the session are the
// production ones.
// ---------------------------------------------------------------------------

function engineWorkbookInput(workbook: OpenedWorkbook['workbook']): EngineWorkbookInput {
  return {
    date1904: workbook.date1904,
    sheetNames: workbook.sheets.map((sheet) => sheet.name),
  };
}

function engineCellInputs(workbook: OpenedWorkbook['workbook']): EngineCellInput[] {
  const cells: EngineCellInput[] = [];
  for (const [sheetIndex, sheet] of workbook.sheets.entries()) {
    for (const [key, cell] of sheet.cells) {
      if (
        cell.value.type === 'blank' &&
        cell.formula === null &&
        cell.formulaSource === undefined
      ) {
        continue;
      }
      const { row, col } = parseCellKey(key);
      const prepared = cell.formula === null ? null : prepareFormula(cell.formula);
      const staleInput =
        cell.formulaSource !== undefined ||
        cell.richText !== null ||
        (cell.formula !== null && prepared?.text === null);
      cells.push({
        sheet: sheetIndex,
        row,
        col,
        formula: staleInput ? null : cell.formula,
        value: engineInputValue(cell, prepared?.skip ?? null),
        staleInput,
      });
    }
  }
  return cells;
}

function engineInputValue(
  cell: Cell,
  skip: 'unsupported-function' | 'unparseable' | null,
): CellValue {
  if (cell.formula === null) return cell.value;
  if (cell.savedResult !== null) return cell.savedResult;
  if (cell.formulaSource !== undefined || cell.richText !== null) return cell.value;
  return skip === 'unsupported-function' ? errorValue('#NAME?') : BLANK;
}

/** Mirror of the controller's `skippedDisplayValue` for the engine's skip list. */
function skippedDisplayValue(cell: Cell): CellValue {
  if (cell.savedResult !== null) return cell.savedResult;
  if (cell.formulaSource !== undefined || cell.richText !== null) return cell.value;
  const prepared = cell.formula === null ? null : prepareFormula(cell.formula);
  return prepared?.skip === 'unsupported-function' ? errorValue('#NAME?') : BLANK;
}

/** Run the load phase against the real HyperFormula adapter and apply results. */
function calculateOnOpen(opened: OpenedWorkbook, session: DocumentSession): EngineBatchResult {
  const engine = new HyperFormulaEngine();
  try {
    engine.start(engineWorkbookInput(opened.workbook));
    engine.load(engineCellInputs(opened.workbook));
    const result = engine.finishLoad();
    const updates = [];
    for (const cell of result.cells) {
      const sheet = opened.workbook.sheets[cell.sheet];
      if (sheet === undefined) continue;
      updates.push({
        sheetId: sheet.id,
        row: cell.row,
        col: cell.col,
        value: cell.value,
        status: cell.status,
        stale: cell.stale,
      });
    }
    for (const ref of result.skipped) {
      const sheet = opened.workbook.sheets[ref.sheet];
      if (sheet === undefined) continue;
      const modelCell = getCell(sheet, ref.row, ref.col);
      if (
        modelCell === null ||
        (modelCell.formula === null && modelCell.formulaSource === undefined)
      ) {
        continue;
      }
      updates.push({
        sheetId: sheet.id,
        row: ref.row,
        col: ref.col,
        value: skippedDisplayValue(modelCell),
        status: 'value' as const,
        stale: true,
      });
    }
    session.applyCalculationResults(updates, { markDirty: false });
    return result;
  } finally {
    engine.dispose();
  }
}

function openCalculated(bytes: Uint8Array): { opened: OpenedWorkbook; session: DocumentSession } {
  const opened = openWorkbook(bytes);
  const session = new DocumentSession(opened.workbook);
  calculateOnOpen(opened, session);
  return { opened, session };
}

function editCell(
  session: DocumentSession,
  sheetId: string,
  row: number,
  col: number,
  value: CellValue,
): void {
  session.dispatch('cell.set', { sheetId, row, col, value });
}

function saveSession(opened: OpenedWorkbook, session: DocumentSession): SaveWorkbookResult {
  return saveWorkbook({
    workbook: opened.workbook,
    package: opened.package,
    dirtyCells: session.dirtyCellMap(),
  });
}

function cellAt(sheet: { readonly cells: Map<string, Cell> }, ref: string): Cell {
  const match = /^([A-Z]+)(\d+)$/.exec(ref);
  if (match === null) throw new Error(`Bad ref ${ref}`);
  let col = 0;
  for (const character of match[1]!) col = col * 26 + (character.charCodeAt(0) - 64);
  const cell = sheet.cells.get(`${Number(match[2]) - 1}:${col - 1}`);
  if (cell === undefined) throw new Error(`Missing cell ${ref}`);
  return cell;
}

const SHEET1 = 'xl/worksheets/sheet1.xml';
const OVERVIEW = 'xl/worksheets/sheet1.xml';
const WORKBOOK_PART = 'xl/workbook.xml';

describe('F1 identity-minimal.xlsx — no-op save returns the original bytes', () => {
  const original = fixtureBytes('identity-minimal.xlsx');

  it('is byte-identical, part-identical and stable on a second save', () => {
    const opened = openWorkbook(original);
    const saved = saveNoOp(opened);
    expect(saved.noOp).toBe(true);
    assertByteIdentical(original, saved.bytes, 'identity-minimal no-op save');
    assertOnlyPartsChanged(original, saved.bytes, []);
    const second = saveWorkbook({
      workbook: opened.workbook,
      package: saved.package,
      dirtyCells: new Map(),
    });
    assertByteIdentical(saved.bytes, second.bytes, 'second identity-minimal no-op save');
  });

  it('does not mutate the retained source parts', () => {
    const opened = openWorkbook(original);
    const before = opened.workbook.sourceParts.get(SHEET1);
    saveNoOp(opened);
    expect(opened.workbook.sourceParts.get(SHEET1)).toBe(before);
  });
});

describe('F2 identity-full.xlsx — byte identity with every Tier A feature', () => {
  const original = fixtureBytes('identity-full.xlsx');

  it('opens every Tier A construct without loss', () => {
    const opened = openWorkbook(original);
    const workbook = opened.workbook;
    expect(workbook.sheets.map((sheet) => sheet.name)).toEqual(['Overview', 'Data', 'Scratch']);
    expect(workbook.date1904).toBe(false);
    expect(workbook.names.map((name) => name.name)).toEqual(['ReportTitle']);
    expect(workbook.sheets[0]!.names.map((name) => name.name)).toEqual([
      '_xlnm._FilterDatabase',
      '_xlnm.Print_Area',
    ]);
    expect(workbook.sheets[1]!.names.map((name) => name.name)).toEqual(['DataRange']);
    const overview = workbook.sheets[0]!;
    expect(overview.merges).toHaveLength(1);
    expect(overview.autoFilter?.range).toEqual({
      startRow: 9,
      startCol: 0,
      endRow: 12,
      endCol: 3,
    });
    expect(overview.conditionalFormats).toHaveLength(2);
    expect(cellAt(overview, 'B5').formula).toBe('SUM(B3:B4)');
    expect(cellAt(overview, 'B5').savedResult).toEqual(numberValue(1292.5));
    expect(cellAt(overview, 'D5').value).toEqual({ type: 'error', value: '#DIV/0!' });
    expect(cellAt(overview, 'A5').richText?.plainText).toBe('Bold lead');
    expect(cellAt(overview, 'D4').value).toEqual({ type: 'boolean', value: true });
    expect(overview.rowDimensions.get(6)?.hidden).toBe(true);
    expect(overview.columnDimensions.get(4)?.hidden).toBe(true);
    expect(overview.view).toMatchObject({
      freezeRows: 3,
      freezeCols: 1,
      showGridLines: false,
      showHeadings: true,
      zoom: 125,
    });
    expect(overview.protection).not.toBeNull();
    expect(workbook.styles.length).toBe(12);
    expect(opened.package.sharedStrings?.entries.map((entry) => entry.text)).toEqual([
      'Report',
      'Region',
      'Q1',
      'Bold lead',
      'Alpha',
      'Beta',
      ' padded ',
    ]);
    for (const part of [
      'xl/comments1.xml',
      'xl/threadedComments/threadedComment1.xml',
      'xl/persons/person.xml',
      'xl/externalLinks/externalLink1.xml',
      'xl/drawings/vmlDrawing1.vml',
    ]) {
      expect(opened.workbook.opaqueParts.has(part)).toBe(true);
    }
  });

  it('returns the original bytes when nothing changed', () => {
    const opened = openWorkbook(original);
    const saved = saveNoOp(opened);
    expect(saved.noOp).toBe(true);
    assertByteIdentical(original, saved.bytes, 'identity-full no-op save');
    assertOnlyPartsChanged(original, saved.bytes, []);
    expect(diffSemantic(semanticManifest(original), semanticManifest(saved.bytes))).toEqual([]);
  });

  it('patches only the edited worksheet when a cell changes, keeping unmodelled children', () => {
    const opened = openWorkbook(original);
    const session = new DocumentSession(opened.workbook);
    const overview = opened.workbook.sheets[0]!;
    editCell(session, overview.id, 1, 3, numberValue(5));

    const saved = saveSession(opened, session);
    assertRelationshipsResolve(saved.bytes);
    expect(saved.noOp).toBe(false);
    // fullCalcOnLoad is declared document state (file-compat §6.1), so the
    // workbook part is rewritten by one attribute; nothing else is.
    assertOnlyPartsChanged(original, saved.bytes, [WORKBOOK_PART, OVERVIEW]);

    const beforeXml = partText(original, OVERVIEW);
    const afterXml = partText(saved.bytes, OVERVIEW);
    assertDirtyPartPreserved(beforeXml, afterXml, {
      changedCells: ['D2'],
      changedRows: [2],
    });
    expect(afterXml).toContain('<c r="D2"><v>5</v></c>');
    expect(afterXml.slice(afterXml.indexOf('<extLst>'))).toBe(
      beforeXml.slice(beforeXml.indexOf('<extLst>')),
    );
    expect(afterXml).toContain('<legacyDrawing r:id="rId2"/>');
    expect(partText(saved.bytes, WORKBOOK_PART)).toContain('fullCalcOnLoad="1"');
    assertSemanticDiff(original, saved.bytes, [
      'calc.fullCalcOnLoad',
      'sheetParts.Overview.cells[7].cached',
      'sheetParts.Overview.cells[7].type',
    ]);
  });
});

describe('F3 styles-append.xlsx — formatting appends styles, indices stay stable', () => {
  const original = fixtureBytes('styles-append.xlsx');

  it('appends one cellXf and one font and leaves every existing record intact', () => {
    const opened = openFixture('styles-append.xlsx');
    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    const beforeStyles = opened.workbook.styles.map((style) => ({ ...style }));
    expect(beforeStyles).toHaveLength(5);

    session.dispatch('format.bold', {
      sheetId: sheet.id,
      range: { startRow: 1, startCol: 0, endRow: 1, endCol: 0 },
    });
    const saved = saveSession(opened, session);
    assertRelationshipsResolve(saved.bytes);
    expect(saved.noOp).toBe(false);
    assertOnlyPartsChanged(original, saved.bytes, ['xl/styles.xml', SHEET1]);

    const beforeXml = partText(original, 'xl/styles.xml');
    const afterXml = partText(saved.bytes, 'xl/styles.xml');
    const beforeXfs = styleElements(beforeXml, 'cellXfs');
    const afterXfs = styleElements(afterXml, 'cellXfs');
    expect(beforeXfs).toHaveLength(5);
    expect(afterXfs.slice(0, 5)).toEqual(beforeXfs);
    expect(afterXfs).toHaveLength(6);
    expect(styleElements(afterXml, 'fonts').slice(0, 3)).toEqual(styleElements(beforeXml, 'fonts'));
    // Unmodelled style sections keep their authored bytes.
    expect(sectionBytes(afterXml, 'dxfs')).toBe(sectionBytes(beforeXml, 'dxfs'));
    expect(sectionBytes(afterXml, 'tableStyles')).toBe(sectionBytes(beforeXml, 'tableStyles'));

    const reopened = openWorkbook(saved.bytes);
    expect(reopened.workbook.styles).toHaveLength(6);
    for (const style of beforeStyles) {
      expect(reopened.workbook.styles[style.id]).toMatchObject({
        numFmtId: style.numFmtId,
        fontId: style.fontId,
        fillId: style.fillId,
        borderId: style.borderId,
        rawXml: style.rawXml,
      });
    }
    expect(reopened.workbook.styles[5]!.font.bold).toBe(true);
    expect(reopened.workbook.styles[5]!.font.italic).toBe(true);
    expect(cellAt(reopened.workbook.sheets[0]!, 'A2').styleId).toBe(5);
    expect(reopened.workbook.styles[4]!.font.bold).toBe(false);
    assertSemanticDiff(original, saved.bytes, [
      'sheetParts.Sheet1.cells[3].styleId',
      'styles.cellXfs.length',
      'styles.fonts.length',
    ]);

    // A second save does not append the same records again.
    adoptSavedPackage(opened.workbook, saved.package);
    const second = saveWorkbook({
      workbook: opened.workbook,
      package: saved.package,
      dirtyCells: new Map(),
    });
    assertByteIdentical(saved.bytes, second.bytes, 'second styles save');
  });
});

describe('F4 shared-strings-append.xlsx — text edits append, indices stay stable', () => {
  const original = fixtureBytes('shared-strings-append.xlsx');

  it('appends a new entry after the untouched originals and points the cell at it', () => {
    const opened = openFixture('shared-strings-append.xlsx');
    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    editCell(session, sheet.id, 0, 0, textValue('Gamma & Co'));

    const saved = saveSession(opened, session);
    assertRelationshipsResolve(saved.bytes);
    assertOnlyPartsChanged(original, saved.bytes, ['xl/sharedStrings.xml', SHEET1]);

    const beforeXml = partText(original, 'xl/sharedStrings.xml');
    const afterXml = partText(saved.bytes, 'xl/sharedStrings.xml');
    const beforeEntries = sharedStringElements(beforeXml);
    const afterEntries = sharedStringElements(afterXml);
    expect(beforeEntries).toHaveLength(5);
    expect(afterEntries.slice(0, 5)).toEqual(beforeEntries);
    expect(afterEntries).toHaveLength(6);
    expect(afterEntries[5]).toContain('Gamma &amp; Co');
    expect(afterXml).toContain('count="8"');
    expect(afterXml).toContain('uniqueCount="6"');
    // The rich entry is not reused for plain text and stays untouched.
    expect(afterEntries[2]).toContain('<r>');

    expect(partText(saved.bytes, SHEET1)).toContain('<c r="A1" t="s"><v>5</v></c>');
    const reopened = openWorkbook(saved.bytes);
    expect(reopened.package.sharedStrings?.entries.map((entry) => entry.text)).toEqual([
      'Name',
      ' padded ',
      'Rich text',
      'Tom & Jerry',
      'Tail',
      'Gamma & Co',
    ]);
    expect(cellAt(reopened.workbook.sheets[0]!, 'A2').value).toEqual({
      type: 'text',
      value: 'Rich text',
    });
    assertSemanticDiff(original, saved.bytes, [
      'sheetParts.Sheet1.cells[0].cached',
      'sharedStrings.length',
    ]);

    // A second save does not append the same entry again.
    adoptSavedPackage(opened.workbook, saved.package);
    const second = saveWorkbook({
      workbook: opened.workbook,
      package: saved.package,
      dirtyCells: new Map(),
    });
    assertByteIdentical(saved.bytes, second.bytes, 'second shared-strings save');
  });

  it('reuses an existing entry without touching the shared-strings part', () => {
    const opened = openFixture('shared-strings-append.xlsx');
    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    editCell(session, sheet.id, 0, 0, textValue('Tail'));

    const saved = saveSession(opened, session);
    assertRelationshipsResolve(saved.bytes);
    assertOnlyPartsChanged(original, saved.bytes, [SHEET1]);
    expect(partText(saved.bytes, 'xl/sharedStrings.xml')).toBe(
      partText(original, 'xl/sharedStrings.xml'),
    );
    expect(partText(saved.bytes, SHEET1)).toContain('<c r="A1" t="s"><v>4</v></c>');
  });
});

describe('F5 formula-cached.xlsx — untouched cached values survive an edit elsewhere', () => {
  const original = fixtureBytes('formula-cached.xlsx');

  it('keeps every cached <v> and leaves calcChain alone', () => {
    const { opened, session } = openCalculated(original);
    const sheet = opened.workbook.sheets[0]!;
    editCell(session, sheet.id, 3, 0, textValue('note'));

    const saved = saveSession(opened, session);
    assertRelationshipsResolve(saved.bytes);
    expect(saved.noOp).toBe(false);
    assertOnlyPartsChanged(original, saved.bytes, [WORKBOOK_PART, SHEET1]);

    const beforeXml = partText(original, SHEET1);
    const afterXml = partText(saved.bytes, SHEET1);
    for (const fragment of [
      '<c r="B1"><f>A1+A2</f><v>30</v></c>',
      '<c r="C1" t="str"><f>IF(B1&gt;20,"big","small")</f><v>big</v></c>',
      '<c r="D1" t="e"><f>B1/0</f><v>#DIV/0!</v></c>',
      '<c r="E1" t="b"><f>B1&gt;25</f><v>1</v></c>',
      '<c r="B2"><f>SUM(A1:A2)*2</f><v>60</v></c>',
      '<c r="C2"><f>AVERAGE(A1:A2)</f><v>15</v></c>',
      '<c r="D2" t="str"><f>CONCATENATE("a","b")</f><v>ab</v></c>',
      '<c r="B3"><f>A3*2</f><v>6.283185307179586</v></c>',
    ]) {
      expect(beforeXml).toContain(fragment);
      expect(afterXml).toContain(fragment);
    }
    assertDirtyPartPreserved(beforeXml, afterXml, { changedRows: [4], changedCells: [] });
    // calcChain is only removed when formula storage changes (§4).
    expect(partText(saved.bytes, 'xl/calcChain.xml')).toBe(partText(original, 'xl/calcChain.xml'));
    assertSemanticDiff(original, saved.bytes, [
      'calc.fullCalcOnLoad',
      'sheetParts.Sheet1.cells.length',
      'sheetParts.Sheet1.rows.length',
    ]);
  });

  it('stays byte-identical when the engine recalculates to the same cached values', () => {
    const { opened, session } = openCalculated(original);
    const sheet = opened.workbook.sheets[0]!;
    expect(cellAt(sheet, 'B1').savedResult).toEqual(numberValue(30));
    expect(cellAt(sheet, 'C1').savedResult).toEqual({ type: 'text', value: 'big' });
    expect(cellAt(sheet, 'D1').savedResult).toEqual({ type: 'error', value: '#DIV/0!' });
    expect(cellAt(sheet, 'E1').savedResult).toEqual({ type: 'boolean', value: true });
    expect(cellAt(sheet, 'B2').stale).toBe(false);
    expect(session.staleCellCount()).toBe(0);

    const saved = saveNoOp(opened);
    expect(saved.noOp).toBe(true);
    assertByteIdentical(original, saved.bytes, 'F5 recalculated no-op save');
  });
});

describe('F6 formula-uncached.xlsx — an uncached formula recalculates and saves a value', () => {
  const original = fixtureBytes('formula-uncached.xlsx');

  it('carries the evaluated result into the saved part and flags recalculation', () => {
    const opened = openWorkbook(original);
    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    // Before the load phase: the file has no cached <v> for B1 or B2.
    expect(cellAt(sheet, 'B1').savedResult).toBeNull();
    expect(cellAt(sheet, 'B2').savedResult).toBeNull();
    calculateOnOpen(opened, session);
    expect(cellAt(sheet, 'B1').value).toEqual(numberValue(42));
    expect(cellAt(sheet, 'B1').savedResult).toEqual(numberValue(42));
    expect(cellAt(sheet, 'B1').stale).toBe(false);
    expect(cellAt(sheet, 'B2').value).toEqual(numberValue(23));
    expect(cellAt(sheet, 'C1').savedResult).toEqual(numberValue(22));

    const saved = saveSession(opened, session);
    assertRelationshipsResolve(saved.bytes);
    expect(saved.noOp).toBe(false);
    // Judgement call: F6 is the one fixture whose no-op save may not be
    // byte-identical. File-compat §6.1 requires the evaluated <v>, so writing
    // the value wins over identity (recorded in the lane report).
    expect(bytesEqual(original, saved.bytes)).toBe(false);
    assertOnlyPartsChanged(original, saved.bytes, [WORKBOOK_PART, SHEET1]);

    const afterXml = partText(saved.bytes, SHEET1);
    expect(afterXml).toContain('<c r="B1"><f>A1*A2</f><v>42</v></c>');
    expect(afterXml).toContain('<c r="B2"><f>SUM(A1:A2)</f><v>23</v></c>');
    // A cached formula the engine agreed with keeps its original bytes.
    expect(afterXml).toContain('<c r="C1"><f>A1+1</f><v>22</v></c>');
    expect(afterXml).toContain(
      '<c r="A3" t="str"><f>IF(A1&gt;10,"large","small")</f><v>large</v></c>',
    );
    expect(partText(saved.bytes, WORKBOOK_PART)).toContain('fullCalcOnLoad="1"');
    assertDirtyPartPreserved(partText(original, SHEET1), afterXml, {
      changedCells: ['B1', 'B2'],
      changedRows: [1, 2],
    });
    assertSemanticDiff(original, saved.bytes, [
      'calc.fullCalcOnLoad',
      'sheetParts.Sheet1.cells[1].cached',
      'sheetParts.Sheet1.cells[4].cached',
    ]);

    const reopened = openWorkbook(saved.bytes);
    expect(cellAt(reopened.workbook.sheets[0]!, 'B1').savedResult).toEqual(numberValue(42));
    // A second save is stable: the written value now equals the model cache.
    adoptSavedPackage(opened.workbook, saved.package);
    const second = saveWorkbook({
      workbook: opened.workbook,
      package: saved.package,
      dirtyCells: new Map(),
    });
    assertByteIdentical(saved.bytes, second.bytes, 'second save after the value settled');
  });
});

describe('F7 unsupported-function.xlsx — _xlfn.XLOOKUP is stale or #NAME?', () => {
  const original = fixtureBytes('unsupported-function.xlsx');

  it('shows the cached result marked stale and never writes a stale value', () => {
    const { opened, session } = openCalculated(original);
    const sheet = opened.workbook.sheets[0]!;
    expect(cellAt(sheet, 'B1').savedResult).toEqual(numberValue(5));
    expect(cellAt(sheet, 'B1').value).toEqual(numberValue(5));
    expect(cellAt(sheet, 'B1').stale).toBe(true);
    expect(cellAt(sheet, 'B2').savedResult).toBeNull();
    expect(cellAt(sheet, 'B2').value).toEqual({ type: 'error', value: '#NAME?' });
    expect(cellAt(sheet, 'B2').stale).toBe(true);
    expect(cellAt(sheet, 'C3').value).toEqual(numberValue(6));
    expect(cellAt(sheet, 'C3').stale).toBe(true);

    const xml = partText(original, SHEET1);
    expect(xml).toContain('<c r="B1"><f>_xlfn.XLOOKUP(A1,A1:A3,A1:A3)</f><v>5</v></c>');
    expect(xml).toContain('<c r="B2"><f>_xlfn.XLOOKUP(A2,A1:A3,A1:A3)</f></c>');

    // Judgement call: an untouched save returns the original file, so it cannot
    // introduce a stale value or a value written from one. Any written file
    // sets fullCalcOnLoad because it holds formulas (file-compat §6.1).
    const noOp = saveNoOp(opened);
    expect(noOp.noOp).toBe(true);
    assertByteIdentical(original, noOp.bytes, 'F7 untouched save');

    editCell(session, sheet.id, 0, 4, numberValue(1));
    const saved = saveSession(opened, session);
    assertRelationshipsResolve(saved.bytes);
    const afterXml = partText(saved.bytes, SHEET1);
    expect(afterXml).toContain('<c r="B1"><f>_xlfn.XLOOKUP(A1,A1:A3,A1:A3)</f><v>5</v></c>');
    expect(afterXml).toContain('<c r="B2"><f>_xlfn.XLOOKUP(A2,A1:A3,A1:A3)</f></c>');
    expect(afterXml).not.toContain('<v>#NAME?</v>');
    expect(partText(saved.bytes, WORKBOOK_PART)).toContain('fullCalcOnLoad="1"');
  });
});

describe('F8 array-shared.xlsx — legacy arrays and shared groups are read-only', () => {
  const original = fixtureBytes('array-shared.xlsx');

  it('preserves t="shared" and t="array" verbatim and refuses content edits', () => {
    const opened = openFixture('array-shared.xlsx');
    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;

    const master = cellAt(sheet, 'B2');
    expect(master.formula).toBe('A2*10');
    expect(master.formulaSource).toMatchObject({
      kind: 'shared-master',
      sharedIndex: '0',
      range: 'B2:B4',
      xml: '<f t="shared" ref="B2:B4" si="0">A2*10</f>',
    });
    const follower = cellAt(sheet, 'B3');
    expect(follower.formula).toBeNull();
    expect(follower.formulaDerived).toBe(true);
    expect(follower.formulaSource).toMatchObject({
      kind: 'shared-follower',
      sharedIndex: '0',
      xml: '<f t="shared" si="0"/>',
    });
    expect(cellAt(sheet, 'D6').formulaSource).toMatchObject({
      kind: 'array',
      range: 'D6:D7',
      xml: '<f t="array" ref="D6:D7">A6:A7*2</f>',
    });

    const followerEdit = session.dispatch('cell.set', {
      sheetId: sheet.id,
      row: 2,
      col: 1,
      value: numberValue(99),
    });
    expect(followerEdit.changes).toEqual([]);
    expect(followerEdit.ui).toMatchObject({ type: 'toast' });
    const arrayEdit = session.dispatch('cell.set', {
      sheetId: sheet.id,
      row: 6,
      col: 3,
      value: numberValue(99),
    });
    expect(arrayEdit.changes).toEqual([]);
    expect(arrayEdit.ui).toMatchObject({ type: 'toast' });

    editCell(session, sheet.id, 8, 0, numberValue(42));
    const saved = saveSession(opened, session);
    assertRelationshipsResolve(saved.bytes);
    assertOnlyPartsChanged(original, saved.bytes, [WORKBOOK_PART, SHEET1]);
    const beforeXml = partText(original, SHEET1);
    const afterXml = partText(saved.bytes, SHEET1);
    for (const fragment of [
      '<f t="shared" ref="B2:B4" si="0">A2*10</f>',
      '<f t="shared" si="0"/>',
      '<f t="array" ref="D6:D7">A6:A7*2</f>',
    ]) {
      expect(beforeXml).toContain(fragment);
      expect(afterXml).toContain(fragment);
    }
    expect(afterXml).toContain('<c r="B2"><f t="shared" ref="B2:B4" si="0">A2*10</f><v>30</v></c>');
    expect(afterXml).toContain('<c r="D6"><f t="array" ref="D6:D7">A6:A7*2</f><v>4</v></c>');
    assertDirtyPartPreserved(beforeXml, afterXml, { changedRows: [9] });
    assertSemanticDiff(original, saved.bytes, [
      'calc.fullCalcOnLoad',
      'sheetParts.Sheet1.cells.length',
      'sheetParts.Sheet1.rows.length',
    ]);

    const reopened = openWorkbook(saved.bytes);
    expect(cellAt(reopened.workbook.sheets[0]!, 'B3').formulaSource).toMatchObject({
      kind: 'shared-follower',
      sharedIndex: '0',
    });
  });

  it('refuses a direct model edit of a shared follower or array member at save', () => {
    const opened = openFixture('array-shared.xlsx');
    const sheet = opened.workbook.sheets[0]!;
    sheet.cells.set('2:1', { ...cellAt(sheet, 'B3'), value: numberValue(99) });
    expect(() =>
      saveWorkbook({
        workbook: opened.workbook,
        package: opened.package,
        dirtyCells: new Map([[sheet.id, new Set(['2:1'])]]),
      }),
    ).toThrow(WorkbookSaveError);

    const openedArray = openFixture('array-shared.xlsx');
    const arraySheet = openedArray.workbook.sheets[0]!;
    arraySheet.cells.set('6:3', { ...cellAt(arraySheet, 'D6'), formula: 'A6:A7*3' });
    expect(() =>
      saveWorkbook({
        workbook: openedArray.workbook,
        package: openedArray.package,
        dirtyCells: new Map([[arraySheet.id, new Set(['6:3'])]]),
      }),
    ).toThrow(WorkbookSaveError);
  });
});

describe('F15 date1904.xlsx — 1904 dates display correctly and the flag survives', () => {
  const original = fixtureBytes('date1904.xlsx');

  it('reads the flag and formats serials from the 1904 epoch', () => {
    const opened = openFixture('date1904.xlsx');
    const workbook = opened.workbook;
    expect(workbook.date1904).toBe(true);
    const sheet = workbook.sheets[0]!;
    const display = (ref: string, date1904: boolean): string => {
      const cell = cellAt(sheet, ref);
      return formatValue(cell.value, workbook.styles[cell.styleId]!.numFmtCode, date1904);
    };
    expect(display('B1', true)).toBe('1904-01-01');
    expect(display('B2', true)).toBe('1904-02-29');
    expect(display('B3', true)).toBe('2026-11-19');
    expect(display('B4', true)).toBe('2026-11-19 12:00');
    // The same serial means a different date in the 1900 system: the flag is
    // what makes the display correct.
    expect(display('B3', false)).not.toBe(display('B3', true));
    expect(display('B2', false)).not.toBe(display('B2', true));
  });

  it('is byte-identical on a no-op save and keeps the flag after an edit', () => {
    const opened = openFixture('date1904.xlsx');
    const noOp = saveNoOp(opened);
    expect(noOp.noOp).toBe(true);
    assertByteIdentical(original, noOp.bytes, 'date1904 no-op save');

    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    editCell(session, sheet.id, 0, 0, textValue('Start'));
    const saved = saveSession(opened, session);
    assertRelationshipsResolve(saved.bytes);
    expect(saved.noOp).toBe(false);
    assertOnlyPartsChanged(original, saved.bytes, [SHEET1]);
    expect(partText(saved.bytes, WORKBOOK_PART)).toBe(partText(original, WORKBOOK_PART));
    expect(partText(saved.bytes, WORKBOOK_PART)).toContain('date1904="1"');
    expect(openWorkbook(saved.bytes).workbook.date1904).toBe(true);
    assertSemanticDiff(original, saved.bytes, ['sheetParts.Sheet1.cells[0].inlineText']);
  });
});

describe('every fixture without an open-phase recalculation is identity-stable', () => {
  const fixtures = [
    'identity-minimal.xlsx',
    'identity-full.xlsx',
    'styles-append.xlsx',
    'shared-strings-append.xlsx',
    'formula-cached.xlsx',
    'unsupported-function.xlsx',
    'array-shared.xlsx',
    'date1904.xlsx',
  ];

  it.each(fixtures)('no-op save of %s is byte-identical', (name) => {
    const original = fixtureBytes(name);
    const saved = saveNoOp(openWorkbook(original));
    expect(saved.noOp).toBe(true);
    assertByteIdentical(original, saved.bytes, `${name} no-op save`);
    assertOnlyPartsChanged(original, saved.bytes, []);
    assertRelationshipsResolve(saved.bytes);
  });
});

describe('identity harness self-checks', () => {
  it('detects a one-byte difference', () => {
    const left = new Uint8Array([1, 2, 3]);
    const right = new Uint8Array([1, 2, 4]);
    expect(() => assertByteIdentical(left, right)).toThrow(/not byte-identical/);
    expect(() => assertByteIdentical(left, left)).not.toThrow();
  });

  it('detects an undeclared part change', () => {
    const original = fixtureBytes('identity-minimal.xlsx');
    const opened = openWorkbook(original);
    const session = new DocumentSession(opened.workbook);
    editCell(session, opened.workbook.sheets[0]!.id, 1, 0, numberValue(9));
    const saved = saveSession(opened, session);
    assertRelationshipsResolve(saved.bytes);
    expect(changedPartNames(original, saved.bytes)).toEqual([SHEET1]);
    expect(() => assertOnlyPartsChanged(original, saved.bytes, [])).toThrow(/Changed parts/);
  });

  it('detects a changed unmodelled child and an undeclared semantic difference', () => {
    const original = fixtureBytes('identity-full.xlsx');
    const overview = partText(original, OVERVIEW);
    const tampered = overview.replace('preserve-me', 'changed-me');
    expect(() => assertDirtyPartPreserved(overview, tampered, {})).toThrow(/<extLst>/);
    expect(diffSemantic({ a: { b: [1, 2] } }, { a: { b: [1, 3] } })).toEqual(['a.b[1]']);
    expect(() => assertSemanticDiff(original, fixtureBytes('identity-minimal.xlsx'), [])).toThrow(
      /Semantic difference/,
    );
  });

  it('detects a relationship that points at a missing part', () => {
    const original = fixtureBytes('identity-minimal.xlsx');
    const parts = new Map(Object.entries(unzipSync(original)));
    const relsPath = 'xl/_rels/workbook.xml.rels';
    const rels = new TextDecoder().decode(parts.get(relsPath)!);
    parts.set(
      relsPath,
      new TextEncoder().encode(rels.replace('worksheets/sheet1.xml', 'worksheets/missing.xml')),
    );
    const tampered = zipSync(Object.fromEntries(parts), {
      level: 6,
      mtime: new Date('1980-01-01T00:00:00Z'),
    });
    expect(() => assertRelationshipsResolve(tampered)).toThrow(/missing part/);
  });
});

function styleElements(xml: string, section: string): string[] {
  const match = new RegExp(`<${section}\\b[^>]*>([\\s\\S]*?)</${section}>`).exec(xml);
  if (match === null) return [];
  const elements = [
    ...match[1]!.matchAll(/<xf\b[^>]*\/>|<xf\b[^>]*>[\s\S]*?<\/xf>|<font\b[^>]*>[\s\S]*?<\/font>/g),
  ].map((entry) => entry[0]!);
  return elements.filter((entry) =>
    section === 'fonts' ? entry.startsWith('<font') : entry.startsWith('<xf'),
  );
}

function sectionBytes(xml: string, section: string): string {
  const match = new RegExp(`<${section}\\b[^>]*(?:/>|>[\\s\\S]*?</${section}>)`).exec(xml);
  return match?.[0] ?? '';
}

function sharedStringElements(xml: string): string[] {
  return [...xml.matchAll(/<si\b[^>]*>[\s\S]*?<\/si>/g)].map((entry) => entry[0]!);
}
