import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { DocumentSession } from '@commands/session';
import { COMMAND_INVENTORY, inventoryCommand } from '@commands/inventory';
import { createCommandRegistry } from '@commands/registry';
import { formulaHasUnsupportedConstruct, isEngineFormulaEligible } from '@engine/formula-text';
import { listUnsupportedConditionalFormats } from '@ooxml/conditional-format-xml';
import { listPreservedParts } from '@ooxml/part-policy';
import { openWorkbook } from '@ooxml/open';
import { saveWorkbook } from '@ooxml/save';
import { findOrAppendStyle, patchedStyleInput } from '@model/style-table';
import type { GridRange, Sheet } from '@model/types';
import { cellKey, getCell, setCellSnapshot } from '@model/types';
import { numberValue, textValue } from '@model/value';
import { createEmptyWorkbook, createSheet } from '@model/workbook';
import { createSheetProjection } from '@view/projection';
import { createPrintSnapshot } from '../src/print/adapter';
import { createPrintPlan } from '../src/print/planner';
import type { PrintPlan } from '../src/print/types';
import { buildXlsx, unzipPart } from './helpers/build-xlsx';
import { fixtureBytes, partText } from './helpers/roundtrip';

/**
 * Pinning tests for docs/known-deviations.md D-1 to D-11. One test per entry:
 * the name carries the D-id and the documented claim, and the body asserts the
 * documented outcome against the code as it stands. If one of these fails, the
 * code and the documented deviation have parted ways; fix one of them, do not
 * weaken the pin silently.
 *
 * D-12 is pinned in tests/m7-page-setup-group.test.ts and D-13 in
 * tests/formula-arity.test.ts, next to the features they describe.
 */

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

function range(startRow: number, startCol: number, endRow: number, endCol: number): GridRange {
  return { startRow, startCol, endRow, endCol };
}

function contextFor(sheetId: string, selection: GridRange) {
  return {
    activeSheetId: sheetId,
    selection,
    activeCell: { row: selection.startRow, col: selection.startCol },
  };
}

function workbookWithSheet() {
  const workbook = createEmptyWorkbook();
  const sheet = createSheet('Sheet1', 'xl/worksheets/sheet1.xml');
  workbook.sheets.push(sheet);
  return { workbook, sheet };
}

function put(sheet: Sheet, row: number, col: number, value: ReturnType<typeof numberValue>): void {
  setCellSnapshot(sheet, row, col, {
    value,
    formula: null,
    formulaDerived: false,
    savedResult: null,
    styleId: 0,
    richText: null,
  });
}

function putFormula(sheet: Sheet, row: number, col: number, formula: string): void {
  setCellSnapshot(sheet, row, col, {
    value: numberValue(0),
    formula,
    formulaDerived: false,
    savedResult: numberValue(0),
    styleId: 0,
    richText: null,
    stale: false,
  });
}

// ---------------------------------------------------------------------------
// D-1 data validation
// ---------------------------------------------------------------------------

describe('known deviations (docs/known-deviations.md)', () => {
  it('pins known deviation D-1: data validation is preserved, not enforced', () => {
    // Excel blocks a value its validation rule forbids. calcplus accepts the
    // value and keeps the rule bytes unchanged, so the file still carries the
    // original rules (file-compat §3).
    const validation =
      '<dataValidations count="1"><dataValidation type="whole" operator="between" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>';
    const worksheet = `${HEADER}
<worksheet xmlns="${NS}">
  <dimension ref="A1"/>
  <sheetData><row r="1"><c r="A1"><v>5</v></c></row></sheetData>
  ${validation}
</worksheet>`;
    const original = buildXlsx(worksheet);
    const opened = openWorkbook(original);
    const sheet = opened.workbook.sheets[0]!;
    expect(sheet.validation).toHaveLength(1);
    expect(sheet.validation[0]!.rawXml).toContain('type="whole"');

    // 99 is outside the allowed 1..10 and the edit is not blocked or warned.
    const session = new DocumentSession(opened.workbook);
    const result = session.dispatch('cell.set', {
      sheetId: sheet.id,
      row: 0,
      col: 0,
      value: numberValue(99),
    });
    expect(result.changes).toHaveLength(1);
    expect(result.ui).toBeUndefined();
    expect(getCell(sheet, 0, 0)?.value).toEqual(numberValue(99));

    // The rule round-trips unchanged around the edited cell.
    const saved = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
    });
    expect(unzipPart(saved.bytes, 'xl/worksheets/sheet1.xml')).toContain(validation);
    const reopened = openWorkbook(saved.bytes);
    expect(reopened.workbook.sheets[0]!.validation[0]!.rawXml).toContain('type="whole"');
    expect(reopened.workbook.sheets[0]!.validation[0]!.rawXml).toContain('sqref="A1"');
  });

  // -------------------------------------------------------------------------
  // D-2 protection
  // -------------------------------------------------------------------------

  it('pins known deviation D-2: protection is preserved, not enforced', () => {
    // Excel refuses an edit to a protected sheet. calcplus lets the edit
    // through and saves the protection XML unchanged (file-compat §6.7).
    const protection =
      '<sheetProtection sheet="1" objects="1" scenarios="1" formatCells="0" formatColumns="1" password="CC1A"/>';
    const worksheet = `${HEADER}
<worksheet xmlns="${NS}">
  <dimension ref="A1:B1"/>
  <sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row></sheetData>
  ${protection}
</worksheet>`;
    const original = buildXlsx(worksheet, {
      workbookExtras: '<workbookProtection workbookPassword="CC1A" lockStructure="1"/>',
    });
    const opened = openWorkbook(original);
    const sheet = opened.workbook.sheets[0]!;
    expect(sheet.protection?.rawXml).toContain('sheet="1"');
    expect(opened.workbook.protection?.rawXml).toContain('lockStructure="1"');

    // The edit lands on the protected sheet.
    const session = new DocumentSession(opened.workbook);
    session.dispatch('cell.set', { sheetId: sheet.id, row: 0, col: 0, value: numberValue(42) });
    expect(getCell(sheet, 0, 0)?.value).toEqual(numberValue(42));

    // Both protection records are saved unchanged.
    const saved = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
    });
    expect(unzipPart(saved.bytes, 'xl/workbook.xml')).toBe(unzipPart(original, 'xl/workbook.xml'));
    expect(unzipPart(saved.bytes, 'xl/worksheets/sheet1.xml')).toContain(protection);
  });

  // -------------------------------------------------------------------------
  // D-3 print scope
  // -------------------------------------------------------------------------

  it('pins known deviation D-3: print honours the print area, orientation, fit-to-pages, uniform scale and margins', () => {
    // Headers and footers, print titles, manual page breaks, page order and
    // independent X/Y scale are v2 (print-pdf §1, §8): the print plan must be
    // identical with and without them. Fit-to-N-pages (Excel's 1-32767 range)
    // and uniform scale 10%-400% are v1 (maintainer ruling 2026-11-22), so
    // the pins below also prove a multi-page fit is honoured and the scale is
    // one factor on both axes.
    // Six rows of twenty columns, so fit-to-width has something to shrink and
    // the print area has a row to drop.
    const rows = [1, 2, 3, 4, 5, 6]
      .map(
        (row) =>
          `<row r="${row}">` +
          Array.from({ length: 20 }, (_, col) => {
            const label = String.fromCharCode(65 + (col % 26));
            return `<c r="${label}${row}"><v>${row * 100 + col}</v></c>`;
          }).join('') +
          '</row>',
      )
      .join('');
    const honoured = `${HEADER}
<worksheet xmlns="${NS}">
  <sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>
  <dimension ref="A1:T6"/>
  <cols><col min="1" max="20" width="20" customWidth="1"/></cols>
  <sheetData>${rows}</sheetData>
  <pageMargins left="0.5" right="0.5" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>
  <pageSetup paperSize="1" orientation="landscape" fitToWidth="1" fitToHeight="0"/>
</worksheet>`;
    const printAreaNames =
      '<definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">Sheet1!$A$1:$T$5</definedName></definedNames>';
    const v2Extras = `${HEADER}
<worksheet xmlns="${NS}">
  <sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>
  <dimension ref="A1:T6"/>
  <cols><col min="1" max="20" width="20" customWidth="1"/></cols>
  <sheetData>${rows}</sheetData>
  <pageMargins left="0.5" right="0.5" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>
  <pageSetup paperSize="1" orientation="landscape" fitToWidth="1" fitToHeight="0" pageOrder="overThenDown"/>
  <headerFooter><oddHeader>&amp;CCalcplus</oddHeader><oddFooter>Page &amp;P</oddFooter></headerFooter>
  <rowBreaks count="1" manualBreakCount="1"><brk id="3" max="16383" man="1"/></rowBreaks>
  <colBreaks count="1" manualBreakCount="1"><brk id="5" max="1048575" man="1"/></colBreaks>
</worksheet>`;
    const v2Names =
      '<definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">Sheet1!$A$1:$T$5</definedName><definedName name="_xlnm.Print_Titles" localSheetId="0">Sheet1!$1:$2</definedName></definedNames>';

    function planOf(worksheetXml: string, workbookExtras: string): PrintPlan {
      const opened = openWorkbook(buildXlsx(worksheetXml, { workbookExtras }));
      const sheet = opened.workbook.sheets[0]!;
      const projection = createSheetProjection(opened.workbook, sheet.id, { version: 1 });
      return createPrintPlan(createPrintSnapshot(opened.workbook, projection));
    }

    const plan = planOf(honoured, printAreaNames);
    const planV2 = planOf(v2Extras, v2Names);
    // The v2 features change nothing: no headers or footers, no repeated title
    // rows, no honoured break and no honoured page order.
    expect({ ...planV2, sheetId: plan.sheetId }).toEqual(plan);

    // The honoured settings drive the plan.
    expect(plan.paper.orientation).toBe('landscape');
    expect(plan.scale).toBeLessThan(1);
    expect(plan.scale).toBeGreaterThanOrEqual(0.1);
    expect(plan.pages[0]!.body.x).toBeCloseTo(0.5 * 72, 6);
    expect(plan.pageSets[0]!.range).toEqual({ startRow: 0, startCol: 0, endRow: 4, endCol: 19 });

    // Fit-to-pages: fitting twenty columns to two pages wide shrinks less
    // than fitting them to one, so the multi-page fit is honoured.
    const fitTwo = planOf(honoured.replace('fitToWidth="1"', 'fitToWidth="2"'), printAreaNames);
    expect(fitTwo.scale).toBeGreaterThan(plan.scale);
    expect(fitTwo.scale).toBeLessThanOrEqual(1);

    // Uniform scale: one factor applied to both axes, clamped to Excel's
    // 10%-400%; the model carries no independent X/Y scale.
    const unscaled = planOf(
      honoured
        .replace('<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>', '')
        .replace(' fitToWidth="1" fitToHeight="0"', ''),
      printAreaNames,
    );
    const scaled = planOf(
      honoured
        .replace('<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>', '')
        .replace('fitToWidth="1" fitToHeight="0"', 'scale="150"'),
      printAreaNames,
    );
    expect(scaled.scale).toBe(1.5);
    expect(scaled.pages[0]!.rows[1]!.height).toBeCloseTo(
      unscaled.pages[0]!.rows[1]!.height * 1.5,
      6,
    );
    // Column widths round to 2 decimals in the model, so compare loosely.
    expect(scaled.pages[0]!.columns[0]!.width).toBeCloseTo(
      unscaled.pages[0]!.columns[0]!.width * 1.5,
      1,
    );
  });

  // -------------------------------------------------------------------------
  // D-4 charts, drawings and pivots
  // -------------------------------------------------------------------------

  it('pins known deviation D-4: charts, drawings and pivots are preserved but not rendered', () => {
    // The grid draws none of these: they live in the model only as opaque
    // parts, they survive an edit byte-for-byte, and the compatibility report
    // says they are not rendered. A pivot is reported "preserved; never
    // refreshed" after an edit to its source range (file-compat §4, §6.4).
    const original = fixtureBytes('charts-images.xlsx');
    const opened = openWorkbook(original);
    const sheet = opened.workbook.sheets[0]!;
    const paths = new Set([...opened.workbook.opaqueParts.keys(), ...sheet.opaqueParts.keys()]);
    expect(paths.has('xl/charts/chart1.xml')).toBe(true);

    const report = listPreservedParts({ package: opened.package, workbook: opened.workbook });
    const chart = report.find((entry) => entry.category === 'chart');
    expect(chart?.note).toContain('not rendered');

    // An unrelated edit keeps the chart and drawing bytes identical.
    const session = new DocumentSession(opened.workbook);
    session.dispatch('cell.set', { sheetId: sheet.id, row: 40, col: 5, value: numberValue(1) });
    const saved = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
      dirtyDimensions: session.dirtyDimensionSnapshot(),
      dirtyMerges: session.dirtyMergeSheets(),
    });
    expect(partText(saved.bytes, 'xl/charts/chart1.xml')).toBe(
      partText(original, 'xl/charts/chart1.xml'),
    );

    // A pivot edited inside its source range is reported preserved and never
    // refreshed, not recalculated.
    const pivotBytes = fixtureBytes('pivot-external.xlsx');
    const pivotOpened = openWorkbook(pivotBytes);
    const pivotSheet = pivotOpened.workbook.sheets[0]!;
    const pivotSession = new DocumentSession(pivotOpened.workbook);
    pivotSession.dispatch('cell.set', {
      sheetId: pivotSheet.id,
      row: 2,
      col: 2,
      value: numberValue(21),
    });
    const listing = listPreservedParts({
      package: pivotOpened.package,
      workbook: pivotOpened.workbook,
      dirtyCells: pivotSession.dirtyCellMap(),
    });
    const cache = listing.find((entry) => entry.category === 'pivotCache');
    expect(cache?.status).toBe('stale');
    expect(cache?.reason).toBe('pivot-cache-not-refreshed');
    expect(cache?.note).toContain('never refreshed');
  });

  // -------------------------------------------------------------------------
  // D-5 conditional formatting
  // -------------------------------------------------------------------------

  it('pins known deviation D-5: conditional formatting renders where supported but cannot be authored', () => {
    // A supported rule paints; a formula-based rule is preserved and reported,
    // never painted; and no command creates or edits rules (conditional
    // formatting §8, §11). The supported-rule paint matrix lives in
    // tests/m5-cf-paint.test.ts.
    const stylesXml = `${HEADER}
<styleSheet xmlns="${NS}">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
  <dxfs count="2">
    <dxf><font><color rgb="FF9C0006"/></font></dxf>
    <dxf><font><b/></font></dxf>
  </dxfs>
</styleSheet>`;
    const unsupportedRule =
      '<cfRule type="expression" dxfId="1" priority="2"><formula>$A2&gt;100</formula></cfRule>';
    const worksheet = `${HEADER}
<worksheet xmlns="${NS}">
  <dimension ref="A1:A3"/>
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Score</t></is></c></row>
    <row r="2"><c r="A2"><v>5</v></c></row>
    <row r="3"><c r="A3"><v>120</v></c></row>
  </sheetData>
  <conditionalFormatting sqref="A2:A3">
    <cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>6</formula></cfRule>
    ${unsupportedRule}
  </conditionalFormatting>
</worksheet>`;
    const original = buildXlsx(worksheet, { stylesXml });
    const opened = openWorkbook(original);
    const sheet = opened.workbook.sheets[0]!;
    const rules = opened.conditionalFormatRules.get(sheet.id)!;
    expect(rules.map((rule) => rule.kind)).toEqual(['cellIs', 'unsupported']);
    const unsupported = rules.find((rule) => rule.kind === 'unsupported')!;
    if (unsupported.kind !== 'unsupported') throw new Error('expected an unsupported rule');
    expect(unsupported.reason).toBe('formula');

    // The supported rule paints; the formula-based rule never paints.
    const projection = createSheetProjection(opened.workbook, sheet.id, {
      version: 1,
      conditionalFormats: { rules, today: { year: 2026, month: 11, day: 19 } },
    });
    const formatting = projection.conditionalFormatting;
    const painted = formatting!.evaluate(2, 0);
    expect(painted.ruleIds).toHaveLength(1);
    expect(painted.ruleIds).toContain(rules.find((rule) => rule.kind === 'cellIs')!.id);
    expect(painted.ruleIds).not.toContain(unsupported.id);
    expect(painted.overlay.font?.color).toBe('#9C0006');
    expect(formatting!.evaluate(1, 0).ruleIds).toEqual([]);

    // The unsupported rule is reported and preserved byte-for-byte.
    expect(
      listUnsupportedConditionalFormats(opened.workbook).map((entry) => [
        entry.sheetName,
        entry.ruleType,
        entry.reason,
      ]),
    ).toEqual([['Sheet1', 'expression', 'formula']]);
    const session = new DocumentSession(opened.workbook);
    session.dispatch('cell.set', { sheetId: sheet.id, row: 5, col: 1, value: numberValue(7) });
    const saved = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
    });
    expect(unzipPart(saved.bytes, 'xl/worksheets/sheet1.xml')).toContain(unsupportedRule);

    // Authoring is v2: the only conditional-formatting command is the
    // deliberately disabled v2 placeholder, and the runtime registry has
    // nothing that creates or edits rules.
    const render = inventoryCommand('format.conditionalRender')!;
    expect(render.implemented).toBe(false);
    expect(render.unimplementedReason).toBe('v2-cut');
    expect(render.ribbon?.[0]?.disabled).toBe(true);
    expect(
      createCommandRegistry()
        .list()
        .map((command) => command.id)
        .filter((id) => id.startsWith('cf.')),
    ).toEqual([]);
  });

  // -------------------------------------------------------------------------
  // D-6 comments, hyperlinks and defined names
  // -------------------------------------------------------------------------

  it('pins known deviation D-6: comments, hyperlinks and defined names are read-only', () => {
    // Reading and following are view behaviour. Nothing in the command surface
    // creates, edits or deletes comments, hyperlinks or defined names
    // (feature matrix §11, file-compat §3).
    expect(
      createCommandRegistry()
        .list()
        .map((command) => command.id)
        .filter((id) => /^(comment|hyperlink|name)\./.test(id)),
    ).toEqual([]);

    // The command map carries only the read-only markers for the three
    // features, and none of them is an implemented command.
    const readOnlyIds = COMMAND_INVENTORY.filter(
      (entry) =>
        /^(comment|hyperlink|name)\./.test(entry.id) || /Comments|Hyperlink/.test(entry.id),
    );
    expect(readOnlyIds.map((entry) => entry.id).sort()).toEqual([
      'cell.followHyperlink',
      'name.preserve',
      'name.resolve',
      'review.showComments',
    ]);
    for (const entry of readOnlyIds) {
      expect(entry.implemented, entry.id).toBe(false);
      expect(entry.unimplementedReason, entry.id).toBe('not-a-command');
    }
  });

  // -------------------------------------------------------------------------
  // D-7 sort, AutoFilter and tables
  // -------------------------------------------------------------------------

  it('pins known deviation D-7: tables stay on the plain-range path and structured references are not evaluated', () => {
    // A structured reference never reaches the engine: the formula keeps its
    // authored text and its stored result (formula semantics §2).
    expect(isEngineFormulaEligible('SUM(Table1[Amount])')).toBe(false);
    expect(formulaHasUnsupportedConstruct('SUM(Table1[Amount])')).toBe(true);

    // A table's own filter is never read into the sheet: filtering and sorting
    // use the sheet's plain range, and the table part is preserved byte for
    // byte (sort and AutoFilter §7). The plain-range filter and sort
    // themselves are pinned in tests/m5-filter-commands.test.ts and
    // tests/m5-sort-commands.test.ts.
    const tablePart = `${HEADER}
<table xmlns="${NS}" id="1" name="Table1" displayName="Table1" ref="A1:B4" totalsRowShown="0">
  <autoFilter ref="A1:B4"><filterColumn colId="0"><filters><filter val="10"/></filters></filterColumn></autoFilter>
  <tableColumns count="2"><tableColumn id="1" name="Name"/><tableColumn id="2" name="Value"/></tableColumns>
</table>`;
    const worksheet = `${HEADER}
<worksheet xmlns="${NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1:C4"/>
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Value</t></is></c></row>
    <row r="2"><c r="A2"><v>10</v></c><c r="B2"><v>4</v></c><c r="C2"><f>SUM(Table1[Amount])</f><v>7</v></c></row>
    <row r="3"><c r="A3"><v>20</v></c><c r="B3"><v>6</v></c></row>
    <row r="4"><c r="A4"><v>30</v></c><c r="B4"><v>8</v></c></row>
  </sheetData>
  <tableParts count="1"><tablePart r:id="rId1"/></tableParts>
</worksheet>`;
    const original = buildXlsx(worksheet, {
      extraParts: {
        'xl/tables/table1.xml': tablePart,
        'xl/worksheets/_rels/sheet1.xml.rels': `${HEADER}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>
</Relationships>`,
      },
      contentTypeOverrides: [
        [
          'xl/tables/table1.xml',
          'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml',
        ],
      ],
    });
    const opened = openWorkbook(original);
    const sheet = opened.workbook.sheets[0]!;
    // The structured-reference formula shows its stored result, not a value.
    expect(getCell(sheet, 1, 2)?.formula).toBe('SUM(Table1[Amount])');
    expect(getCell(sheet, 1, 2)?.value).toEqual(numberValue(7));
    // The table's own filter is not the sheet's filter.
    expect(sheet.autoFilter).toBeNull();

    const session = new DocumentSession(opened.workbook);
    session.dispatch('cell.set', { sheetId: sheet.id, row: 2, col: 1, value: numberValue(9) });
    const saved = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
    });
    expect(unzipPart(saved.bytes, 'xl/tables/table1.xml')).toBe(tablePart);
    expect(unzipPart(saved.bytes, 'xl/worksheets/sheet1.xml')).toContain(
      '<f>SUM(Table1[Amount])</f>',
    );

    // Tables are preserved but not editable: no command authors them.
    expect(inventoryCommand('fidelity.tables')?.implemented).toBe(false);
    expect(inventoryCommand('fidelity.tables')?.unimplementedReason).toBe('not-a-command');
  });

  // -------------------------------------------------------------------------
  // D-8 structural edits and #REF!
  // -------------------------------------------------------------------------

  it('pins known deviation D-8: a structural edit marks a broken reference #REF! instead of refusing', () => {
    // The edit applies, the reference whose target is gone reads #REF!, and
    // the user is told once. One undo restores both the data and the
    // reference (ticket 32). The formulas, names, hyperlinks and chart-series
    // matrix is pinned in tests/m3c-range-transform.test.ts and
    // tests/m6-transforms.test.ts.
    const { workbook, sheet } = workbookWithSheet();
    put(sheet, 2, 0, numberValue(3));
    putFormula(sheet, 0, 1, 'A3+1');
    const session = new DocumentSession(workbook);

    const result = session.dispatch(
      'structure.deleteRows',
      {},
      contextFor(sheet.id, range(2, 0, 2, 0)),
    );
    expect(result.structure?.state).toBe('applied');
    expect(sheet.cells.get(cellKey(0, 1))?.formula).toBe('#REF!+1');
    expect(result.ui).toMatchObject({ type: 'toast' });
    expect(result.ui?.message).toContain('1 reference');
    expect(result.ui?.message).toContain('#REF!');

    // One undo restores the data and the reference together.
    session.dispatch('edit.undo', {}, { activeSheetId: sheet.id });
    expect(sheet.cells.get(cellKey(0, 1))?.formula).toBe('A3+1');
    expect(sheet.cells.get(cellKey(2, 0))?.value).toEqual(numberValue(3));
    expect(session.transactionCount).toBe(0);

    // Refusal remains only for data loss: an edit splitting a merged cell is
    // still blocked and writes no transaction.
    const merged = workbookWithSheet();
    merged.sheet.merges.push({ range: range(0, 0, 1, 1) });
    const mergedSession = new DocumentSession(merged.workbook);
    const blocked = mergedSession.dispatch(
      'structure.insertCells',
      { shift: 'right' },
      contextFor(merged.sheet.id, range(0, 0, 0, 0)),
    );
    expect(blocked.structure?.state).toBe('blocked');
    expect(blocked.ui).toEqual({ type: 'toast', message: 'Cannot change part of a merged cell.' });
    expect(mergedSession.transactionCount).toBe(0);
  });

  // -------------------------------------------------------------------------
  // D-9 broken internal hyperlinks
  // -------------------------------------------------------------------------

  it('pins known deviation D-9: a broken internal hyperlink is written as location="#REF!"', () => {
    // calcplus's own convention (ticket 32, open question 3, settled
    // 2026-11-22): a target the edit removes is written as
    // `location="#REF!"`, matching formulas and names. Excel instead freezes
    // the location verbatim; that oracle side is pinned in the next test and
    // the difference is documented in docs/known-deviations.md D-9. The
    // fixture side of the same convention is pinned in
    // tests/ticket32-multi-part-saves.test.ts.
    const worksheet = `${HEADER}
<worksheet xmlns="${NS}">
  <dimension ref="A1:B5"/>
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
    <row r="5"><c r="B5" t="inlineStr"><is><t>up</t></is></c></row>
  </sheetData>
  <hyperlinks><hyperlink ref="B5" location="Sheet1!A1" display="up"/></hyperlinks>
</worksheet>`;
    const opened = openWorkbook(buildXlsx(worksheet));
    const sheet = opened.workbook.sheets[0]!;
    const session = new DocumentSession(opened.workbook);
    const result = session.dispatch(
      'structure.deleteRows',
      {},
      contextFor(sheet.id, range(0, 0, 0, 0)),
    );
    expect(result.structure?.state).toBe('applied');
    // The hyperlink moves up with its row and its target collapses.
    expect(sheet.hyperlinks.get(cellKey(3, 1))?.location).toBe('#REF!');

    const saved = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
      dirtyDimensions: session.dirtyDimensionSnapshot(),
      dirtyMerges: session.dirtyMergeSheets(),
    });
    const afterXml = unzipPart(saved.bytes, 'xl/worksheets/sheet1.xml');
    expect(afterXml).toContain('location="#REF!"');
    expect(afterXml).toContain('ref="B4"');
  });

  it('pins the oracle side of D-9: Excel freezes internal hyperlink locations', () => {
    // The Excel 365 capture of 2026-11-22
    // (scripts/oracle/capture-broken-hyperlink.ps1, evidence in
    // fixtures/oracle/d9-broken-hyperlink/) saved one workbook per case and
    // read the bytes back: across a deleted target sheet, a deleted target
    // name, a never-valid location, a deleted target row, rows inserted above
    // the target and a shifted-up target cell, Excel wrote the authored
    // location verbatim in every file and never wrote `#REF!`. The maintainer
    // ruled the same day that calcplus keeps its marker as the deliberate
    // difference D-9 documents.
    const report = JSON.parse(
      readFileSync(
        new URL(
          '../fixtures/oracle/d9-broken-hyperlink/broken-hyperlink-report.json',
          import.meta.url,
        ),
        'utf8',
      ),
    ) as { cases: { case: string; hyperlinksXml: string }[] };
    const byCase = new Map(report.cases.map((entry) => [entry.case, entry.hyperlinksXml]));
    expect([...byCase.keys()].sort()).toEqual([
      'deleted-cells-shift-up',
      'deleted-name',
      'deleted-row',
      'deleted-sheet',
      'inserted-rows',
      'never-existed',
      'valid',
    ]);
    expect(byCase.get('deleted-sheet')).toContain('location="Target!A1"');
    expect(byCase.get('deleted-name')).toContain('location="GoneName"');
    expect(byCase.get('never-existed')).toContain('location="NoSuchSheet!A1"');
    expect(byCase.get('deleted-row')).toContain('location="Target!A5"');
    // Rows inserted above the target moved it to A7; the location did not follow.
    expect(byCase.get('inserted-rows')).toContain('location="Target!A5"');
    for (const xml of byCase.values()) {
      expect(xml).not.toContain('#REF!');
    }
  });

  // -------------------------------------------------------------------------
  // D-10 patched style records
  // -------------------------------------------------------------------------

  it('pins known deviation D-10: a patched style record carries only what the model describes', () => {
    // A patch copies the modelled parts and not the source `rawXml`, identity
    // matching compares the visual parts and not transport details, and an
    // appended `xf` drops source attributes outside the model (ticket 35;
    // `src/model/style-table.ts` owns the fix).
    const workbook = createEmptyWorkbook();
    const base = {
      ...workbook.styles[0]!,
      rawXml:
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" quotePrefix="1" pivotButton="1"/>',
    };
    const patched = patchedStyleInput(base, { font: { bold: true } });
    expect(patched.rawXml).toBe('');
    expect(patched.font?.bold).toBe(true);

    // Identity compares the visual parts, not the rawXml.
    const styleId = findOrAppendStyle(workbook, patched);
    expect(findOrAppendStyle(workbook, { ...patched, rawXml: '<xf custom="1"/>' })).toBe(styleId);

    // On save the appended record drops the unmodelled `xf` attributes while
    // the untouched source record keeps them.
    const stylesXml = `${HEADER}
<styleSheet xmlns="${NS}">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" quotePrefix="1" pivotButton="1"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;
    const worksheet = `${HEADER}
<worksheet xmlns="${NS}">
  <dimension ref="A1"/>
  <sheetData><row r="1"><c r="A1" s="1"><v>1</v></c></row></sheetData>
</worksheet>`;
    const opened = openWorkbook(buildXlsx(worksheet, { stylesXml }));
    const sheet = opened.workbook.sheets[0]!;
    const session = new DocumentSession(opened.workbook);
    session.dispatch(
      'format.bold',
      { sheetId: sheet.id, range: range(0, 0, 0, 0) },
      contextFor(sheet.id, range(0, 0, 0, 0)),
    );
    // The patched cell's new record carries no source rawXml.
    expect(
      session.workbook.styles.find((style) => style.id === getCell(sheet, 0, 0)?.styleId)?.rawXml,
    ).toBe('');
    const saved = saveWorkbook({
      workbook: session.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
    });
    const savedStyles = unzipPart(saved.bytes, 'xl/styles.xml');
    const cellXfs = savedStyles.slice(
      savedStyles.indexOf('<cellXfs'),
      savedStyles.indexOf('</cellXfs>'),
    );
    const xfs = cellXfs.match(/<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g)!;
    expect(xfs).toHaveLength(3);
    expect(xfs[1]).toContain('quotePrefix="1"');
    expect(xfs[2]).toContain('applyFont="1"');
    expect(xfs[2]).not.toContain('quotePrefix');
    expect(xfs[2]).not.toContain('pivotButton');
    // The patched cell points at the appended record.
    expect(unzipPart(saved.bytes, 'xl/worksheets/sheet1.xml')).toContain('<c r="A1" s="2">');
  });

  // -------------------------------------------------------------------------
  // D-11 rich text
  // -------------------------------------------------------------------------

  it('pins known deviation D-11: rich text is flattened when edited; untouched runs are preserved', () => {
    // Editing a cell with mixed runs keeps the plain text, warns, and loses
    // the runs. An untouched rich cell keeps its runs byte for byte
    // (file-compat §3).
    const runs =
      '<r><rPr><b/><sz val="11"/><rFont val="Calibri"/></rPr><t>Bold</t></r><r><t xml:space="preserve"> plain</t></r>';
    const worksheet = `${HEADER}
<worksheet xmlns="${NS}">
  <dimension ref="A1:B1"/>
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is>${runs}</is></c>
      <c r="B1" t="inlineStr"><is>${runs}</is></c>
    </row>
  </sheetData>
</worksheet>`;
    const original = buildXlsx(worksheet);
    const opened = openWorkbook(original);
    const sheet = opened.workbook.sheets[0]!;
    expect(getCell(sheet, 0, 0)?.richText?.plainText).toBe('Bold plain');
    expect(getCell(sheet, 0, 0)?.richText?.xml).toContain('<rPr><b/>');

    const session = new DocumentSession(opened.workbook);
    const result = session.dispatch('cell.set', {
      sheetId: sheet.id,
      row: 0,
      col: 0,
      value: textValue('Bold plain'),
    });
    expect(result.ui).toMatchObject({ type: 'toast' });
    expect(result.ui?.message).toMatch(/rich text/i);
    expect(result.ui?.message).toMatch(/plain text/i);
    // The edit keeps the plain text and loses the runs.
    expect(getCell(sheet, 0, 0)?.richText).toBeNull();
    expect(getCell(sheet, 0, 0)?.value).toEqual(textValue('Bold plain'));
    // The untouched neighbour keeps its runs.
    expect(getCell(sheet, 0, 1)?.richText?.xml).toContain('<rPr><b/>');

    const saved = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
    });
    expect(unzipPart(saved.bytes, 'xl/worksheets/sheet1.xml')).toContain(
      `<c r="B1" t="inlineStr"><is>${runs}</is></c>`,
    );
    const reopened = openWorkbook(saved.bytes);
    expect(getCell(reopened.workbook.sheets[0]!, 0, 0)?.richText).toBeNull();
    expect(getCell(reopened.workbook.sheets[0]!, 0, 0)?.value).toEqual(textValue('Bold plain'));
    expect(getCell(reopened.workbook.sheets[0]!, 0, 1)?.richText).not.toBeNull();
  });
});
