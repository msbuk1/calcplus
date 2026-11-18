import { describe, expect, it } from 'vitest';
import { openWorkbook } from '@ooxml/open';
import { saveWorkbook } from '@ooxml/save';
import { WorkbookSaveError } from '@ooxml/errors';
import type { ConditionalFormatRule } from '@model/conditional-format';
import {
  createDxfResolver,
  isConditionalFormatRuleSource,
  listUnsupportedConditionalFormats,
  parseConditionalFormats,
  parseDxfTable,
  parseDifferentialFormat,
  serialiseConditionalFormatElement,
  serialiseConditionalFormatRule,
  unsupportedConditionalFormatsInRules,
} from '@ooxml/conditional-format-xml';
import { DocumentSession } from '@commands/session';
import { numberValue } from '@model/value';
import { buildXlsx, elementRanges, unzipPart, type BuildXlsxOptions } from './helpers/build-xlsx';

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const X14_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
const HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

function worksheetPart(bytes: Uint8Array): string {
  return unzipPart(bytes, 'xl/worksheets/sheet1.xml');
}

function stylesPart(bytes: Uint8Array): string {
  return unzipPart(bytes, 'xl/styles.xml');
}

/** Every `conditionalFormatting` element (base and x14) in document order. */
function conditionalFormatSlices(xml: string): string[] {
  return elementRanges(xml, 'conditionalFormatting').map((range) =>
    xml.slice(range.start, range.end),
  );
}

/** The worksheet's `extLst` that carries the x14 conditional formats. */
function x14BlockSlice(xml: string): string | null {
  for (const range of elementRanges(xml, 'extLst')) {
    const slice = xml.slice(range.start, range.end);
    if (slice.includes('x14:conditionalFormattings')) return slice;
  }
  return null;
}

function openFixture(worksheetXml: string, options: BuildXlsxOptions = {}) {
  return openWorkbook(buildXlsx(worksheetXml, options));
}

function saveOpened(
  opened: ReturnType<typeof openWorkbook>,
  dirtyCells: Map<string, Set<string>> = new Map(),
) {
  return saveWorkbook({
    workbook: opened.workbook,
    package: opened.package,
    dirtyCells,
  });
}

function payloadOnly(rule: ConditionalFormatRule): unknown {
  return {
    kind: rule.kind,
    ranges: rule.ranges,
    priority: rule.priority,
    stopIfTrue: rule.stopIfTrue,
    dxfId: rule.dxfId,
    ...('operator' in rule ? { operator: rule.operator } : {}),
    ...('value1' in rule ? { value1: rule.value1 } : {}),
  };
}

// ---------------------------------------------------------------------------
// CF14 cf-unsupported
// ---------------------------------------------------------------------------

const UNSUPPORTED_WORKSHEET = `${HEADER}
<worksheet xmlns="${NS}" xmlns:r="${REL_NS}">
  <dimension ref="A1:A12"/>
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Score</t></is></c></row>
    <row r="2"><c r="A2"><v>5</v></c></row>
    <row r="3"><c r="A3"><v>120</v></c></row>
  </sheetData>
  <conditionalFormatting sqref="A2:A12" data-note="keep-me">
    <cfRule type="expression" dxfId="0" priority="4" customAttr="x"><formula>$A2&gt;100</formula><extLst><ext uri="{AAA}"><foo bar="1"/></ext></extLst></cfRule>
    <cfRule type="aboveAverage" priority="3" aboveAverage="0" equalAverage="1"><formula>1</formula></cfRule>
    <cfRule type="containsBlanks" priority="2"><formula>LEN(TRIM(A2))=0</formula></cfRule>
    <cfRule type="colorScale" priority="1">
      <colorScale><cfvo type="formula" val="$Z$1"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale>
    </cfRule>
  </conditionalFormatting>
</worksheet>`;

describe('CF14 cf-unsupported', () => {
  it('reads formula, above-average, blank and formula-threshold rules as unsupported', () => {
    const opened = openFixture(UNSUPPORTED_WORKSHEET);
    const sheet = opened.workbook.sheets[0]!;
    expect(sheet.conditionalFormats).toHaveLength(1);
    const rules = opened.conditionalFormatRules.get(sheet.id)!;
    expect(rules).toHaveLength(4);

    expect(
      rules.map((rule) =>
        rule.kind === 'unsupported' ? [rule.reason, rule.ruleType] : [rule.kind, null],
      ),
    ).toEqual([
      ['formula', 'expression'],
      ['aboveAverage', 'aboveAverage'],
      ['blankOrError', 'containsBlanks'],
      ['formulaThreshold', 'colorScale'],
    ]);
    for (const rule of rules) {
      expect(rule.ranges).toEqual([{ startRow: 1, startCol: 0, endRow: 11, endCol: 0 }]);
      // §1: preserved, never painted.
      expect(rule.kind).toBe('unsupported');
    }
    expect(rules[0]?.dxfId).toBe(0);
    expect(rules[0]?.priority).toBe(4);
  });

  it('lists them for the compatibility report with sheet, ranges and type', () => {
    const opened = openFixture(UNSUPPORTED_WORKSHEET);
    const listing = listUnsupportedConditionalFormats(opened.workbook);
    expect(
      listing.map((entry) => [entry.sheetName, entry.ruleType, entry.reason, entry.source]),
    ).toEqual([
      ['Sheet1', 'expression', 'formula', 'base'],
      ['Sheet1', 'aboveAverage', 'aboveAverage', 'base'],
      ['Sheet1', 'containsBlanks', 'blankOrError', 'base'],
      ['Sheet1', 'colorScale', 'formulaThreshold', 'base'],
    ]);
    for (const entry of listing) {
      expect(entry.ranges).toEqual([{ startRow: 1, startCol: 0, endRow: 11, endCol: 0 }]);
    }
  });

  it('preserves the rules byte-for-byte on a no-op save and after an edit', () => {
    const original = buildXlsx(UNSUPPORTED_WORKSHEET);
    const opened = openWorkbook(original);
    const noOp = saveOpened(opened);
    expect(noOp.noOp).toBe(true);
    expect(Buffer.compare(Buffer.from(noOp.bytes), Buffer.from(original))).toBe(0);

    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    session.dispatch('cell.set', { sheetId: sheet.id, row: 4, col: 1, value: numberValue(7) });
    const edited = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
    });
    expect(conditionalFormatSlices(worksheetPart(edited.bytes))).toEqual(
      conditionalFormatSlices(worksheetPart(original)),
    );
    const slice = conditionalFormatSlices(worksheetPart(edited.bytes))[0]!;
    expect(slice).toContain('data-note="keep-me"');
    expect(slice).toContain('customAttr="x"');
    expect(slice).toContain('<foo bar="1"/>');
  });

  it('keeps the listing unchanged after a save and reopen', () => {
    const opened = openFixture(UNSUPPORTED_WORKSHEET);
    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    session.dispatch('cell.set', { sheetId: sheet.id, row: 5, col: 0, value: numberValue(9) });
    const saved = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
    });
    const reopened = openWorkbook(saved.bytes);
    expect(
      listUnsupportedConditionalFormats(reopened.workbook).map((entry) => entry.reason),
    ).toEqual(['formula', 'aboveAverage', 'blankOrError', 'formulaThreshold']);
  });
});

// ---------------------------------------------------------------------------
// CF15 cf-x14
// ---------------------------------------------------------------------------

const DATA_BAR_ID = '{11111111-1111-1111-1111-111111111111}';
const ICON_SET_ID = '{22222222-2222-2222-2222-222222222222}';

const X14_LINK_EXT = (guid: string) =>
  `<extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" xmlns:x14="${X14_NS}"><x14:id>${guid}</x14:id></ext></extLst>`;

const X14_WORKSHEET = `${HEADER}
<worksheet xmlns="${NS}" xmlns:r="${REL_NS}">
  <dimension ref="A1:C10"/>
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Label</t></is></c></row>
    <row r="2"><c r="B2"><v>10</v></c><c r="C2"><v>1</v></c></row>
    <row r="3"><c r="B3"><v>20</v></c><c r="C3"><v>2</v></c></row>
  </sheetData>
  <conditionalFormatting sqref="B2:B10">
    <cfRule type="dataBar" priority="1" id="${DATA_BAR_ID}"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar>${X14_LINK_EXT(DATA_BAR_ID)}</cfRule>
  </conditionalFormatting>
  <conditionalFormatting sqref="C2:C10">
    <cfRule type="iconSet" priority="2" id="${ICON_SET_ID}"><iconSet iconSet="3TrafficLights1"><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet>${X14_LINK_EXT(ICON_SET_ID)}</cfRule>
  </conditionalFormatting>
  <extLst>
    <ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="${X14_NS}">
      <x14:conditionalFormattings>
        <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
          <x14:cfRule type="dataBar" id="${DATA_BAR_ID}">
            <x14:dataBar minLength="0" maxLength="100" border="1" gradient="0" showValue="0" direction="rightToLeft" negativeBarColorSameAsPositive="0">
              <x14:cfvo type="autoMin"/>
              <x14:cfvo type="autoMax"/>
              <x14:fillColor rgb="FF4E9A06"/>
              <x14:negativeFillColor rgb="FFCC0000"/>
              <x14:axisColor rgb="FF000000"/>
            </x14:dataBar>
          </x14:cfRule>
        </x14:conditionalFormatting>
        <x14:conditionalFormatting>
          <x14:cfRule type="iconSet" id="${ICON_SET_ID}">
            <x14:iconSet iconSet="4Rating" showValue="0" reverse="1">
              <x14:cfvo type="percent" val="0"/>
              <x14:cfvo type="percent" val="25"/>
              <x14:cfvo type="percent" val="50"/>
              <x14:cfvo type="percent" val="75"/>
            </x14:iconSet>
          </x14:cfRule>
        </x14:conditionalFormatting>
      </x14:conditionalFormattings>
    </ext>
  </extLst>
</worksheet>`;

describe('CF15 cf-x14', () => {
  it('merges x14 data bars and icon sets with their base rules', () => {
    const opened = openFixture(X14_WORKSHEET);
    const sheet = opened.workbook.sheets[0]!;
    const rules = opened.conditionalFormatRules.get(sheet.id)!;
    // One merged rule per base element: the x14 copy is not a second rule.
    expect(rules).toHaveLength(2);
    expect(rules.map((rule) => rule.id)).toEqual([DATA_BAR_ID, ICON_SET_ID]);

    const dataBar = rules[0]!;
    expect(dataBar.kind).toBe('dataBar');
    expect(dataBar.ranges).toEqual([{ startRow: 1, startCol: 1, endRow: 9, endCol: 1 }]);
    if (dataBar.kind !== 'dataBar') throw new Error('expected a data bar');
    expect(dataBar).toMatchObject({
      min: { type: 'automatic' },
      max: { type: 'automatic' },
      // The x14 payload drives rendering (§7).
      color: '#4E9A06',
      negativeColor: '#CC0000',
      axisColor: '#000000',
      fill: 'solid',
      border: true,
      showValue: false,
      direction: 'rightToLeft',
    });

    const iconSet = rules[1]!;
    expect(iconSet.kind).toBe('iconSet');
    expect(iconSet.ranges).toEqual([{ startRow: 1, startCol: 2, endRow: 9, endCol: 2 }]);
    if (iconSet.kind !== 'iconSet') throw new Error('expected an icon set');
    expect(iconSet).toMatchObject({
      iconSet: '4Rating',
      reverse: true,
      iconOnly: true,
      thresholds: [
        { type: 'percent', value: 0 },
        { type: 'percent', value: 25 },
        { type: 'percent', value: 50 },
        { type: 'percent', value: 75 },
      ],
    });
  });

  it('keeps both XML forms on the rule raw payload', () => {
    const opened = openFixture(X14_WORKSHEET);
    const sheet = opened.workbook.sheets[0]!;
    const rules = opened.conditionalFormatRules.get(sheet.id)!;
    for (const rule of rules) {
      expect(isConditionalFormatRuleSource(rule.raw)).toBe(true);
      if (!isConditionalFormatRuleSource(rule.raw)) continue;
      expect(rule.raw.base).not.toBeNull();
      expect(rule.raw.x14).not.toBeNull();
      expect(rule.raw.base?.source).toBe('base');
      expect(rule.raw.x14?.source).toBe('x14');
      expect(rule.raw.base?.xml.startsWith('<cfRule')).toBe(true);
      expect(rule.raw.x14?.xml).toContain('<x14:cfRule');
    }
  });

  it('leaves both XML forms byte-identical on a no-op save and after an edit', () => {
    const original = buildXlsx(X14_WORKSHEET);
    const opened = openWorkbook(original);
    const noOp = saveOpened(opened);
    expect(noOp.noOp).toBe(true);
    expect(Buffer.compare(Buffer.from(noOp.bytes), Buffer.from(original))).toBe(0);

    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    session.dispatch('cell.set', { sheetId: sheet.id, row: 4, col: 0, value: numberValue(3) });
    const edited = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
    });
    const before = worksheetPart(original);
    const after = worksheetPart(edited.bytes);
    expect(conditionalFormatSlices(after)).toEqual(conditionalFormatSlices(before));
    expect(x14BlockSlice(after)).toBe(x14BlockSlice(before));
    expect(x14BlockSlice(after)).not.toBeNull();
  });

  it('falls back to same priority and family when the file carries no rule ids', () => {
    const worksheet = X14_WORKSHEET.replaceAll(` id="${DATA_BAR_ID}"`, '')
      .replaceAll(` id="${ICON_SET_ID}"`, '')
      .replace(
        `<extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" xmlns:x14="${X14_NS}"><x14:id>${DATA_BAR_ID}</x14:id></ext></extLst>`,
        '',
      )
      .replace(
        `<extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" xmlns:x14="${X14_NS}"><x14:id>${ICON_SET_ID}</x14:id></ext></extLst>`,
        '',
      );
    const opened = openFixture(worksheet);
    const sheet = opened.workbook.sheets[0]!;
    const rules = opened.conditionalFormatRules.get(sheet.id)!;
    expect(rules).toHaveLength(2);
    expect(rules.map((rule) => rule.kind)).toEqual(['dataBar', 'iconSet']);
    // Without ids the id is positional, but the x14 payload still wins.
    if (rules[0]?.kind === 'dataBar') expect(rules[0].fill).toBe('solid');
  });

  it('transforms the base sqref once and leaves the x14 block byte-identical', () => {
    const original = buildXlsx(X14_WORKSHEET);
    const opened = openWorkbook(original);
    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    const selection = { startRow: 1, startCol: 0, endRow: 1, endCol: 2 };
    session.dispatch(
      'structure.insertRows',
      {},
      {
        activeSheetId: sheet.id,
        selection,
        activeCell: { row: 1, col: 0 },
      },
    );
    const result = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
      dirtyDimensions: session.dirtyDimensionSnapshot(),
      dirtyMerges: new Set(session.dirtyMergeSnapshot().keys()),
    });
    const before = worksheetPart(original);
    const after = worksheetPart(result.bytes);
    expect(after).toContain('<conditionalFormatting sqref="B3:B11">');
    expect(after).toContain('<conditionalFormatting sqref="C3:C11">');
    expect(x14BlockSlice(after)).toBe(x14BlockSlice(before));

    const reopened = openWorkbook(result.bytes);
    const reopenedSheet = reopened.workbook.sheets[0]!;
    const rules = reopened.conditionalFormatRules.get(reopenedSheet.id)!;
    expect(rules.map((rule) => rule.id)).toEqual([DATA_BAR_ID, ICON_SET_ID]);
    expect(rules[0]?.ranges).toEqual([{ startRow: 2, startCol: 1, endRow: 10, endCol: 1 }]);
    // The x14 payload still drives rendering after the move.
    if (rules[0]?.kind === 'dataBar') expect(rules[0].fill).toBe('solid');
  });
});

// ---------------------------------------------------------------------------
// CF16 cf-dxf-roundtrip
// ---------------------------------------------------------------------------

const STYLES_XML = `${HEADER}
<styleSheet xmlns="${NS}">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
  <dxfs count="2">
    <dxf><font><b/><color rgb="FFFF0000"/></font><fill><patternFill><bgColor rgb="FFFFFF00"/></patternFill></fill></dxf>
    <dxf><numFmt numFmtId="164" formatCode="0.00%"/></dxf>
  </dxfs>
</styleSheet>`;

const DXF_WORKSHEET = `${HEADER}
<worksheet xmlns="${NS}" xmlns:r="${REL_NS}">
  <dimension ref="A1:C10"/>
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c></row>
    <row r="2"><c r="A2"><v>5</v></c><c r="B2"><v>1</v></c></row>
    <row r="3"><c r="A3"><v>7</v></c><c r="B3"><v>1</v></c><c r="C3"><v>2</v></c></row>
    <row r="4"><c r="A4"><v>7</v></c><c r="B4"><v>3</v></c><c r="C4"><v>2</v></c></row>
    <row r="5"><c r="A5"><v>9</v></c><c r="B5"><v>4</v></c><c r="C5"><v>5</v></c></row>
  </sheetData>
  <conditionalFormatting sqref="A2:A10" data-note="keep-me">
    <cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan" stopIfTrue="1" pivot="1" customAttr="x"><formula>6</formula><extLst><ext uri="{AAA}"><foo bar="1"/></ext></extLst></cfRule>
    <cfRule type="top10" dxfId="1" priority="3" rank="5" percent="1" top="0"/>
  </conditionalFormatting>
  <conditionalFormatting sqref="B2:B4 C2:C5">
    <cfRule type="duplicateValues" priority="2" dxfId="0"/>
  </conditionalFormatting>
</worksheet>`;

const DXF_OPTIONS: BuildXlsxOptions = { stylesXml: STYLES_XML };

describe('CF16 cf-dxf-roundtrip', () => {
  it('parses the dxfs table by index and resolves the four merge properties', () => {
    const opened = openFixture(DXF_WORKSHEET, DXF_OPTIONS);
    const table = parseDxfTable(opened.package.entries.get('xl/styles.xml')!, 'xl/styles.xml');
    expect(table.entries.map((entry) => entry.id)).toEqual([0, 1]);
    expect(table.declaredCount).toBe(2);
    expect(table.entries[0]!.rawXml).toContain(
      '<patternFill><bgColor rgb="FFFFFF00"/></patternFill>',
    );

    const resolver = createDxfResolver(table);
    expect(resolver(0)).toMatchObject({
      font: { bold: true, color: '#FF0000' },
      fill: { backgroundColor: '#FFFF00' },
    });
    expect(resolver(1)).toMatchObject({ numberFormat: { numFmtId: 164, code: '0.00%' } });
    expect(resolver(2)).toBeNull();

    const parsed = parseDifferentialFormat(table.entries[1]!.rawXml, 'xl/styles.xml');
    expect(parsed.numberFormat).toEqual({ numFmtId: 164, code: '0.00%' });
    expect(parsed.fill).toBeUndefined();
  });

  it('reads the rules with their stable dxf ids, priorities and multi-area sqref', () => {
    const opened = openFixture(DXF_WORKSHEET, DXF_OPTIONS);
    const sheet = opened.workbook.sheets[0]!;
    expect(sheet.conditionalFormats).toHaveLength(2);
    expect(sheet.conditionalFormats[1]!.range).toEqual({
      startRow: 1,
      startCol: 1,
      endRow: 4,
      endCol: 2,
    });

    const rules = opened.conditionalFormatRules.get(sheet.id)!;
    expect(rules.map((rule) => rule.dxfId)).toEqual([0, 1, 0]);
    expect(rules.map((rule) => rule.priority)).toEqual([1, 3, 2]);
    expect(rules[0]?.stopIfTrue).toBe(true);
    expect(rules[1]).toMatchObject({ kind: 'top10', rank: 5, top: false, percent: true });
    expect(rules[2]?.ranges).toEqual([
      { startRow: 1, startCol: 1, endRow: 3, endCol: 1 },
      { startRow: 1, startCol: 2, endRow: 4, endCol: 2 },
    ]);
    expect(unsupportedConditionalFormatsInRules(sheet, rules)).toEqual([]);
  });

  it('leaves the whole file byte-identical on a no-op save', () => {
    const original = buildXlsx(DXF_WORKSHEET, DXF_OPTIONS);
    const opened = openWorkbook(original);
    const noOp = saveOpened(opened);
    expect(noOp.noOp).toBe(true);
    expect(Buffer.compare(Buffer.from(noOp.bytes), Buffer.from(original))).toBe(0);
    expect(stylesPart(noOp.bytes)).toBe(STYLES_XML);
  });

  it('keeps rule XML verbatim, dxf ids stable and the styles part untouched after an edit', () => {
    const original = buildXlsx(DXF_WORKSHEET, DXF_OPTIONS);
    const opened = openWorkbook(original);
    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    session.dispatch('cell.set', { sheetId: sheet.id, row: 9, col: 3, value: numberValue(42) });
    const edited = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
    });
    const after = worksheetPart(edited.bytes);
    const before = worksheetPart(original);
    expect(conditionalFormatSlices(after)).toEqual(conditionalFormatSlices(before));

    // Unknown attributes and children survive in their authored order.
    const slice = conditionalFormatSlices(after)[0]!;
    expect(slice).toContain(
      '<cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan" stopIfTrue="1" pivot="1" customAttr="x"><formula>6</formula><extLst><ext uri="{AAA}"><foo bar="1"/></ext></extLst></cfRule>',
    );
    expect(slice).toContain('data-note="keep-me"');

    // dxf ids are never renumbered and the dxfs table never moves.
    const table = parseDxfTable(edited.package.entries.get('xl/styles.xml')!, 'xl/styles.xml');
    expect(table.entries.map((entry) => entry.id)).toEqual([0, 1]);
    expect(stylesPart(edited.bytes)).toBe(STYLES_XML);

    const reopened = openWorkbook(edited.bytes).workbook.sheets[0]!;
    expect(openWorkbook(edited.bytes).conditionalFormatRules.get(reopened.id)).toHaveLength(3);
  });

  it('writes a removed rule without renumbering the remaining dxf ids', () => {
    const opened = openFixture(DXF_WORKSHEET, DXF_OPTIONS);
    const sheet = opened.workbook.sheets[0]!;
    // Drop the second element: the remaining rules still index dxf 0 and 1.
    sheet.conditionalFormats.splice(1, 1);
    const result = saveOpened(opened);
    expect(result.noOp).toBe(false);
    const slices = conditionalFormatSlices(worksheetPart(result.bytes));
    expect(slices).toHaveLength(1);
    expect(slices[0]).toBe(
      conditionalFormatSlices(worksheetPart(buildXlsx(DXF_WORKSHEET, DXF_OPTIONS)))[0],
    );
    expect(stylesPart(result.bytes)).toBe(STYLES_XML);

    const reopened = openWorkbook(result.bytes);
    const reopenedSheet = reopened.workbook.sheets[0]!;
    expect(reopenedSheet.conditionalFormats).toHaveLength(1);
    expect(
      reopened.conditionalFormatRules.get(reopenedSheet.id)!.map((rule) => rule.dxfId),
    ).toEqual([0, 1]);
  });

  it('inserts a new rule element before the later worksheet children', () => {
    const opened = openFixture(DXF_WORKSHEET, DXF_OPTIONS);
    const sheet = opened.workbook.sheets[0]!;
    sheet.conditionalFormats.push({
      range: { startRow: 1, startCol: 0, endRow: 3, endCol: 0 },
      rawXml:
        '<conditionalFormatting sqref="A2:A4"><cfRule type="duplicateValues" priority="9" dxfId="0"/></conditionalFormatting>',
    });
    const result = saveOpened(opened);
    const part = worksheetPart(result.bytes);
    expect(conditionalFormatSlices(part)).toHaveLength(3);
    expect(part.indexOf('sqref="A2:A4"')).toBeGreaterThan(part.indexOf('sqref="B2:B4 C2:C5"'));
    expect(part.indexOf('sqref="A2:A4"')).toBeLessThan(part.indexOf('</worksheet>'));
    expect(openWorkbook(result.bytes).workbook.sheets[0]!.conditionalFormats).toHaveLength(3);
  });

  it('rewrites only the sqref and changed attributes when a record is edited', () => {
    const opened = openFixture(DXF_WORKSHEET, DXF_OPTIONS);
    const sheet = opened.workbook.sheets[0]!;
    const parsed = parseConditionalFormats(worksheetPart(opened.package.originalBytes!), sheet);
    const element = parsed.elements[0]!;
    const rule = element.rules[0]!;

    // State-key optimisation: untouched source is re-emitted verbatim.
    expect(serialiseConditionalFormatElement(element.element, element.rules)).toBe(
      element.element.xml,
    );
    expect(serialiseConditionalFormatRule(rule)).toBe(
      isConditionalFormatRuleSource(rule.raw) ? rule.raw.base!.xml : '',
    );

    // An attribute-only change keeps every unknown attribute and child.
    const retargeted = { ...rule, dxfId: 1 } as ConditionalFormatRule;
    expect(serialiseConditionalFormatRule(retargeted)).toContain(
      '<cfRule type="cellIs" dxfId="1" priority="1" operator="greaterThan" stopIfTrue="1" pivot="1" customAttr="x"><formula>6</formula><extLst><ext uri="{AAA}"><foo bar="1"/></ext></extLst></cfRule>',
    );

    // A moved range rewrites sqref, nothing else.
    const moved = {
      ...element.element,
      ranges: [{ startRow: 9, startCol: 0, endRow: 19, endCol: 0 }],
    };
    const serialised = serialiseConditionalFormatElement(moved, element.rules);
    expect(
      serialised.startsWith('<conditionalFormatting sqref="A10:A20" data-note="keep-me">'),
    ).toBe(true);
    expect(serialised).toContain('<foo bar="1"/>');

    // Rule authoring is v2: a payload change is refused, not guessed.
    const reauthored = { ...rule, operator: 'lessThan' } as ConditionalFormatRule;
    expect(() => serialiseConditionalFormatRule(reauthored)).toThrow(WorkbookSaveError);
  });

  it('keeps the parser output stable when re-run on the same part', () => {
    const opened = openFixture(DXF_WORKSHEET, DXF_OPTIONS);
    const sheet = opened.workbook.sheets[0]!;
    const part = worksheetPart(opened.package.originalBytes!);
    const first = parseConditionalFormats(part, sheet);
    const second = parseConditionalFormats(part, sheet);
    expect(first.records).toEqual(second.records);
    expect(first.rules.map(payloadOnly)).toEqual(second.rules.map(payloadOnly));
  });
});
