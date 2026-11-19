import { writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { unzipSync, zipSync } from 'fflate';
import { XML_HEADER, SPREADSHEET_NS, REL_NS, PACKAGE_REL_NS, packageEntries } from '../parts.mjs';

/**
 * M6 lane 3 fixtures: sheet and content features.
 *
 * - F12 `threaded-comments.xlsx` — legacy note, threaded comments, persons.
 * - F13 `validation-cf.xlsx` — every supported validation type and every
 *   supported conditional-format rule type, plus internal and external links.
 * - F14 `protection-hidden.xlsx` — workbook and sheet protection, a hidden and
 *   a veryHidden sheet, a hidden row and a hidden column.
 * - F16 `print-setup.xlsx` — print area, orientation, fit-to-width and margins.
 * - F17 `rich-text.xlsx` — rich shared-string and inline runs.
 * - F23 `names-defined.xlsx` — workbook and sheet names with quoted references.
 *
 * Every fixture is hand-authored OOXML with fixed zip timestamps, so
 * `npm run fixtures` is byte-identical on every run (groups/README.md). The
 * assertions live in tests/m6-sheet-features.test.ts.
 */

const DAY_MS = 86_400_000;

/** 1900 date system serial, matching Excel's epoch. */
function serial1900(year, month, day, fraction = 0) {
  return Math.floor((Date.UTC(year, month - 1, day) - Date.UTC(1899, 11, 30)) / DAY_MS) + fraction;
}

const CONTENT_TYPE = {
  comments: 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml',
  threadedComments: 'application/vnd.ms-excel.threadedcomments+xml',
  persons: 'application/vnd.ms-excel.person+xml',
};

/**
 * Give comment parts their registered content types. `packageEntries` leaves an
 * unknown extra part at `application/xml`; OPC allows one override per part, so
 * an existing override is replaced rather than duplicated.
 */
function withContentTypes(bytes, overrides) {
  if (overrides.length === 0) return bytes;
  const entries = unzipSync(bytes);
  let contentTypes = new TextDecoder().decode(entries['[Content_Types].xml']);
  for (const [part, type] of overrides) {
    const replacement = `<Override PartName="/${part}" ContentType="${type}"/>`;
    const pattern = new RegExp(`<Override PartName="/${escapeRegExp(part)}"[^>]*/>`);
    contentTypes = pattern.test(contentTypes)
      ? contentTypes.replace(pattern, replacement)
      : contentTypes.replace('</Types>', `  ${replacement}\n</Types>`);
  }
  entries['[Content_Types].xml'] = new TextEncoder().encode(contentTypes);
  return zipSync(entries, { level: 6, mtime: new Date('1980-01-01T00:00:00Z') });
}

function escapeRegExp(text) {
  return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function buildPackage({ workbook, workbookRels, sheets, sharedStrings, extra = {}, overrides = [] }) {
  return withContentTypes(
    packageEntries({ workbook, workbookRels, sheets, sharedStrings, extra }),
    overrides,
  );
}

function workbookXml({ sheets, definedNames = '', workbookProtection = '', calcPr = '' }) {
  const sheetList = sheets
    .map(
      ([name, state], index) =>
        `    <sheet name="${escapeAttribute(name)}" sheetId="${index + 1}"${
          state === 'visible' ? '' : ` state="${state}"`
        } r:id="rId${index + 1}"/>`,
    )
    .join('\n');
  return `${XML_HEADER}
<workbook xmlns="${SPREADSHEET_NS}" xmlns:r="${REL_NS}">
  <workbookPr date1904="0"/>
  ${workbookProtection}
  <bookViews><workbookView activeTab="0"/></bookViews>
  <sheets>
${sheetList}
  </sheets>
  ${definedNames}
  ${calcPr === '' ? '<calcPr calcId="191029" fullCalcOnLoad="0"/>' : calcPr}
</workbook>`;
}

function workbookRelsXml(entries) {
  return `${XML_HEADER}
<Relationships xmlns="${PACKAGE_REL_NS}">
${entries
  .map(([id, type, target, mode]) => {
    const targetMode = mode === undefined ? '' : ` TargetMode="${mode}"`;
    return `  <Relationship Id="${id}" Type="${type}" Target="${target}"${targetMode}/>`;
  })
  .join('\n')}
</Relationships>`;
}

function sheetRelsXml(entries) {
  return `${XML_HEADER}
<Relationships xmlns="${PACKAGE_REL_NS}">
${entries
  .map(([id, type, target, mode]) => {
    const targetMode = mode === undefined ? '' : ` TargetMode="${mode}"`;
    return `  <Relationship Id="${id}" Type="${type}" Target="${target}"${targetMode}/>`;
  })
  .join('\n')}
</Relationships>`;
}

function escapeAttribute(text) {
  return text
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;');
}

/** The minimal style sheet the fixtures share, with optional dxfs. */
function stylesXml(dxfs) {
  return `${XML_HEADER}
<styleSheet xmlns="${SPREADSHEET_NS}">
  <fonts count="2"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
  ${dxfs}
</styleSheet>`;
}

const TWO_STYLES = stylesXml('');

// ---------------------------------------------------------------------------
// F12 threaded-comments.xlsx — threaded comments and persons preserved; the
// comment indicator renders (file-compat §3, §6.6, §9 F12).
// ---------------------------------------------------------------------------

const f12Workbook = workbookXml({ sheets: [['Discussions', 'visible']] });

const f12WorkbookRels = workbookRelsXml([
  ['rId1', `${REL_NS}/worksheet`, 'worksheets/sheet1.xml'],
  ['rId2', `${REL_NS}/styles`, 'styles.xml'],
  ['rId3', 'http://schemas.microsoft.com/office/2017/10/relationships/person', 'persons/person.xml'],
]);

const f12SheetRels = sheetRelsXml([
  ['rId1', `${REL_NS}/comments`, '../comments1.xml'],
  ['rId2', `${REL_NS}/vmlDrawing`, '../drawings/vmlDrawing1.vml'],
  [
    'rId3',
    'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment',
    '../threadedComments/threadedComment1.xml',
  ],
]);

const f12Sheet = `${XML_HEADER}
<worksheet xmlns="${SPREADSHEET_NS}" xmlns:r="${REL_NS}">
  <dimension ref="A1:E4"/>
  <sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Topic</t></is></c><c r="B1" t="inlineStr"><is><t>Owner</t></is></c><c r="C1" t="inlineStr"><is><t>Status</t></is></c><c r="D1" t="inlineStr"><is><t>Note text</t></is></c><c r="E1" t="inlineStr"><is><t>Follow-up</t></is></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>Budget</t></is></c><c r="B2" t="inlineStr"><is><t>Ana</t></is></c><c r="C2" t="inlineStr"><is><t>open</t></is></c><c r="D2" t="inlineStr"><is><t>legacy note cell</t></is></c></row>
    <row r="3"><c r="A3" t="inlineStr"><is><t>Hiring</t></is></c><c r="B3" t="inlineStr"><is><t>Ana</t></is></c><c r="C3" t="inlineStr"><is><t>discussed</t></is></c><c r="D3" t="inlineStr"><is><t>thread root cell</t></is></c></row>
    <row r="4"><c r="A4" t="inlineStr"><is><t>Launch</t></is></c><c r="B4" t="inlineStr"><is><t>Bo</t></is></c><c r="C4" t="inlineStr"><is><t>blocked</t></is></c><c r="D4" t="inlineStr"><is><t>reply cell</t></is></c></row>
  </sheetData>
  <legacyDrawing r:id="rId2"/>
</worksheet>`;

const f12Comments = `${XML_HEADER}
<comments xmlns="${SPREADSHEET_NS}">
  <authors><author>Ana Silva</author></authors>
  <commentList>
    <comment ref="D2" authorId="0"><text><r><rPr><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><family val="2"/></rPr><t>Legacy note: confirm the budget.</t></r></text></comment>
  </commentList>
</comments>`;

const f12ThreadedComments = `${XML_HEADER}
<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <threadedComment ref="D3" dT="2026-11-18T09:15:00Z" personId="{11111111-1111-1111-1111-111111111111}" id="{22222222-2222-2222-2222-222222222222}">
    <text>Threaded: two candidates shortlisted.</text>
  </threadedComment>
  <threadedComment ref="D4" dT="2026-11-18T10:02:00Z" personId="{33333333-3333-3333-3333-333333333333}" id="{44444444-4444-4444-4444-444444444444}" parentId="{22222222-2222-2222-2222-222222222222}">
    <text>Reply: the launch date is blocked.</text>
  </threadedComment>
  <threadedComment ref="E2" dT="2026-11-18T11:30:00Z" personId="{33333333-3333-3333-3333-333333333333}" id="{55555555-5555-5555-5555-555555555555}" done="1">
    <text>Follow-up on an empty cell.</text>
  </threadedComment>
</ThreadedComments>`;

const f12Persons = `${XML_HEADER}
<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <person displayName="Ana Silva" id="{11111111-1111-1111-1111-111111111111}" userId="ana@example.invalid" providerId="None"/>
  <person displayName="Bo Chen" id="{33333333-3333-3333-3333-333333333333}" userId="bo@example.invalid" providerId="None"/>
</personList>`;

const f12Vml = `${XML_HEADER}
<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
  <o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>
  <v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>
  <v:shape id="_x0000_s1025" type="#_x0000_t202" style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:96pt;height:55.5pt;z-index:1;visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto">
    <v:fill color2="#ffffe1"/>
    <v:shadow on="t" color="black" obscured="t"/>
    <v:path o:connecttype="none"/>
    <v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>
    <x:ClientData ObjectType="Note">
      <x:MoveWithCells/>
      <x:SizeWithCells/>
      <x:Anchor>3, 15, 1, 2, 6, 31, 4, 13</x:Anchor>
      <x:AutoFill>False</x:AutoFill>
      <x:Row>1</x:Row>
      <x:Column>3</x:Column>
    </x:ClientData>
  </v:shape>
</xml>`;

function f12Package() {
  return buildPackage({
    workbook: f12Workbook,
    workbookRels: f12WorkbookRels,
    sheets: [['xl/worksheets/sheet1.xml', f12Sheet]],
    extra: {
      'xl/styles.xml': TWO_STYLES,
      'xl/worksheets/_rels/sheet1.xml.rels': f12SheetRels,
      'xl/comments1.xml': f12Comments,
      'xl/drawings/vmlDrawing1.vml': f12Vml,
      'xl/threadedComments/threadedComment1.xml': f12ThreadedComments,
      'xl/persons/person.xml': f12Persons,
    },
    overrides: [
      ['xl/comments1.xml', CONTENT_TYPE.comments],
      ['xl/threadedComments/threadedComment1.xml', CONTENT_TYPE.threadedComments],
      ['xl/persons/person.xml', CONTENT_TYPE.persons],
    ],
  });
}

// ---------------------------------------------------------------------------
// F13 validation-cf.xlsx — every supported validation type and every supported
// conditional-format rule type (file-compat §9 F13).
// ---------------------------------------------------------------------------

const f13Styles = stylesXml(`<dxfs count="3">
    <dxf><font><color rgb="FF9C0006"/></font></dxf>
    <dxf><font><b/><color rgb="FF1F4E78"/></font></dxf>
    <dxf><fill><patternFill patternType="solid"><bgColor rgb="FFC6EFCE"/></patternFill></fill></dxf>
  </dxfs>`);

const f13Sheet = `${XML_HEADER}
<worksheet xmlns="${SPREADSHEET_NS}" xmlns:r="${REL_NS}">
  <dimension ref="A1:J10"/>
  <sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Text</t></is></c><c r="B1" t="inlineStr"><is><t>CellIs</t></is></c><c r="C1" t="inlineStr"><is><t>Top10</t></is></c><c r="D1" t="inlineStr"><is><t>Duplicate</t></is></c><c r="E1" t="inlineStr"><is><t>Unique</t></is></c><c r="F1" t="inlineStr"><is><t>Today</t></is></c><c r="G1" t="inlineStr"><is><t>Scale</t></is></c><c r="H1" t="inlineStr"><is><t>Bar</t></is></c><c r="I1" t="inlineStr"><is><t>Icons</t></is></c><c r="J1" t="inlineStr"><is><t>Formula</t></is></c>
    </row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>urgent call</t></is></c><c r="B2"><v>50</v></c><c r="C2"><v>10</v></c><c r="D2"><v>7</v></c><c r="E2"><v>1</v></c><c r="F2"><v>${serial1900(2026, 11, 19)}</v></c><c r="G2"><v>10</v></c><c r="H2"><v>20</v></c><c r="I2"><v>20</v></c><c r="J2"><v>1</v></c></row>
    <row r="3"><c r="A3" t="inlineStr"><is><t>normal</t></is></c><c r="B3"><v>150</v></c><c r="C3"><v>40</v></c><c r="D3"><v>7</v></c><c r="E3"><v>2</v></c><c r="F3"><v>${serial1900(2026, 11, 19)}</v></c><c r="G3"><v>50</v></c><c r="H3"><v>50</v></c><c r="I3"><v>50</v></c><c r="J3"><v>2</v></c></row>
    <row r="4"><c r="A4" t="inlineStr"><is><t>URGENT issue</t></is></c><c r="B4"><v>20</v></c><c r="C4"><v>20</v></c><c r="D4"><v>8</v></c><c r="E4"><v>3</v></c><c r="F4"><v>${serial1900(2026, 11, 18)}</v></c><c r="G4"><v>90</v></c><c r="H4"><v>80</v></c><c r="I4"><v>80</v></c><c r="J4"><v>3</v></c></row>
    <row r="5"><c r="A5" t="inlineStr"><is><t>quiet</t></is></c><c r="B5"><v>90</v></c><c r="C5"><v>30</v></c><c r="D5"><v>8</v></c><c r="E5"><v>4</v></c><c r="F5"><v>${serial1900(2026, 11, 20)}</v></c><c r="G5"><v>30</v></c><c r="H5"><v>5</v></c><c r="I5"><v>5</v></c><c r="J5"><v>4</v></c></row>
    <row r="6"><c r="A6" t="inlineStr"><is><t>urgent follow-up</t></is></c><c r="B6"><v>200</v></c><c r="C6"><v>5</v></c><c r="D6"><v>8</v></c><c r="E6"><v>5</v></c><c r="F6"><v>${serial1900(2026, 11, 19)}</v></c><c r="G6"><v>70</v></c><c r="H6"><v>95</v></c><c r="I6"><v>95</v></c><c r="J6"><v>5</v></c></row>
    <row r="8"><c r="A8" t="inlineStr"><is><t>External link</t></is></c><c r="B8" t="inlineStr"><is><t>Internal link</t></is></c></row>
    <row r="9"><c r="A9" t="inlineStr"><is><t>List source</t></is></c><c r="B9" t="inlineStr"><is><t>Alpha</t></is></c><c r="C9" t="inlineStr"><is><t>Beta</t></is></c><c r="D9" t="inlineStr"><is><t>Gamma</t></is></c></row>
  </sheetData>
  <conditionalFormatting sqref="B2:B6"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>100</formula></cfRule></conditionalFormatting>
  <conditionalFormatting sqref="A2:A6"><cfRule type="containsText" dxfId="1" priority="2" operator="containsText" text="urgent"><formula>NOT(ISERROR(SEARCH("urgent",A2)))</formula></cfRule></conditionalFormatting>
  <conditionalFormatting sqref="C2:C6"><cfRule type="top10" dxfId="1" priority="3" rank="2"/></conditionalFormatting>
  <conditionalFormatting sqref="D2:D6"><cfRule type="duplicateValues" dxfId="2" priority="4"/></conditionalFormatting>
  <conditionalFormatting sqref="E2:E6"><cfRule type="uniqueValues" dxfId="2" priority="5"/></conditionalFormatting>
  <conditionalFormatting sqref="F2:F6"><cfRule type="timePeriod" dxfId="1" priority="6" timePeriod="today"><formula>FLOOR(F2,1)=TODAY()</formula></cfRule></conditionalFormatting>
  <conditionalFormatting sqref="G2:G6"><cfRule type="colorScale" priority="7"><colorScale><cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FFFFEB84"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>
  <conditionalFormatting sqref="H2:H6"><cfRule type="dataBar" priority="8"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar></cfRule></conditionalFormatting>
  <conditionalFormatting sqref="I2:I6"><cfRule type="iconSet" priority="9"><iconSet iconSet="3TrafficLights1"><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet></cfRule></conditionalFormatting>
  <conditionalFormatting sqref="J2:J6"><cfRule type="expression" dxfId="0" priority="10"><formula>$J2&gt;2</formula></cfRule></conditionalFormatting>
  <dataValidations count="10">
    <dataValidation type="whole" operator="between" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B2:B6"><formula1>0</formula1><formula2>1000</formula2></dataValidation>
    <dataValidation type="decimal" operator="greaterThan" allowBlank="1" sqref="C2:C6"><formula1>0</formula1></dataValidation>
    <dataValidation type="list" allowBlank="1" showDropDown="0" sqref="D2:D6"><formula1>"Alpha,Beta,Gamma"</formula1></dataValidation>
    <dataValidation type="list" allowBlank="1" sqref="E2:E6"><formula1>$B$9:$D$9</formula1></dataValidation>
    <dataValidation type="date" operator="between" allowBlank="1" sqref="F2:F6"><formula1>${serial1900(2026, 1, 1)}</formula1><formula2>${serial1900(2026, 12, 31)}</formula2></dataValidation>
    <dataValidation type="time" operator="lessThan" allowBlank="1" sqref="G2:G6"><formula1>0.75</formula1></dataValidation>
    <dataValidation type="textLength" operator="equal" allowBlank="1" sqref="A2:A6"><formula1>6</formula1></dataValidation>
    <dataValidation type="custom" allowBlank="1" sqref="J2:J6"><formula1>AND(ISNUMBER(J2),J2&gt;=0)</formula1></dataValidation>
    <dataValidation type="none" allowBlank="1" showInputMessage="1" promptTitle="Hint" prompt="Enter a value" sqref="I2:I6"><formula1/></dataValidation>
    <dataValidation type="whole" operator="notBetween" allowBlank="1" sqref="A8 B8"><formula1>1</formula1><formula2>2</formula2></dataValidation>
  </dataValidations>
  <hyperlinks>
    <hyperlink ref="A8" r:id="rId1" tooltip="External report"/>
    <hyperlink ref="B8" location="Rules!A1" display="Rules sheet"/>
  </hyperlinks>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>`;

const f13SheetRels = sheetRelsXml([
  ['rId1', `${REL_NS}/hyperlink`, 'https://example.invalid/report', 'External'],
]);

function f13Package() {
  return buildPackage({
    workbook: workbookXml({ sheets: [['Rules', 'visible']] }),
    workbookRels: workbookRelsXml([
      ['rId1', `${REL_NS}/worksheet`, 'worksheets/sheet1.xml'],
      ['rId2', `${REL_NS}/styles`, 'styles.xml'],
    ]),
    sheets: [['xl/worksheets/sheet1.xml', f13Sheet]],
    extra: {
      'xl/styles.xml': f13Styles,
      'xl/worksheets/_rels/sheet1.xml.rels': f13SheetRels,
    },
  });
}

// ---------------------------------------------------------------------------
// F14 protection-hidden.xlsx — protection, hidden and veryHidden sheets, a
// hidden row and column are preserved, not enforced (file-compat §6.7, §9 F14).
// ---------------------------------------------------------------------------

const f14Workbook = `${XML_HEADER}
<workbook xmlns="${SPREADSHEET_NS}" xmlns:r="${REL_NS}">
  <workbookPr date1904="0"/>
  <workbookProtection workbookPassword="CC1A" lockStructure="1" lockWindows="1"/>
  <bookViews><workbookView activeTab="0"/></bookViews>
  <sheets>
    <sheet name="Board" sheetId="1" r:id="rId1"/>
    <sheet name="Members" sheetId="2" state="hidden" r:id="rId2"/>
    <sheet name="Archive" sheetId="3" state="veryHidden" r:id="rId3"/>
  </sheets>
  <calcPr calcId="191029" fullCalcOnLoad="0"/>
</workbook>`;

const f14WorkbookRels = workbookRelsXml([
  ['rId1', `${REL_NS}/worksheet`, 'worksheets/sheet1.xml'],
  ['rId2', `${REL_NS}/worksheet`, 'worksheets/sheet2.xml'],
  ['rId3', `${REL_NS}/worksheet`, 'worksheets/sheet3.xml'],
  ['rId4', `${REL_NS}/styles`, 'styles.xml'],
]);

const f14Board = `${XML_HEADER}
<worksheet xmlns="${SPREADSHEET_NS}">
  <sheetPr><tabColor rgb="FF1F4E78"/></sheetPr>
  <dimension ref="A1:D5"/>
  <sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="15" defaultColWidth="12"/>
  <cols><col min="2" max="3" width="14" customWidth="1"/><col min="4" max="4" width="9" hidden="1" customWidth="1"/></cols>
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Seat</t></is></c><c r="B1" t="inlineStr"><is><t>Member</t></is></c><c r="C1" t="inlineStr"><is><t>Role</t></is></c><c r="D1" t="inlineStr"><is><t>Notes</t></is></c></row>
    <row r="2" hidden="1"><c r="A2"><v>2</v></c><c r="B2" t="inlineStr"><is><t>Hidden row</t></is></c><c r="C2" t="inlineStr"><is><t>observer</t></is></c><c r="D2" t="inlineStr"><is><t>private</t></is></c></row>
    <row r="3"><c r="A3"><v>3</v></c><c r="B3" t="inlineStr"><is><t>Ana</t></is></c><c r="C3" t="inlineStr"><is><t>chair</t></is></c><c r="D3" t="inlineStr"><is><t>secret</t></is></c></row>
    <row r="4"><c r="A4"><v>4</v></c><c r="B4" t="inlineStr"><is><t>Bo</t></is></c><c r="C4" t="inlineStr"><is><t>member</t></is></c><c r="D4" t="inlineStr"><is><t>hidden column</t></is></c></row>
    <row r="5"><c r="A5"><v>5</v></c><c r="B5" t="inlineStr"><is><t>Cy</t></is></c><c r="C5" t="inlineStr"><is><t>member</t></is></c><c r="D5" t="inlineStr"><is><t>done</t></is></c></row>
  </sheetData>
  <sheetProtection sheet="1" objects="1" scenarios="1" formatCells="0" formatColumns="1" password="CC1A"/>
</worksheet>`;

const f14Members = `${XML_HEADER}
<worksheet xmlns="${SPREADSHEET_NS}">
  <dimension ref="A1:B2"/>
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Joined</t></is></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>Ana</t></is></c><c r="B2"><v>${serial1900(2020, 1, 15)}</v></c></row>
  </sheetData>
</worksheet>`;

const f14Archive = `${XML_HEADER}
<worksheet xmlns="${SPREADSHEET_NS}">
  <dimension ref="A1:A1"/>
  <sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>archived</t></is></c></row></sheetData>
</worksheet>`;

function f14Package() {
  return buildPackage({
    workbook: f14Workbook,
    workbookRels: f14WorkbookRels,
    sheets: [
      ['xl/worksheets/sheet1.xml', f14Board],
      ['xl/worksheets/sheet2.xml', f14Members],
      ['xl/worksheets/sheet3.xml', f14Archive],
    ],
    extra: { 'xl/styles.xml': TWO_STYLES },
  });
}

// ---------------------------------------------------------------------------
// F16 print-setup.xlsx — print area, orientation, fit-to-width and margins
// (file-compat §3, §9 F16).
// ---------------------------------------------------------------------------

const f16Workbook = workbookXml({
  sheets: [['Report', 'visible']],
  definedNames: `<definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">'Report'!$A$1:$D$20</definedName></definedNames>`,
});

const f16Sheet = `${XML_HEADER}
<worksheet xmlns="${SPREADSHEET_NS}">
  <sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>
  <dimension ref="A1:D6"/>
  <sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="15" defaultColWidth="11"/>
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Region</t></is></c><c r="B1" t="inlineStr"><is><t>Q1</t></is></c><c r="C1" t="inlineStr"><is><t>Q2</t></is></c><c r="D1" t="inlineStr"><is><t>Total</t></is></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>North</t></is></c><c r="B2"><v>120</v></c><c r="C2"><v>140</v></c><c r="D2"><v>260</v></c></row>
    <row r="3"><c r="A3" t="inlineStr"><is><t>South</t></is></c><c r="B3"><v>90</v></c><c r="C3"><v>110</v></c><c r="D3"><v>200</v></c></row>
    <row r="4"><c r="A4" t="inlineStr"><is><t>East</t></is></c><c r="B4"><v>80</v></c><c r="C4"><v>95</v></c><c r="D4"><v>175</v></c></row>
  </sheetData>
  <printOptions horizontalCentered="1" verticalCentered="0"/>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
  <pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0" scale="85" firstPageNumber="1" useFirstPageNumber="1" cellComments="none"/>
</worksheet>`;

function f16Package() {
  return buildPackage({
    workbook: f16Workbook,
    workbookRels: workbookRelsXml([
      ['rId1', `${REL_NS}/worksheet`, 'worksheets/sheet1.xml'],
      ['rId2', `${REL_NS}/styles`, 'styles.xml'],
    ]),
    sheets: [['xl/worksheets/sheet1.xml', f16Sheet]],
    extra: { 'xl/styles.xml': TWO_STYLES },
  });
}

// ---------------------------------------------------------------------------
// F17 rich-text.xlsx — untouched runs survive; editing warns and flattens
// (file-compat §3, §9 F17; workbook-model §9).
// ---------------------------------------------------------------------------

const f17SharedStrings = `${XML_HEADER}
<sst xmlns="${SPREADSHEET_NS}" count="3" uniqueCount="2">
  <si><r><rPr><b/><sz val="11"/><color rgb="FF1F4E78"/><rFont val="Calibri"/><family val="2"/></rPr><t>Bold lead</t></r><r><rPr><sz val="11"/><rFont val="Calibri"/><family val="2"/></rPr><t xml:space="preserve"> plain tail</t></r></si>
  <si><t>plain cell</t></si>
</sst>`;

const f17Sheet = `${XML_HEADER}
<worksheet xmlns="${SPREADSHEET_NS}">
  <dimension ref="A1:C3"/>
  <sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetData>
    <row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1"><v>1</v></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><r><rPr><b/><sz val="11"/><rFont val="Calibri"/></rPr><t>Rich</t></r><r><t xml:space="preserve"> inline</t></r></is></c><c r="B2"><v>7</v></c><c r="C2" t="s" s="1"><v>0</v></c></row>
    <row r="3"><c r="A3" t="inlineStr"><is><t>plain inline</t></is></c><c r="B3"><v>9</v></c></row>
  </sheetData>
</worksheet>`;

function f17Package() {
  return buildPackage({
    workbook: workbookXml({ sheets: [['Notes', 'visible']] }),
    workbookRels: workbookRelsXml([
      ['rId1', `${REL_NS}/worksheet`, 'worksheets/sheet1.xml'],
      ['rId2', `${REL_NS}/styles`, 'styles.xml'],
      ['rId3', `${REL_NS}/sharedStrings`, 'sharedStrings.xml'],
    ]),
    sheets: [['xl/worksheets/sheet1.xml', f17Sheet]],
    sharedStrings: f17SharedStrings,
    extra: { 'xl/styles.xml': TWO_STYLES },
  });
}

// ---------------------------------------------------------------------------
// F23 names-defined.xlsx — workbook and sheet names, quoted sheet references
// and their range transforms (file-compat §3, §9 F23).
// ---------------------------------------------------------------------------

const f23Workbook = workbookXml({
  sheets: [
    ['Sales Data', 'visible'],
    ['Summary', 'visible'],
    ["O'Brien's Log", 'visible'],
  ],
  definedNames: `<definedNames>
    <definedName name="TotalSales">'Sales Data'!$B$2:$B$5</definedName>
    <definedName name="TaxRate" localSheetId="0" hidden="1">0.2</definedName>
    <definedName name="BoldTitle" localSheetId="1">'Summary'!$A$1</definedName>
    <definedName name="LocalFigure" localSheetId="0">$B$2</definedName>
    <definedName name="QuotedApostrophe">'O''Brien''s Log'!$A$1:$A$2</definedName>
  </definedNames>`,
});

const f23WorkbookRels = workbookRelsXml([
  ['rId1', `${REL_NS}/worksheet`, 'worksheets/sheet1.xml'],
  ['rId2', `${REL_NS}/worksheet`, 'worksheets/sheet2.xml'],
  ['rId3', `${REL_NS}/worksheet`, 'worksheets/sheet3.xml'],
  ['rId4', `${REL_NS}/styles`, 'styles.xml'],
]);

const f23Sales = `${XML_HEADER}
<worksheet xmlns="${SPREADSHEET_NS}">
  <dimension ref="A1:D5"/>
  <sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Month</t></is></c><c r="B1" t="inlineStr"><is><t>Amount</t></is></c><c r="C1" t="inlineStr"><is><t>Region</t></is></c><c r="D1" t="inlineStr"><is><t>Rep</t></is></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>Jan</t></is></c><c r="B2"><v>100</v></c><c r="C2" t="inlineStr"><is><t>North</t></is></c><c r="D2" t="inlineStr"><is><t>Ana</t></is></c></row>
    <row r="3"><c r="A3" t="inlineStr"><is><t>Feb</t></is></c><c r="B3"><v>200</v></c><c r="C3" t="inlineStr"><is><t>South</t></is></c><c r="D3" t="inlineStr"><is><t>Bo</t></is></c></row>
    <row r="4"><c r="A4" t="inlineStr"><is><t>Mar</t></is></c><c r="B4"><v>300</v></c><c r="C4" t="inlineStr"><is><t>East</t></is></c><c r="D4" t="inlineStr"><is><t>Cy</t></is></c></row>
    <row r="5"><c r="A5" t="inlineStr"><is><t>Apr</t></is></c><c r="B5"><v>400</v></c><c r="C5" t="inlineStr"><is><t>West</t></is></c><c r="D5" t="inlineStr"><is><t>Di</t></is></c></row>
  </sheetData>
</worksheet>`;

const f23Summary = `${XML_HEADER}
<worksheet xmlns="${SPREADSHEET_NS}">
  <dimension ref="A1:C4"/>
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Summary</t></is></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>Total</t></is></c><c r="B2"><f>SUM('Sales Data'!$B$2:$B$5)</f><v>1000</v></c></row>
    <row r="3"><c r="A3" t="inlineStr"><is><t>Check</t></is></c><c r="B3"><v>1</v></c></row>
  </sheetData>
</worksheet>`;

const f23Log = `${XML_HEADER}
<worksheet xmlns="${SPREADSHEET_NS}">
  <dimension ref="A1:B2"/>
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Entry</t></is></c><c r="B1" t="inlineStr"><is><t>Owner</t></is></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>Quoted</t></is></c><c r="B2" t="inlineStr"><is><t>O'Brien</t></is></c></row>
  </sheetData>
</worksheet>`;

function f23Package() {
  return buildPackage({
    workbook: f23Workbook,
    workbookRels: f23WorkbookRels,
    sheets: [
      ['xl/worksheets/sheet1.xml', f23Sales],
      ['xl/worksheets/sheet2.xml', f23Summary],
      ['xl/worksheets/sheet3.xml', f23Log],
    ],
    extra: { 'xl/styles.xml': TWO_STYLES },
  });
}

export function build({ compatDir }) {
  const fixtures = [
    ['threaded-comments.xlsx', f12Package()],
    ['validation-cf.xlsx', f13Package()],
    ['protection-hidden.xlsx', f14Package()],
    ['print-setup.xlsx', f16Package()],
    ['rich-text.xlsx', f17Package()],
    ['names-defined.xlsx', f23Package()],
  ];
  for (const [name, bytes] of fixtures) {
    const path = join(compatDir, name);
    writeFileSync(path, bytes);
    console.log(`Wrote ${path}`);
  }
}
