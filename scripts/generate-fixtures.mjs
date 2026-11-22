import { mkdirSync, readdirSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import {
  XML_HEADER,
  SPREADSHEET_NS,
  REL_NS,
  PACKAGE_REL_NS,
  themeXml,
  stylesXml,
  sharedStringsXml,
  workbookXml,
  workbookRels,
  sheetRels,
  packageEntries,
} from './fixtures/parts.mjs';

const root = dirname(dirname(fileURLToPath(import.meta.url)));
const compatDir = join(root, 'fixtures', 'compat');
const expectedDir = join(root, 'fixtures', 'expected');
const manualDir = join(root, 'fixtures', 'manual');
const oracleDir = join(root, 'fixtures', 'oracle');
mkdirSync(compatDir, { recursive: true });
mkdirSync(expectedDir, { recursive: true });
mkdirSync(manualDir, { recursive: true });
mkdirSync(oracleDir, { recursive: true });

const drawingRels = `${XML_HEADER}
<Relationships xmlns="${PACKAGE_REL_NS}">
  <Relationship Id="rId1" Type="${REL_NS}/chart" Target="../charts/chart1.xml"/>
</Relationships>`;

const chartXml = `${XML_HEADER}
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${REL_NS}">
  <c:chart>
    <c:plotArea>
      <c:layout/>
      <c:barChart>
        <c:barDir val="col"/>
        <c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numRef><c:f>Sheet1!$B$2:$B$3</c:f></c:numRef></c:val></c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:barChart>
      <c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/></c:catAx>
      <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/></c:valAx>
    </c:plotArea>
    <c:legend/>
    <c:plotVisOnly val="1"/>
  </c:chart>
</c:chartSpace>`;

const drawingXml = `${XML_HEADER}
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${REL_NS}">
  <xdr:oneCellAnchor>
    <xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:ext cx="2400000" cy="1600000"/>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
      <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="2400000" cy="1600000"/></xdr:xfrm>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId1"/></a:graphicData></a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:oneCellAnchor>
</xdr:wsDr>`;

const vmlXml = `${XML_HEADER}
<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
  <o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>
  <v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>
</xml>`;

const sheetOneXml = `${XML_HEADER}
<worksheet xmlns="${SPREADSHEET_NS}" xmlns:r="${REL_NS}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" mc:Ignorable="x14">
  <sheetPr><tabColor rgb="FF22AA55"/><pageSetUpPr fitToPage="1"/></sheetPr>
  <dimension ref="A1:C4"/>
  <sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <cols><col min="1" max="1" width="18.5" customWidth="1"/><col min="2" max="2" width="14" customWidth="1"/></cols>
  <sheetData>
    <row r="1" ht="22" customHeight="1" s="1" customFormat="1">
      <c r="A1" t="s" s="1"><v>0</v></c>
      <c r="B1" t="s"><v>1</v></c>
    </row>
    <row r="2">
      <c r="A2" t="s"><v>2</v></c>
      <c r="B2" x14:foo="bar"><v>42.5</v></c>
      <c r="C2" t="inlineStr"><is><t>inline text</t></is></c>
    </row>
    <row r="3">
      <c r="A3" t="s"><v>3</v></c>
      <c r="B3"><f>B2*2</f><v>85</v></c>
      <c r="C3"><v>3.141592653589793</v></c>
    </row>
    <row r="4">
      <c r="A4" t="b"><v>1</v></c>
      <c r="B4" t="e"><v>#DIV/0!</v></c>
    </row>
  </sheetData>
  <mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>
  <phoneticPr fontId="1" type="noConversion"/>
  <drawing r:id="rId1"/>
  <legacyDrawing r:id="rId2"/>
  <extLst>
    <ext uri="{5C3D1A0B-6E2F-4C7B-9A11-CALCPLUS0001}" xmlns:cp="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties">
      <cp:properties>
        <cp:property name="calcplus.unknown" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2">
          <vt:lpwstr xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">preserve-me</vt:lpwstr>
        </cp:property>
      </cp:properties>
    </ext>
  </extLst>
</worksheet>`;

// Excel-openable value/error fixture for the manual Excel capture. The chart
// fixture is a preservation fixture: its hand-authored chart and extLst are
// deliberately minimal and Excel 365 refuses to open it. This workbook keeps
// the same M1 cell shapes without the unmodelled parts so the COM harness can
// capture real Excel values, errors and formula status.
const oracleWorkbookRels = `${XML_HEADER}
<Relationships xmlns="${PACKAGE_REL_NS}">
  <Relationship Id="rId1" Type="${REL_NS}/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="${REL_NS}/styles" Target="styles.xml"/>
  <Relationship Id="rId3" Type="${REL_NS}/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`;

const oracleSharedStringsXml = `${XML_HEADER}
<sst xmlns="${SPREADSHEET_NS}" count="6" uniqueCount="6">
  <si><t>Name</t></si>
  <si><t>Value</t></si>
  <si><t>Alpha</t></si>
  <si><t>Beta</t></si>
  <si><t>=not a formula</t></si>
  <si><t>=SUM(A1:A2)</t></si>
</sst>`;

const oracleSheetXml = `${XML_HEADER}
<worksheet xmlns="${SPREADSHEET_NS}" xmlns:r="${REL_NS}">
  <dimension ref="A1:E5"/>
  <cols><col min="4" max="4" width="22" customWidth="1"/><col min="5" max="5" width="4" customWidth="1"/></cols>
  <sheetData>
    <row r="1">
      <c r="A1" t="s" s="1"><v>0</v></c>
      <c r="B1" t="s"><v>1</v></c>
    </row>
    <row r="2">
      <c r="A2" t="s"><v>2</v></c>
      <c r="B2"><v>42.5</v></c>
      <c r="C2" t="inlineStr"><is><t>inline text</t></is></c>
      <c r="D2" t="s"><v>4</v></c>
      <c r="E2" t="e"><v>#N/A</v></c>
    </row>
    <row r="3">
      <c r="A3" t="s"><v>3</v></c>
      <c r="B3"><f>B2*2</f><v>85</v></c>
      <c r="C3"><v>3.141592653589793</v></c>
      <c r="D3" t="s"><v>5</v></c>
      <c r="E3" t="e"><v>#VALUE!</v></c>
    </row>
    <row r="4">
      <c r="A4" t="b"><v>1</v></c>
      <c r="B4" t="e"><v>#DIV/0!</v></c>
      <c r="C4" t="str"><f>"=literal"</f><v>=literal</v></c>
      <c r="D4" t="e"><f>1/0</f><v>#DIV/0!</v></c>
      <c r="E4"><f>SUM(B2:B3)</f><v>127.5</v></c>
    </row>
    <row r="5">
      <c r="A5" t="e"><v>#REF!</v></c>
      <c r="B5" t="e"><v>#NAME?</v></c>
      <c r="C5" t="e"><v>#NUM!</v></c>
      <c r="D5" t="e"><v>#NULL!</v></c>
    </row>
  </sheetData>
</worksheet>`;

const identityWorkbookXml = `${XML_HEADER}
<workbook xmlns="${SPREADSHEET_NS}" xmlns:r="${REL_NS}">
  <workbookPr date1904="0"/>
  <bookViews><workbookView activeTab="0"/></bookViews>
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`;

const identityWorkbookRels = `${XML_HEADER}
<Relationships xmlns="${PACKAGE_REL_NS}">
  <Relationship Id="rId1" Type="${REL_NS}/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="${REL_NS}/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`;

const identitySharedStrings = `${XML_HEADER}
<sst xmlns="${SPREADSHEET_NS}" count="1" uniqueCount="1"><si><t>Hello</t></si></sst>`;

const identitySheetXml = `${XML_HEADER}
<worksheet xmlns="${SPREADSHEET_NS}" xmlns:r="${REL_NS}">
  <dimension ref="A1:B2"/>
  <sheetData>
    <row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>1</v></c></row>
    <row r="2"><c r="A2"><v>2.5</v></c></row>
  </sheetData>
</worksheet>`;

const m1Package = packageEntries({
  workbook: workbookXml,
  workbookRels,
  sheets: [['xl/worksheets/sheet1.xml', sheetOneXml]],
  sharedStrings: sharedStringsXml,
  extra: {
    'xl/styles.xml': stylesXml,
    'xl/theme/theme1.xml': themeXml,
    'xl/worksheets/_rels/sheet1.xml.rels': sheetRels,
    'xl/drawings/drawing1.xml': drawingXml,
    'xl/drawings/_rels/drawing1.xml.rels': drawingRels,
    'xl/charts/chart1.xml': chartXml,
    'xl/drawings/vmlDrawing1.vml': vmlXml,
  },
});

const identityPackage = packageEntries({
  workbook: identityWorkbookXml,
  workbookRels: identityWorkbookRels,
  sheets: [['xl/worksheets/sheet1.xml', identitySheetXml]],
  sharedStrings: identitySharedStrings,
});

// Ticket 32: a cross-sheet link (and formula, defined name and chart series)
// that must follow an insert/delete on another sheet. The `Links` sheet holds a
// hyperlink and no cell that a structural edit moves, so a save that forgets a
// sibling part leaves it pointing at the old row.
const crossSheetWorkbookXml = `${XML_HEADER}
<workbook xmlns="${SPREADSHEET_NS}" xmlns:r="${REL_NS}">
  <workbookPr date1904="0"/>
  <bookViews><workbookView activeTab="0"/></bookViews>
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    <sheet name="Sheet2" sheetId="2" r:id="rId2"/>
    <sheet name="Links" sheetId="3" r:id="rId3"/>
  </sheets>
  <definedNames><definedName name="TargetCell">Sheet1!$A$3</definedName></definedNames>
  <calcPr calcId="191029" fullCalcOnLoad="0"/>
</workbook>`;

const crossSheetWorkbookRels = `${XML_HEADER}
<Relationships xmlns="${PACKAGE_REL_NS}">
  <Relationship Id="rId1" Type="${REL_NS}/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="${REL_NS}/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="${REL_NS}/worksheet" Target="worksheets/sheet3.xml"/>
</Relationships>`;

const crossSheetSheet1Xml = `${XML_HEADER}
<worksheet xmlns="${SPREADSHEET_NS}">
  <dimension ref="A1:B4"/>
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Target</t></is></c><c r="B1" t="inlineStr"><is><t>Value</t></is></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>One</t></is></c><c r="B2"><v>10</v></c></row>
    <row r="3"><c r="A3" t="inlineStr"><is><t>Two</t></is></c><c r="B3"><v>20</v></c></row>
    <row r="4"><c r="A4" t="inlineStr"><is><t>Three</t></is></c><c r="B4"><v>30</v></c></row>
  </sheetData>
</worksheet>`;

const crossSheetSheet2Xml = `${XML_HEADER}
<worksheet xmlns="${SPREADSHEET_NS}" xmlns:r="${REL_NS}">
  <dimension ref="A1:C4"/>
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Mirror</t></is></c></row>
    <row r="2"><c r="A2"><f>Sheet1!B3</f><v>20</v></c></row>
    <row r="3"><c r="B3" t="inlineStr"><is><t>jump</t></is></c></row>
  </sheetData>
  <hyperlinks><hyperlink ref="B3" location="Sheet1!A3" display="jump"/></hyperlinks>
  <drawing r:id="rId1"/>
</worksheet>`;

const crossSheetLinksXml = `${XML_HEADER}
<worksheet xmlns="${SPREADSHEET_NS}">
  <dimension ref="A1:B2"/>
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Links</t></is></c></row>
    <row r="2"><c r="B2" t="inlineStr"><is><t>go</t></is></c></row>
  </sheetData>
  <hyperlinks><hyperlink ref="B2" location="Sheet1!A3" display="go"/></hyperlinks>
</worksheet>`;

const crossSheetSheet2Rels = `${XML_HEADER}
<Relationships xmlns="${PACKAGE_REL_NS}">
  <Relationship Id="rId1" Type="${REL_NS}/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`;

const crossSheetDrawingXml = `${XML_HEADER}
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="${REL_NS}">
  <xdr:twoCellAnchor editAs="twoCell">
    <xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>9</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
      <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>`;

const crossSheetDrawingRels = `${XML_HEADER}
<Relationships xmlns="${PACKAGE_REL_NS}">
  <Relationship Id="rId1" Type="${REL_NS}/chart" Target="../charts/chart1.xml"/>
</Relationships>`;

const crossSheetChartXml = `${XML_HEADER}
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${REL_NS}">
  <c:chart>
    <c:plotArea>
      <c:layout/>
      <c:barChart>
        <c:barDir val="col"/>
        <c:ser><c:idx val="0"/><c:order val="0"/>
          <c:cat><c:strRef><c:f>Sheet1!$A$3</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Two</c:v></c:pt></c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$B$3</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="1"/><c:pt idx="0"><c:v>20</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:barChart>
      <c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/></c:catAx>
      <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/></c:valAx>
    </c:plotArea>
    <c:legend/>
    <c:plotVisOnly val="1"/>
  </c:chart>
</c:chartSpace>`;

const crossSheetPackage = packageEntries({
  workbook: crossSheetWorkbookXml,
  workbookRels: crossSheetWorkbookRels,
  sheets: [
    ['xl/worksheets/sheet1.xml', crossSheetSheet1Xml],
    ['xl/worksheets/sheet2.xml', crossSheetSheet2Xml],
    ['xl/worksheets/sheet3.xml', crossSheetLinksXml],
  ],
  extra: {
    'xl/worksheets/_rels/sheet2.xml.rels': crossSheetSheet2Rels,
    'xl/drawings/drawing1.xml': crossSheetDrawingXml,
    'xl/drawings/_rels/drawing1.xml.rels': crossSheetDrawingRels,
    'xl/charts/chart1.xml': crossSheetChartXml,
  },
});

// Ticket 32: the broken-internal-link convention (known deviation D-9, settled
// 2026-11-22). The target is gone and the author tool wrote `#REF!` in
// `location`; opening, saving and a later structural edit must keep the marker
// rather than block the sheet. The Excel-oracle capture in
// fixtures/oracle/d9-broken-hyperlink/ records that Excel freezes the authored
// location instead; calcplus keeps the marker as the deliberate difference.
const brokenLinkWorkbookXml = `${XML_HEADER}
<workbook xmlns="${SPREADSHEET_NS}" xmlns:r="${REL_NS}">
  <workbookPr date1904="0"/>
  <bookViews><workbookView activeTab="0"/></bookViews>
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
  <calcPr calcId="191029" fullCalcOnLoad="0"/>
</workbook>`;

const brokenLinkWorkbookRels = `${XML_HEADER}
<Relationships xmlns="${PACKAGE_REL_NS}">
  <Relationship Id="rId1" Type="${REL_NS}/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`;

const brokenLinkSheetXml = `${XML_HEADER}
<worksheet xmlns="${SPREADSHEET_NS}">
  <dimension ref="A1:B3"/>
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Broken link</t></is></c></row>
    <row r="3"><c r="B3" t="inlineStr"><is><t>gone</t></is></c></row>
  </sheetData>
  <hyperlinks><hyperlink ref="B3" location="#REF!" display="gone"/></hyperlinks>
</worksheet>`;

const brokenLinkPackage = packageEntries({
  workbook: brokenLinkWorkbookXml,
  workbookRels: brokenLinkWorkbookRels,
  sheets: [['xl/worksheets/sheet1.xml', brokenLinkSheetXml]],
});

// M2 grid fixture: frozen row and column, a merged range, imported row and
// column sizes, and the font, fill, alignment and number-format styles the
// grid renders in the grid-foundation milestone.
const m2StylesXml = `${XML_HEADER}
<styleSheet xmlns="${SPREADSHEET_NS}">
  <numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0.00"/></numFmts>
  <fonts count="4">
    <font><sz val="11"/><name val="Calibri"/><family val="2"/></font>
    <font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>
    <font><i/><sz val="14"/><color rgb="FFB00020"/><name val="Calibri"/><family val="2"/></font>
    <font><b/><sz val="14"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="6">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="right"/></xf>
    <xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="0" fontId="3" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
    <xf numFmtId="10" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;

const m2SharedStringsXml = `${XML_HEADER}
<sst xmlns="${SPREADSHEET_NS}" count="0" uniqueCount="0"/>`;

const m2Rows = [];
m2Rows.push(
  '<row r="1" ht="30" customHeight="1"><c r="A1" t="inlineStr" s="4"><is><t>Quarterly report</t></is></c><c r="B1" s="4"/></row>',
);
m2Rows.push(
  '<row r="2">' +
    '<c r="A2" t="inlineStr" s="1"><is><t>Item</t></is></c>' +
    '<c r="B2" t="inlineStr" s="1"><is><t>Amount</t></is></c>' +
    '<c r="C2" t="inlineStr" s="1"><is><t>Share</t></is></c>' +
    '<c r="D2" t="inlineStr" s="1"><is><t>Note</t></is></c>' +
    '</row>',
);
for (let index = 0; index < 300; index += 1) {
  const row = index + 3;
  const amount = ((index * 37) % 1000) + 0.25;
  const share = ((index * 13) % 100) / 100;
  const note = index % 25 === 0 ? 'checked' : '';
  m2Rows.push(
    `<row r="${row}">` +
      `<c r="A${row}" t="inlineStr"><is><t>Item ${index + 1}</t></is></c>` +
      `<c r="B${row}" s="2"><v>${amount}</v></c>` +
      `<c r="C${row}" s="5"><v>${share}</v></c>` +
      (note === '' ? '' : `<c r="D${row}" t="inlineStr" s="3"><is><t>${note}</t></is></c>`) +
      '</row>',
  );
}

const m2SheetXml = `${XML_HEADER}
<worksheet xmlns="${SPREADSHEET_NS}" xmlns:r="${REL_NS}">
  <dimension ref="A1:D302"/>
  <sheetViews><sheetView tabSelected="1" workbookViewId="0">\${m2Pane}</sheetView></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <cols><col min="1" max="1" width="24" customWidth="1"/><col min="2" max="2" width="14" customWidth="1"/></cols>
  <sheetData>
    ${m2Rows.join('\n    ')}
  </sheetData>
  <mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>
</worksheet>`;

const m2WorkbookRels = `${XML_HEADER}
<Relationships xmlns="${PACKAGE_REL_NS}">
  <Relationship Id="rId1" Type="${REL_NS}/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="${REL_NS}/styles" Target="styles.xml"/>
  <Relationship Id="rId3" Type="${REL_NS}/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`;

// Three freeze shapes over the same sheet: both axes, rows only, columns only.
const m2Package = packageEntries({
  workbook: workbookXml,
  workbookRels: m2WorkbookRels,
  sheets: [
    [
      'xl/worksheets/sheet1.xml',
      m2SheetXml.replace(
        '${m2Pane}',
        '<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/>',
      ),
    ],
  ],
  sharedStrings: m2SharedStringsXml,
  extra: { 'xl/styles.xml': m2StylesXml },
});

const m2FreezeRowPackage = packageEntries({
  workbook: workbookXml,
  workbookRels: m2WorkbookRels,
  sheets: [
    [
      'xl/worksheets/sheet1.xml',
      m2SheetXml.replace(
        '${m2Pane}',
        '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>',
      ),
    ],
  ],
  sharedStrings: m2SharedStringsXml,
  extra: { 'xl/styles.xml': m2StylesXml },
});

const m2FreezeColumnPackage = packageEntries({
  workbook: workbookXml,
  workbookRels: m2WorkbookRels,
  sheets: [
    [
      'xl/worksheets/sheet1.xml',
      m2SheetXml.replace(
        '${m2Pane}',
        '<pane xSplit="1" topLeftCell="B1" activePane="topRight" state="frozen"/>',
      ),
    ],
  ],
  sharedStrings: m2SharedStringsXml,
  extra: { 'xl/styles.xml': m2StylesXml },
});

const oraclePackage = packageEntries({
  workbook: workbookXml,
  workbookRels: oracleWorkbookRels,
  sheets: [['xl/worksheets/sheet1.xml', oracleSheetXml]],
  sharedStrings: oracleSharedStringsXml,
  extra: {
    'xl/styles.xml': stylesXml,
  },
});

// Manual smoke workbook. It is not a conformance fixture: a person opens it,
// scrolls, selects, finds, and checks the frozen panes, fills, formats and
// cached formula results by eye. Keep every value plain and synthetic.
const manualStylesXml = `${XML_HEADER}
<styleSheet xmlns="${SPREADSHEET_NS}">
  <numFmts count="5">
    <numFmt numFmtId="164" formatCode="yyyy-mm-dd"/>
    <numFmt numFmtId="165" formatCode="yyyy-mm-dd hh:mm"/>
    <numFmt numFmtId="166" formatCode="#,##0.00"/>
    <numFmt numFmtId="167" formatCode="0.0%"/>
    <numFmt numFmtId="168" formatCode="$#,##0.00"/>
  </numFmts>
  <fonts count="5">
    <font><sz val="11"/><name val="Calibri"/><family val="2"/></font>
    <font><b/><sz val="16"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>
    <font><b/><sz val="11"/><color rgb="FF1F1F1F"/><name val="Calibri"/><family val="2"/></font>
    <font><b/><sz val="11"/><color rgb="FFC00000"/><name val="Calibri"/><family val="2"/></font>
    <font><i/><sz val="10"/><color rgb="FF595959"/><name val="Calibri"/><family val="2"/></font>
  </fonts>
  <fills count="6">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFD9E2F3"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFE2EFDA"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border><left/><right/><top/><bottom style="thin"><color rgb="FFBFBFBF"/></bottom><diagonal/></border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="15">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="2" fillId="4" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="166" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="1" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center"/></xf>
    <xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="167" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="168" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="166" fontId="0" fillId="5" borderId="0" xfId="0" applyNumberFormat="1" applyFill="1"/>
    <xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyFont="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;

const manualDateSerial = Math.floor(Date.UTC(2026, 10, 15) / 86_400_000) + 25_569;

const manualCoverRows = [
  '<row r="1" ht="28" customHeight="1">' +
    '<c r="A1" t="inlineStr" s="1"><is><t>General manual test workbook</t></is></c>' +
    '<c r="B1" s="1"/><c r="C1" s="1"/><c r="D1" s="1"/>' +
    '</row>',
  '<row r="2" ht="20" customHeight="1">' +
    '<c r="A2" t="inlineStr" s="2"><is><t>Case</t></is></c>' +
    '<c r="B2" t="inlineStr" s="2"><is><t>Value</t></is></c>' +
    '<c r="C2" t="inlineStr" s="2"><is><t>Format</t></is></c>' +
    '<c r="D2" t="inlineStr" s="2"><is><t>Notes</t></is></c>' +
    '</row>',
  '<row r="3">' +
    '<c r="A3" t="inlineStr" s="3"><is><t>Text</t></is></c>' +
    '<c r="B3" t="inlineStr" s="4"><is><t>Plain synthetic text</t></is></c>' +
    '<c r="C3" t="inlineStr" s="4"><is><t>General</t></is></c>' +
    '<c r="D3" t="inlineStr" s="14"><is><t>Left-aligned text</t></is></c>' +
    '</row>',
  '<row r="4">' +
    '<c r="A4" t="inlineStr" s="3"><is><t>Number</t></is></c>' +
    '<c r="B4" s="5"><v>1234.5678</v></c>' +
    '<c r="C4" t="inlineStr" s="4"><is><t>#,##0.00</t></is></c>' +
    '<c r="D4" t="inlineStr" s="14"><is><t>Two decimal places</t></is></c>' +
    '</row>',
  '<row r="5">' +
    '<c r="A5" t="inlineStr" s="3"><is><t>Integer</t></is></c>' +
    '<c r="B5" s="6"><v>42</v></c>' +
    '<c r="C5" t="inlineStr" s="4"><is><t>0</t></is></c>' +
    '<c r="D5" t="inlineStr" s="14"><is><t>Whole number</t></is></c>' +
    '</row>',
  '<row r="6">' +
    '<c r="A6" t="inlineStr" s="3"><is><t>Boolean true</t></is></c>' +
    '<c r="B6" t="b" s="7"><v>1</v></c>' +
    '<c r="C6" t="inlineStr" s="4"><is><t>General</t></is></c>' +
    '<c r="D6" t="inlineStr" s="14"><is><t>Boolean cell</t></is></c>' +
    '</row>',
  '<row r="7">' +
    '<c r="A7" t="inlineStr" s="3"><is><t>Boolean false</t></is></c>' +
    '<c r="B7" t="b" s="7"><v>0</v></c>' +
    '<c r="C7" t="inlineStr" s="4"><is><t>General</t></is></c>' +
    '<c r="D7" t="inlineStr" s="14"><is><t>Boolean cell</t></is></c>' +
    '</row>',
  '<row r="8">' +
    '<c r="A8" t="inlineStr" s="3"><is><t>Error #DIV/0!</t></is></c>' +
    '<c r="B8" t="e" s="8"><v>#DIV/0!</v></c>' +
    '<c r="C8" t="inlineStr" s="4"><is><t>Error</t></is></c>' +
    '<c r="D8" t="inlineStr" s="14"><is><t>Error cell</t></is></c>' +
    '</row>',
  '<row r="9">' +
    '<c r="A9" t="inlineStr" s="3"><is><t>Error #N/A</t></is></c>' +
    '<c r="B9" t="e" s="8"><v>#N/A</v></c>' +
    '<c r="C9" t="inlineStr" s="4"><is><t>Error</t></is></c>' +
    '<c r="D9" t="inlineStr" s="14"><is><t>Error cell</t></is></c>' +
    '</row>',
  '<row r="10">' +
    '<c r="A10" t="inlineStr" s="3"><is><t>Date</t></is></c>' +
    `<c r="B10" s="9"><v>${manualDateSerial}</v></c>` +
    '<c r="C10" t="inlineStr" s="4"><is><t>yyyy-mm-dd</t></is></c>' +
    '<c r="D10" t="inlineStr" s="14"><is><t>Date number format</t></is></c>' +
    '</row>',
  '<row r="11">' +
    '<c r="A11" t="inlineStr" s="3"><is><t>Date and time</t></is></c>' +
    `<c r="B11" s="10"><v>${manualDateSerial + 0.4895833333333333}</v></c>` +
    '<c r="C11" t="inlineStr" s="4"><is><t>yyyy-mm-dd hh:mm</t></is></c>' +
    '<c r="D11" t="inlineStr" s="14"><is><t>Date and time</t></is></c>' +
    '</row>',
  '<row r="12">' +
    '<c r="A12" t="inlineStr" s="3"><is><t>Percent</t></is></c>' +
    '<c r="B12" s="11"><v>0.256</v></c>' +
    '<c r="C12" t="inlineStr" s="4"><is><t>0.0%</t></is></c>' +
    '<c r="D12" t="inlineStr" s="14"><is><t>25.6%</t></is></c>' +
    '</row>',
  '<row r="13">' +
    '<c r="A13" t="inlineStr" s="3"><is><t>Currency</t></is></c>' +
    '<c r="B13" s="12"><v>1234.5</v></c>' +
    '<c r="C13" t="inlineStr" s="4"><is><t>$#,##0.00</t></is></c>' +
    '<c r="D13" t="inlineStr" s="14"><is><t>Currency display</t></is></c>' +
    '</row>',
  '<row r="14">' +
    '<c r="A14" t="inlineStr" s="3"><is><t>Formula</t></is></c>' +
    '<c r="B14" s="13"><f>B4*2</f><v>2469.1356</v></c>' +
    '<c r="C14" t="inlineStr" s="4"><is><t>#,##0.00</t></is></c>' +
    '<c r="D14" t="inlineStr" s="14"><is><t>Cached result 2469.14</t></is></c>' +
    '</row>',
  '<row r="15">' +
    '<c r="A15" t="inlineStr" s="3"><is><t>Formula sum</t></is></c>' +
    '<c r="B15" s="13"><f>SUM(B4:B5)</f><v>1276.5678</v></c>' +
    '<c r="C15" t="inlineStr" s="4"><is><t>#,##0.00</t></is></c>' +
    '<c r="D15" t="inlineStr" s="14"><is><t>Cached result 1276.57</t></is></c>' +
    '</row>',
];

const manualCoverSheetXml = `${XML_HEADER}
<worksheet xmlns="${SPREADSHEET_NS}">
  <dimension ref="A1:D15"/>
  <sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/><selection pane="bottomRight" activeCell="B2" sqref="B2"/></sheetView></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <cols>
    <col min="1" max="1" width="20" customWidth="1"/>
    <col min="2" max="2" width="22" customWidth="1"/>
    <col min="3" max="3" width="22" customWidth="1"/>
    <col min="4" max="4" width="30" customWidth="1"/>
  </cols>
  <sheetData>
    ${manualCoverRows.join('\n    ')}
  </sheetData>
  <mergeCells count="1"><mergeCell ref="A1:D1"/></mergeCells>
</worksheet>`;

const manualDataRows = [];
manualDataRows.push(
  '<row r="1" ht="20" customHeight="1">' +
    '<c r="A1" t="inlineStr" s="2"><is><t>ID</t></is></c>' +
    '<c r="B1" t="inlineStr" s="2"><is><t>Name</t></is></c>' +
    '<c r="C1" t="inlineStr" s="2"><is><t>Amount</t></is></c>' +
    '<c r="D1" t="inlineStr" s="2"><is><t>Ratio</t></is></c>' +
    '<c r="E1" t="inlineStr" s="2"><is><t>Date</t></is></c>' +
    '<c r="F1" t="inlineStr" s="2"><is><t>Flag</t></is></c>' +
    '<c r="G1" t="inlineStr" s="2"><is><t>Status</t></is></c>' +
    '<c r="H1" t="inlineStr" s="2"><is><t>Error</t></is></c>' +
    '<c r="I1" t="inlineStr" s="2"><is><t>Double</t></is></c>' +
    '</row>',
);
const manualStatuses = ['Open', 'Closed', 'Pending', 'Review'];
const manualErrors = ['#N/A', '#VALUE!', '#DIV/0!'];
for (let index = 0; index < 400; index += 1) {
  const row = index + 2;
  const amount = ((((index + 1) * 37) % 10_000) / 100 + 0.25).toFixed(2);
  const ratio = ((((index + 1) * 13) % 100) / 100).toFixed(2);
  const date = manualDateSerial + (index % 365);
  const flag = index % 2 === 0 ? 1 : 0;
  const status = manualStatuses[index % manualStatuses.length];
  const error =
    index > 0 && index % 31 === 0
      ? manualErrors[Math.floor(index / 31) % manualErrors.length]
      : null;
  manualDataRows.push(
    `<row r="${row}">` +
      `<c r="A${row}" s="6"><v>${index + 1}</v></c>` +
      `<c r="B${row}" t="inlineStr" s="4"><is><t>Item ${String(index + 1).padStart(3, '0')}</t></is></c>` +
      `<c r="C${row}" s="5"><v>${amount}</v></c>` +
      `<c r="D${row}" s="11"><v>${ratio}</v></c>` +
      `<c r="E${row}" s="9"><v>${date}</v></c>` +
      `<c r="F${row}" t="b" s="7"><v>${flag}</v></c>` +
      (index > 0 && index % 17 === 0
        ? ''
        : `<c r="G${row}" t="inlineStr" s="4"><is><t>${status}</t></is></c>`) +
      (error ? `<c r="H${row}" t="e" s="8"><v>${error}</v></c>` : '') +
      `<c r="I${row}" s="13"><f>C${row}*2</f><v>${(Number(amount) * 2).toFixed(2)}</v></c>` +
      '</row>',
  );
}

const manualDataSheetXml = `${XML_HEADER}
<worksheet xmlns="${SPREADSHEET_NS}">
  <dimension ref="A1:I401"/>
  <sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <cols>
    <col min="1" max="1" width="8" customWidth="1"/>
    <col min="2" max="2" width="18" customWidth="1"/>
    <col min="3" max="3" width="14" customWidth="1"/>
    <col min="4" max="4" width="10" customWidth="1"/>
    <col min="5" max="5" width="12" customWidth="1"/>
    <col min="6" max="6" width="10" customWidth="1"/>
    <col min="7" max="7" width="12" customWidth="1"/>
    <col min="8" max="8" width="10" customWidth="1"/>
    <col min="9" max="9" width="14" customWidth="1"/>
  </cols>
  <sheetData>
    ${manualDataRows.join('\n    ')}
  </sheetData>
</worksheet>`;

const manualWorkbookXml = `${XML_HEADER}
<workbook xmlns="${SPREADSHEET_NS}" xmlns:r="${REL_NS}">
  <workbookPr date1904="0"/>
  <bookViews><workbookView activeTab="0"/></bookViews>
  <sheets>
    <sheet name="Cover" sheetId="1" r:id="rId1"/>
    <sheet name="Data" sheetId="2" r:id="rId2"/>
  </sheets>
  <calcPr calcId="191029" fullCalcOnLoad="0"/>
</workbook>`;

const manualWorkbookRels = `${XML_HEADER}
<Relationships xmlns="${PACKAGE_REL_NS}">
  <Relationship Id="rId1" Type="${REL_NS}/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="${REL_NS}/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="${REL_NS}/styles" Target="styles.xml"/>
</Relationships>`;

const manualPackage = packageEntries({
  workbook: manualWorkbookXml,
  workbookRels: manualWorkbookRels,
  sheets: [
    ['xl/worksheets/sheet1.xml', manualCoverSheetXml],
    ['xl/worksheets/sheet2.xml', manualDataSheetXml],
  ],
  extra: {
    'xl/styles.xml': manualStylesXml,
  },
});

writeFileSync(join(compatDir, 'm1-chart-unknown.xlsx'), m1Package);
writeFileSync(join(compatDir, 'identity-minimal.xlsx'), identityPackage);
writeFileSync(join(compatDir, 'cross-sheet-link.xlsx'), crossSheetPackage);
writeFileSync(join(compatDir, 'broken-internal-link.xlsx'), brokenLinkPackage);
writeFileSync(join(compatDir, 'm2-grid.xlsx'), m2Package);
writeFileSync(join(compatDir, 'm2-freeze-row.xlsx'), m2FreezeRowPackage);
writeFileSync(join(compatDir, 'm2-freeze-column.xlsx'), m2FreezeColumnPackage);
writeFileSync(join(manualDir, 'general-test.xlsx'), manualPackage);
writeFileSync(join(oracleDir, 'm1-values-errors.xlsx'), oraclePackage);

const expectedValues = {
  fixture: 'm1-chart-unknown.xlsx',
  provenance: {
    kind: 'hand-authored',
    generator: 'scripts/generate-fixtures.mjs',
    oracle: null,
    manifest: 'fixtures/oracle-manifest.json',
    generatedAt: '2026-11-14',
    note: 'Hand-authored M1 expectations for the fixture generator output. These are not Excel 365 oracle goldens.',
  },
  tolerance: { relative: 1e-9, absolute: 1e-12 },
  cells: [
    { sheet: 'Sheet1', ref: 'A1', kind: 'text', value: 'Name' },
    { sheet: 'Sheet1', ref: 'B1', kind: 'text', value: 'Value' },
    { sheet: 'Sheet1', ref: 'A2', kind: 'text', value: 'Alpha' },
    { sheet: 'Sheet1', ref: 'B2', kind: 'number', value: 42.5 },
    { sheet: 'Sheet1', ref: 'C2', kind: 'text', value: 'inline text' },
    { sheet: 'Sheet1', ref: 'A3', kind: 'text', value: 'Beta' },
    { sheet: 'Sheet1', ref: 'B3', kind: 'formula', value: 'B2*2', cached: 85 },
    { sheet: 'Sheet1', ref: 'C3', kind: 'number', value: 3.141592653589793 },
    { sheet: 'Sheet1', ref: 'A4', kind: 'boolean', value: true },
    { sheet: 'Sheet1', ref: 'B4', kind: 'error', value: '#DIV/0!' },
  ],
};
writeFileSync(join(expectedDir, 'm1-values.json'), `${JSON.stringify(expectedValues, null, 2)}\n`);

const oracleManifest = {
  schemaVersion: 1,
  status: 'not-configured',
  oracle: {
    product: 'Microsoft Excel 365',
    version: null,
    build: null,
    os: null,
    locale: null,
    dateSystem: null,
  },
  generatedAt: null,
  goldens: [],
  note: 'No Excel 365 oracle goldens are committed or configured yet. See scripts/oracle/README.md. Hand-authored expectations live in fixtures/expected/ and are not oracle goldens; M3B uses fixtures/expected/m3b-formulas.json.',
};
writeFileSync(
  join(root, 'fixtures', 'oracle-manifest.json'),
  `${JSON.stringify(oracleManifest, null, 2)}\n`,
);

console.log(`Wrote ${join(compatDir, 'm1-chart-unknown.xlsx')}`);
console.log(`Wrote ${join(compatDir, 'identity-minimal.xlsx')}`);
console.log(`Wrote ${join(compatDir, 'cross-sheet-link.xlsx')}`);
console.log(`Wrote ${join(compatDir, 'broken-internal-link.xlsx')}`);
console.log(`Wrote ${join(compatDir, 'm2-grid.xlsx')}`);
console.log(`Wrote ${join(compatDir, 'm2-freeze-row.xlsx')}`);
console.log(`Wrote ${join(compatDir, 'm2-freeze-column.xlsx')}`);
console.log(`Wrote ${join(manualDir, 'general-test.xlsx')}`);
console.log(`Wrote ${join(oracleDir, 'm1-values-errors.xlsx')}`);
console.log(`Wrote ${join(expectedDir, 'm1-values.json')}`);
console.log(`Wrote ${join(root, 'fixtures', 'oracle-manifest.json')}`);

const groupsDir = join(root, 'scripts', 'fixtures', 'groups');
for (const name of readdirSync(groupsDir)
  .filter((file) => file.endsWith('.mjs'))
  .sort()) {
  const group = await import(pathToFileURL(join(groupsDir, name)).href);
  await group.build({ root, compatDir, expectedDir, manualDir, oracleDir });
}
