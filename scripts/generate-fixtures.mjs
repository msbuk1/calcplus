import { mkdirSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { zipSync } from 'fflate';

const root = dirname(dirname(fileURLToPath(import.meta.url)));
const compatDir = join(root, 'fixtures', 'compat');
const expectedDir = join(root, 'fixtures', 'expected');
mkdirSync(compatDir, { recursive: true });
mkdirSync(expectedDir, { recursive: true });

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const SPREADSHEET_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PACKAGE_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

const contentTypes = (overrides) => `${XML_HEADER}
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>
  ${overrides.map(([part, type]) => `<Override PartName="/${part}" ContentType="${type}"/>`).join('\n  ')}
</Types>`;

const rootRels = `${XML_HEADER}
<Relationships xmlns="${PACKAGE_REL_NS}">
  <Relationship Id="rId1" Type="${REL_NS}/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="${REL_NS}/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;

const docPropsCore = `${XML_HEADER}
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>calcplus fixtures</dc:creator>
  <cp:lastModifiedBy>calcplus fixtures</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">2026-11-14T00:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2026-11-14T00:00:00Z</dcterms:modified>
</cp:coreProperties>`;

const docPropsApp = `${XML_HEADER}
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>calcplus fixture generator</Application>
</Properties>`;

const themeXml = `${XML_HEADER}
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="calcplus">
  <a:themeElements>
    <a:clrScheme name="calcplus">
      <a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="calcplus">
      <a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="calcplus"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>
  </a:themeElements>
</a:theme>`;

const stylesXml = `${XML_HEADER}
<styleSheet xmlns="${SPREADSHEET_NS}">
  <numFmts count="1"><numFmt numFmtId="164" formatCode="0.00"/></numFmts>
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;

const sharedStringsXml = `${XML_HEADER}
<sst xmlns="${SPREADSHEET_NS}" count="4" uniqueCount="4">
  <si><t>Name</t></si>
  <si><t>Value</t></si>
  <si><t>Alpha</t></si>
  <si><t>Beta</t></si>
</sst>`;

const workbookXml = `${XML_HEADER}
<workbook xmlns="${SPREADSHEET_NS}" xmlns:r="${REL_NS}">
  <workbookPr date1904="0"/>
  <bookViews><workbookView activeTab="0"/></bookViews>
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
  <definedNames><definedName name="FixtureRange">Sheet1!$A$1:$B$2</definedName></definedNames>
  <calcPr calcId="191029" fullCalcOnLoad="0"/>
</workbook>`;

const workbookRels = `${XML_HEADER}
<Relationships xmlns="${PACKAGE_REL_NS}">
  <Relationship Id="rId1" Type="${REL_NS}/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="${REL_NS}/styles" Target="styles.xml"/>
  <Relationship Id="rId3" Type="${REL_NS}/sharedStrings" Target="sharedStrings.xml"/>
  <Relationship Id="rId4" Type="${REL_NS}/theme" Target="theme/theme1.xml"/>
</Relationships>`;

const sheetRels = `${XML_HEADER}
<Relationships xmlns="${PACKAGE_REL_NS}">
  <Relationship Id="rId1" Type="${REL_NS}/drawing" Target="../drawings/drawing1.xml"/>
  <Relationship Id="rId2" Type="${REL_NS}/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>
</Relationships>`;

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

function packageEntries({ workbook, workbookRels, sheets, sharedStrings, extra = {} }) {
  const overrides = [
    [
      'xl/workbook.xml',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
    ],
    ...sheets.map(([path]) => [
      path,
      'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
    ]),
    ...(sharedStrings
      ? [
          [
            'xl/sharedStrings.xml',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
          ],
        ]
      : []),
    ...Object.keys(extra)
      .filter((path) => path.endsWith('.xml') || path.endsWith('.vml'))
      .filter((path) => path !== 'docProps/core.xml' && path !== 'docProps/app.xml')
      .map((path) => [path, contentTypeFor(path)]),
    ['docProps/core.xml', 'application/vnd.openxmlformats-package.core-properties+xml'],
    ['docProps/app.xml', 'application/vnd.openxmlformats-officedocument.extended-properties+xml'],
  ];
  const files = {
    '[Content_Types].xml': contentTypes(overrides),
    '_rels/.rels': rootRels,
    'docProps/core.xml': docPropsCore,
    'docProps/app.xml': docPropsApp,
    'xl/workbook.xml': workbook,
    'xl/_rels/workbook.xml.rels': workbookRels,
    ...Object.fromEntries(sheets),
    ...(sharedStrings ? { 'xl/sharedStrings.xml': sharedStrings } : {}),
    ...extra,
  };
  return zipSync(
    Object.fromEntries(
      Object.entries(files).map(([name, content]) => [name, new TextEncoder().encode(content)]),
    ),
    { level: 6, mtime: new Date('1980-01-01T00:00:00Z') },
  );
}

function contentTypeFor(path) {
  if (path.startsWith('xl/styles'))
    return 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml';
  if (path.startsWith('xl/theme/'))
    return 'application/vnd.openxmlformats-officedocument.theme+xml';
  if (path.startsWith('xl/drawings/drawing'))
    return 'application/vnd.openxmlformats-officedocument.drawing+xml';
  if (path.startsWith('xl/charts/'))
    return 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
  if (path.endsWith('.vml')) return 'application/vnd.openxmlformats-officedocument.vmlDrawing';
  return 'application/xml';
}

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

writeFileSync(join(compatDir, 'm1-chart-unknown.xlsx'), m1Package);
writeFileSync(join(compatDir, 'identity-minimal.xlsx'), identityPackage);

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
  note: 'No Excel 365 oracle goldens are committed or configured yet. See scripts/oracle/README.md. Hand-authored M1 fixture expectations live in fixtures/expected/ and are not oracle goldens.',
};
writeFileSync(
  join(root, 'fixtures', 'oracle-manifest.json'),
  `${JSON.stringify(oracleManifest, null, 2)}\n`,
);

console.log(`Wrote ${join(compatDir, 'm1-chart-unknown.xlsx')}`);
console.log(`Wrote ${join(compatDir, 'identity-minimal.xlsx')}`);
console.log(`Wrote ${join(expectedDir, 'm1-values.json')}`);
console.log(`Wrote ${join(root, 'fixtures', 'oracle-manifest.json')}`);
