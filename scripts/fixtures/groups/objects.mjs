import { writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { unzipSync, zipSync } from 'fflate';
import {
  PACKAGE_REL_NS,
  REL_NS,
  SPREADSHEET_NS,
  XML_HEADER,
  packageEntries,
} from '../parts.mjs';

/**
 * M6 lane 4 fixtures: preserved parts and macros (file-compat §9 F9-F11).
 *
 *   F9  charts-images.xlsx  chart, image, shape, VML, OLE parts and the dirty
 *                           worksheet's unmodelled children
 *   F10 pivot-external.xlsx pivot table, pivot cache and external link
 *   F11 macros.xlsm         vbaProject.bin and the `.xlsx` Save As path
 *
 * Every fixture is hand-authored OOXML with fixed zip timestamps and stable
 * ordering, so `npm run fixtures` is byte-identical on every run
 * (scripts/fixtures/groups/README.md). The assertions live in
 * tests/m6-preserved-parts.test.ts.
 */

const FIXED_MTIME = new Date('1980-01-01T00:00:00Z');
const DRAWING_NS = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
const DRAWING_MAIN_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const CHART_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const CUSTOM_PROP_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties';
const VT_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes';
const MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const X14_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';

const CONTENT_TYPE = {
  chart: 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
  pivotTable: 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml',
  pivotCacheDefinition:
    'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml',
  pivotCacheRecords:
    'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml',
  externalLink: 'application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml',
  oleObject: 'application/vnd.openxmlformats-officedocument.oleObject',
  vbaProject: 'application/vnd.ms-office.vbaProject',
  macroWorkbook: 'application/vnd.ms-excel.sheet.macroEnabled.main+xml',
};

function escapeRegExp(text) {
  return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * `packageEntries` encodes every extra part as UTF-8 text, so binary parts
 * (images, OLE embeddings, the VBA project) are spliced in afterwards. The
 * rezip keeps the generator's fixed timestamp and level, so output stays
 * deterministic.
 */
function addBinaryParts(bytes, parts) {
  if (Object.keys(parts).length === 0) return bytes;
  const entries = unzipSync(bytes);
  for (const [path, content] of Object.entries(parts)) entries[path] = content;
  return zipSync(entries, { level: 6, mtime: FIXED_MTIME });
}

/**
 * Replace or add `[Content_Types].xml` entries. OPC allows one override per
 * part, so an existing override for the same part is replaced, never
 * duplicated. Unknown extensions (a PNG, a VBA project) get a Default or an
 * Override here because `parts.mjs` only knows the XML part types.
 */
function patchContentTypes(bytes, { defaults = [], overrides = [] }) {
  if (defaults.length === 0 && overrides.length === 0) return bytes;
  const entries = unzipSync(bytes);
  let xml = new TextDecoder().decode(entries['[Content_Types].xml']);
  for (const [extension, type] of defaults) {
    const line = `<Default Extension="${extension}" ContentType="${type}"/>`;
    const pattern = new RegExp(`<Default Extension="${escapeRegExp(extension)}"[^>]*/>`);
    xml = pattern.test(xml) ? xml.replace(pattern, line) : xml.replace('</Types>', `  ${line}\n</Types>`);
  }
  for (const [part, type] of overrides) {
    const line = `<Override PartName="/${part}" ContentType="${type}"/>`;
    const pattern = new RegExp(`<Override PartName="/${escapeRegExp(part)}"[^>]*/>`);
    xml = pattern.test(xml) ? xml.replace(pattern, line) : xml.replace('</Types>', `  ${line}\n</Types>`);
  }
  entries['[Content_Types].xml'] = new TextEncoder().encode(xml);
  return zipSync(entries, { level: 6, mtime: FIXED_MTIME });
}

function buildPackage({ workbook, workbookRels, sheets, sharedStrings, extra = {}, binary = {}, defaults = [], overrides = [] }) {
  const base = packageEntries({ workbook, workbookRels, sheets, sharedStrings, extra });
  return patchContentTypes(addBinaryParts(base, binary), { defaults, overrides });
}

function simpleWorkbook({ sheetName = 'Sheet1', calcPr = '' } = {}) {
  return `${XML_HEADER}
<workbook xmlns="${SPREADSHEET_NS}" xmlns:r="${REL_NS}">
  <workbookPr date1904="0"/>
  <bookViews><workbookView activeTab="0"/></bookViews>
  <sheets><sheet name="${sheetName}" sheetId="1" r:id="rId1"/></sheets>
  ${calcPr}
</workbook>`;
}

function relationships(entries) {
  return `${XML_HEADER}
<Relationships xmlns="${PACKAGE_REL_NS}">
${entries.map(([id, type, target]) => `  <Relationship Id="${id}" Type="${type}" Target="${target}"/>`).join('\n')}
</Relationships>`;
}

// ---------------------------------------------------------------------------
// F9 charts-images.xlsx — chart, image and shape parts are byte-identical
// after an unrelated edit; the dirty worksheet keeps its unmodelled children
// and their order; an insert above the charted range rewrites the series
// (file-compat §4, §9 F9, §10.15-16).
// ---------------------------------------------------------------------------

const f9Sheet = `${XML_HEADER}
<worksheet xmlns="${SPREADSHEET_NS}" xmlns:r="${REL_NS}" xmlns:mc="${MC_NS}" xmlns:x14="${X14_NS}" mc:Ignorable="x14">
  <dimension ref="A1:D4"/>
  <sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <cols><col min="1" max="2" width="14" customWidth="1"/><col min="4" max="4" width="10" customWidth="1"/></cols>
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Month</t></is></c><c r="B1" t="inlineStr"><is><t>Revenue</t></is></c><c r="D1" t="inlineStr"><is><t>Scratch</t></is></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>Jan</t></is></c><c r="B2"><v>10</v></c><c r="D2"><v>0</v></c></row>
    <row r="3"><c r="A3" t="inlineStr"><is><t>Feb</t></is></c><c r="B3"><v>20</v></c></row>
    <row r="4"><c r="A4" t="inlineStr"><is><t>Mar</t></is></c><c r="B4"><v>30</v></c></row>
  </sheetData>
  <drawing r:id="rId1"/>
  <legacyDrawing r:id="rId2"/>
  <oleObjects><oleObject progId="Package" shapeId="1025" r:id="rId3"/></oleObjects>
  <extLst>
    <ext uri="{5C3D1A0B-6E2F-4C7B-9A11-CALCPLUSF9001}" xmlns:cp="${CUSTOM_PROP_NS}">
      <cp:properties>
        <cp:property name="calcplus.chartFixture" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2">
          <vt:lpwstr xmlns:vt="${VT_NS}">preserve-chart-me</vt:lpwstr>
        </cp:property>
      </cp:properties>
    </ext>
  </extLst>
</worksheet>`;

const f9SheetRels = relationships([
  ['rId1', `${REL_NS}/drawing`, '../drawings/drawing1.xml'],
  ['rId2', `${REL_NS}/vmlDrawing`, '../drawings/vmlDrawing1.vml'],
  ['rId3', `${REL_NS}/oleObject`, '../embeddings/oleObject1.bin'],
]);

const f9Drawing = `${XML_HEADER}
<xdr:wsDr xmlns:xdr="${DRAWING_NS}" xmlns:a="${DRAWING_MAIN_NS}" xmlns:c="${CHART_NS}" xmlns:r="${REL_NS}">
  <xdr:twoCellAnchor editAs="twoCell">
    <xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>9</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
      <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
      <a:graphic><a:graphicData uri="${CHART_NS}"><c:chart r:id="rId1"/></a:graphicData></a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
  <xdr:twoCellAnchor editAs="oneCell">
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>12</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>20</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:pic>
      <xdr:nvPicPr><xdr:cNvPr id="3" name="Image 1"/><xdr:cNvPicPr/></xdr:nvPicPr>
      <xdr:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>
      <xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
  <xdr:twoCellAnchor editAs="oneCell">
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>22</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>26</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:sp macro="" textlink="">
      <xdr:nvSpPr><xdr:cNvPr id="4" name="Shape 1"/><xdr:cNvSpPr/></xdr:nvSpPr>
      <xdr:spPr><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom></xdr:spPr>
      <xdr:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>Shape</a:t></a:r></a:p></xdr:txBody>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>`;

const f9DrawingRels = relationships([
  ['rId1', `${REL_NS}/chart`, '../charts/chart1.xml'],
  ['rId2', `${REL_NS}/image`, '../media/image1.png'],
]);

const f9Chart = `${XML_HEADER}
<c:chartSpace xmlns:c="${CHART_NS}" xmlns:a="${DRAWING_MAIN_NS}" xmlns:r="${REL_NS}">
  <c:chart>
    <c:plotArea>
      <c:layout/>
      <c:barChart>
        <c:barDir val="col"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$4</c:f><c:strCache><c:ptCount val="3"/><c:pt idx="0"><c:v>Jan</c:v></c:pt><c:pt idx="1"><c:v>Feb</c:v></c:pt><c:pt idx="2"><c:v>Mar</c:v></c:pt></c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$B$2:$B$4</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt><c:pt idx="2"><c:v>30</c:v></c:pt></c:numCache></c:numRef></c:val>
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

/** A shape-only legacy drawing: no anchors to transform on a structural edit. */
const f9Vml = `${XML_HEADER}
<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
  <o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>
  <v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>
  <v:shape id="_x0000_s1025" type="#_x0000_t202" style="position:absolute;visibility:hidden"/>
</xml>`;

/** A 1x1 PNG: the smallest honest image part, so openpyxl can read it. */
const f9Png = new Uint8Array(
  Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
    'base64',
  ),
);

function f9Package() {
  return buildPackage({
    workbook: simpleWorkbook(),
    workbookRels: relationships([['rId1', `${REL_NS}/worksheet`, 'worksheets/sheet1.xml']]),
    sheets: [['xl/worksheets/sheet1.xml', f9Sheet]],
    extra: {
      'xl/worksheets/_rels/sheet1.xml.rels': f9SheetRels,
      'xl/drawings/drawing1.xml': f9Drawing,
      'xl/drawings/_rels/drawing1.xml.rels': f9DrawingRels,
      'xl/charts/chart1.xml': f9Chart,
      'xl/drawings/vmlDrawing1.vml': f9Vml,
    },
    binary: { 'xl/media/image1.png': f9Png, 'xl/embeddings/oleObject1.bin': binaryPayload('oleObject1', 512) },
    defaults: [['png', 'image/png']],
    overrides: [
      ['xl/charts/chart1.xml', CONTENT_TYPE.chart],
      ['xl/embeddings/oleObject1.bin', CONTENT_TYPE.oleObject],
    ],
  });
}

// ---------------------------------------------------------------------------
// F10 pivot-external.xlsx — pivot and external-link parts are preserved, a
// pivot source-range edit is blocked, external-link references stay untouched
// and the listing marks staleness (file-compat §6.2, §6.4, §9 F10).
// ---------------------------------------------------------------------------

const f10Sheet = `${XML_HEADER}
<worksheet xmlns="${SPREADSHEET_NS}" xmlns:r="${REL_NS}">
  <dimension ref="A1:H5"/>
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Region</t></is></c><c r="B1" t="inlineStr"><is><t>Product</t></is></c><c r="C1" t="inlineStr"><is><t>Amount</t></is></c><c r="F1" t="inlineStr"><is><t>Scratch</t></is></c><c r="G1" t="inlineStr"><is><t>Linked</t></is></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>North</t></is></c><c r="B2" t="inlineStr"><is><t>Widget</t></is></c><c r="C2"><v>10</v></c><c r="F2"><v>0</v></c><c r="G2" t="str"><v>North</v></c><c r="H2"><v>30</v></c></row>
    <row r="3"><c r="A3" t="inlineStr"><is><t>South</t></is></c><c r="B3" t="inlineStr"><is><t>Gadget</t></is></c><c r="C3"><v>20</v></c><c r="G3" t="str"><v>South</v></c><c r="H3"><v>30</v></c></row>
    <row r="4"><c r="A4" t="inlineStr"><is><t>North</t></is></c><c r="B4" t="inlineStr"><is><t>Gadget</t></is></c><c r="C4"><v>30</v></c><c r="G4" t="inlineStr"><is><t>Total</t></is></c><c r="H4"><v>60</v></c></row>
    <row r="5"><c r="A5" t="inlineStr"><is><t>South</t></is></c><c r="B5" t="inlineStr"><is><t>Widget</t></is></c><c r="C5"><v>40</v></c><c r="J5"><f>[1]Linked!$A$1</f><v>99</v></c></row>
  </sheetData>
</worksheet>`;

const f10SheetRels = relationships([
  ['rId1', `${REL_NS}/pivotTable`, '../pivotTables/pivotTable1.xml'],
]);

const f10PivotTable = `${XML_HEADER}
<pivotTableDefinition xmlns="${SPREADSHEET_NS}" xmlns:mc="${MC_NS}" xmlns:x14="${X14_NS}" mc:Ignorable="x14" name="PivotTable1" cacheId="1" dataCaption="Values" updatedVersion="6" minRefreshableVersion="3" createdVersion="6" indent="0" outline="1" outlineData="1" multipleFieldFilters="0">
  <location ref="G2:H4" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/>
  <pivotFields count="3">
    <pivotField axis="axisRow" showAll="0"><items count="2"><item x="0"/><item x="1"/></items></pivotField>
    <pivotField dataField="1" showAll="0"/>
    <pivotField dataField="1" showAll="0"/>
  </pivotFields>
  <rowFields count="1"><field x="0"/></rowFields>
  <rowItems count="2"><i><x/></i><i><x v="1"/></i></rowItems>
  <colItems count="1"><i/></colItems>
  <dataFields count="1"><dataField name="Sum of Amount" fld="2" baseField="0" baseItem="0"/></dataFields>
</pivotTableDefinition>`;

const f10PivotTableRels = relationships([
  ['rId1', `${REL_NS}/pivotCacheDefinition`, '../pivotCache/pivotCacheDefinition1.xml'],
]);

const f10PivotCache = `${XML_HEADER}
<pivotCacheDefinition xmlns="${SPREADSHEET_NS}" xmlns:r="${REL_NS}" r:id="rId1" refreshOnLoad="0" refreshedBy="calcplus fixtures" createdVersion="6" refreshedVersion="6" minRefreshableVersion="3" recordCount="4">
  <cacheSource type="worksheet"><worksheetSource ref="A1:C5" sheet="Sales"/></cacheSource>
  <cacheFields count="3">
    <cacheField name="Region" numFmtId="0"><sharedItems count="2"><s v="North"/><s v="South"/></sharedItems></cacheField>
    <cacheField name="Product" numFmtId="0"><sharedItems count="2"><s v="Widget"/><s v="Gadget"/></sharedItems></cacheField>
    <cacheField name="Amount" numFmtId="0"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="1" minValue="10" maxValue="40"/></cacheField>
  </cacheFields>
</pivotCacheDefinition>`;

const f10PivotCacheRels = relationships([
  ['rId1', `${REL_NS}/pivotCacheRecords`, 'pivotCacheRecords1.xml'],
]);

const f10PivotRecords = `${XML_HEADER}
<pivotCacheRecords xmlns="${SPREADSHEET_NS}" xmlns:r="${REL_NS}" count="4">
  <r><s v="North"/><s v="Widget"/><n v="10"/></r>
  <r><s v="South"/><s v="Gadget"/><n v="20"/></r>
  <r><s v="North"/><s v="Gadget"/><n v="30"/></r>
  <r><s v="South"/><s v="Widget"/><n v="40"/></r>
</pivotCacheRecords>`;

const f10ExternalLink = `${XML_HEADER}
<externalLink xmlns="${SPREADSHEET_NS}" xmlns:r="${REL_NS}">
  <externalBook r:id="rId1">
    <sheetNames><sheetName val="Linked"/></sheetNames>
    <sheetDataSet>
      <sheetData sheetId="0"><row r="1"><cell r="1" t="n"><v>99</v></cell></row></sheetData>
    </sheetDataSet>
  </externalBook>
</externalLink>`;

const f10ExternalLinkRels = `${XML_HEADER}
<Relationships xmlns="${PACKAGE_REL_NS}">
  <Relationship Id="rId1" Type="${REL_NS}/externalLinkPath" Target="book1.xlsx" TargetMode="External"/>
</Relationships>`;

function f10Package() {
  return buildPackage({
    workbook: simpleWorkbook({
      sheetName: 'Sales',
      calcPr:
        '<calcPr calcId="191029" fullCalcOnLoad="0"/>\n  <pivotCaches><pivotCache cacheId="1" r:id="rId2"/></pivotCaches>',
    }),
    workbookRels: relationships([
      ['rId1', `${REL_NS}/worksheet`, 'worksheets/sheet1.xml'],
      ['rId2', `${REL_NS}/pivotCacheDefinition`, 'pivotCache/pivotCacheDefinition1.xml'],
      ['rId3', `${REL_NS}/externalLink`, 'externalLinks/externalLink1.xml'],
    ]),
    sheets: [['xl/worksheets/sheet1.xml', f10Sheet]],
    extra: {
      'xl/worksheets/_rels/sheet1.xml.rels': f10SheetRels,
      'xl/pivotTables/pivotTable1.xml': f10PivotTable,
      'xl/pivotTables/_rels/pivotTable1.xml.rels': f10PivotTableRels,
      'xl/pivotCache/pivotCacheDefinition1.xml': f10PivotCache,
      'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels': f10PivotCacheRels,
      'xl/pivotCache/pivotCacheRecords1.xml': f10PivotRecords,
      'xl/externalLinks/externalLink1.xml': f10ExternalLink,
      'xl/externalLinks/_rels/externalLink1.xml.rels': f10ExternalLinkRels,
    },
    overrides: [
      ['xl/pivotTables/pivotTable1.xml', CONTENT_TYPE.pivotTable],
      ['xl/pivotCache/pivotCacheDefinition1.xml', CONTENT_TYPE.pivotCacheDefinition],
      ['xl/pivotCache/pivotCacheRecords1.xml', CONTENT_TYPE.pivotCacheRecords],
      ['xl/externalLinks/externalLink1.xml', CONTENT_TYPE.externalLink],
    ],
  });
}

// ---------------------------------------------------------------------------
// F11 macros.xlsm — vbaProject.bin is byte-identical, never executed, and
// `.xlsm` -> `.xlsx` requires confirmation (file-compat §6.3, §9 F11).
// ---------------------------------------------------------------------------

const f11Sheet = `${XML_HEADER}
<worksheet xmlns="${SPREADSHEET_NS}">
  <dimension ref="A1:B2"/>
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Input</t></is></c><c r="B1" t="inlineStr"><is><t>Doubled</t></is></c></row>
    <row r="2"><c r="A2"><v>10</v></c><c r="B2"><f>A2*2</f><v>20</v></c></row>
  </sheetData>
</worksheet>`;

function f11Package() {
  return buildPackage({
    workbook: simpleWorkbook({ calcPr: '<calcPr calcId="191029" fullCalcOnLoad="0"/>' }),
    workbookRels: relationships([
      ['rId1', `${REL_NS}/worksheet`, 'worksheets/sheet1.xml'],
      [
        'rId2',
        'http://schemas.microsoft.com/office/2006/relationships/vbaProject',
        'vbaProject.bin',
      ],
    ]),
    sheets: [['xl/worksheets/sheet1.xml', f11Sheet]],
    binary: { 'xl/vbaProject.bin': binaryPayload('vbaProject', 1024) },
    overrides: [
      ['xl/workbook.xml', CONTENT_TYPE.macroWorkbook],
      ['xl/vbaProject.bin', CONTENT_TYPE.vbaProject],
    ],
  });
}

/**
 * Deterministic pseudo-binary payloads: a marker string followed by a fixed
 * byte pattern, so the parts are inert, compress poorly enough to stay well
 * under the package ratio cap, and never execute.
 */
function binaryPayload(marker, length) {
  const markerBytes = new TextEncoder().encode(`calcplus ${marker} fixture; inert bytes. `);
  const bytes = new Uint8Array(length);
  for (let index = 0; index < length; index += 1) {
    bytes[index] = index < markerBytes.length ? markerBytes[index] : (index * 37) % 251;
  }
  return bytes;
}

export function build({ compatDir }) {
  const fixtures = [
    ['charts-images.xlsx', f9Package()],
    ['pivot-external.xlsx', f10Package()],
    ['macros.xlsm', f11Package()],
  ];
  for (const [name, bytes] of fixtures) {
    const path = join(compatDir, name);
    writeFileSync(path, bytes);
    console.log(`Wrote ${path}`);
  }
}
