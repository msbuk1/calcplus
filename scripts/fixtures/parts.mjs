import { zipSync } from 'fflate';

export const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
export const SPREADSHEET_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
export const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export const PACKAGE_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

export const contentTypes = (overrides) => `${XML_HEADER}
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>
  ${overrides.map(([part, type]) => `<Override PartName="/${part}" ContentType="${type}"/>`).join('\n  ')}
</Types>`;

export const rootRels = `${XML_HEADER}
<Relationships xmlns="${PACKAGE_REL_NS}">
  <Relationship Id="rId1" Type="${REL_NS}/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="${REL_NS}/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;

export const docPropsCore = `${XML_HEADER}
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>calcplus fixtures</dc:creator>
  <cp:lastModifiedBy>calcplus fixtures</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">2026-11-14T00:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2026-11-14T00:00:00Z</dcterms:modified>
</cp:coreProperties>`;

export const docPropsApp = `${XML_HEADER}
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>calcplus fixture generator</Application>
</Properties>`;

export const themeXml = `${XML_HEADER}
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

export const stylesXml = `${XML_HEADER}
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

export const sharedStringsXml = `${XML_HEADER}
<sst xmlns="${SPREADSHEET_NS}" count="4" uniqueCount="4">
  <si><t>Name</t></si>
  <si><t>Value</t></si>
  <si><t>Alpha</t></si>
  <si><t>Beta</t></si>
</sst>`;

export const workbookXml = `${XML_HEADER}
<workbook xmlns="${SPREADSHEET_NS}" xmlns:r="${REL_NS}">
  <workbookPr date1904="0"/>
  <bookViews><workbookView activeTab="0"/></bookViews>
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
  <definedNames><definedName name="FixtureRange">Sheet1!$A$1:$B$2</definedName></definedNames>
  <calcPr calcId="191029" fullCalcOnLoad="0"/>
</workbook>`;

export const workbookRels = `${XML_HEADER}
<Relationships xmlns="${PACKAGE_REL_NS}">
  <Relationship Id="rId1" Type="${REL_NS}/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="${REL_NS}/styles" Target="styles.xml"/>
  <Relationship Id="rId3" Type="${REL_NS}/sharedStrings" Target="sharedStrings.xml"/>
  <Relationship Id="rId4" Type="${REL_NS}/theme" Target="theme/theme1.xml"/>
</Relationships>`;

export const sheetRels = `${XML_HEADER}
<Relationships xmlns="${PACKAGE_REL_NS}">
  <Relationship Id="rId1" Type="${REL_NS}/drawing" Target="../drawings/drawing1.xml"/>
  <Relationship Id="rId2" Type="${REL_NS}/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>
</Relationships>`;


export function packageEntries({ workbook, workbookRels, sheets, sharedStrings, extra = {} }) {
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

export function contentTypeFor(path) {
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
