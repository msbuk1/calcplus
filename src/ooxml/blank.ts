import { openWorkbook, type OpenedWorkbook } from './open';
import { XML_DECLARATION_UTF8 } from './xml';
import { writePackage } from './zip';

const SPREADSHEET_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PACKAGE_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const RELS_TYPE = 'application/vnd.openxmlformats-package.relationships+xml';
const XML_TYPE = 'application/xml';

/**
 * The M1 new-workbook template. It is a real OOXML package that goes through
 * the same `openWorkbook` path as a user file, so there is no second model
 * construction path to drift.
 */
export function createBlankWorkbookBytes(): Uint8Array {
  const workbook = `${XML_DECLARATION_UTF8}
<workbook xmlns="${SPREADSHEET_NS}" xmlns:r="${REL_NS}">
  <workbookPr date1904="0"/>
  <bookViews><workbookView activeTab="0"/></bookViews>
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
  <calcPr calcId="191029" fullCalcOnLoad="0"/>
</workbook>`;
  const sheet = `${XML_DECLARATION_UTF8}
<worksheet xmlns="${SPREADSHEET_NS}" xmlns:r="${REL_NS}">
  <dimension ref="A1"/>
  <sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetData/>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>`;
  const styles = `${XML_DECLARATION_UTF8}
<styleSheet xmlns="${SPREADSHEET_NS}">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
  <dxfs count="0"/>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>`;
  const workbookRels = `${XML_DECLARATION_UTF8}
<Relationships xmlns="${PACKAGE_REL_NS}">
  <Relationship Id="rId1" Type="${REL_NS}/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="${REL_NS}/styles" Target="styles.xml"/>
</Relationships>`;
  const rootRels = `${XML_DECLARATION_UTF8}
<Relationships xmlns="${PACKAGE_REL_NS}">
  <Relationship Id="rId1" Type="${REL_NS}/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;
  const contentTypes = `${XML_DECLARATION_UTF8}
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="${RELS_TYPE}"/>
  <Default Extension="xml" ContentType="${XML_TYPE}"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`;

  return writePackage(
    new Map([
      ['[Content_Types].xml', encode(contentTypes)],
      ['_rels/.rels', encode(rootRels)],
      ['xl/workbook.xml', encode(workbook)],
      ['xl/_rels/workbook.xml.rels', encode(workbookRels)],
      ['xl/worksheets/sheet1.xml', encode(sheet)],
      ['xl/styles.xml', encode(styles)],
    ]),
  );
}

export function openBlankWorkbook(): OpenedWorkbook {
  return openWorkbook(createBlankWorkbookBytes(), 'Untitled.xlsx');
}

function encode(xml: string): Uint8Array {
  return new TextEncoder().encode(xml);
}
