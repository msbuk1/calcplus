import { writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { XML_HEADER, SPREADSHEET_NS, REL_NS, PACKAGE_REL_NS, packageEntries } from '../parts.mjs';

/**
 * M7 lane 1 fixtures: the print planner's inputs.
 *
 * - `print-areas.xlsx` — two print areas on one sheet, `printOptions`
 *   gridlines and headings, a merged range, a date and a currency cell, a
 *   bold/filled/bordered style, landscape with no `fitToPage` (so the planner
 *   reads `scale` and the model can add `sheetPr/pageSetUpPr`).
 *
 * Hand-authored with fixed zip timestamps, so `npm run fixtures` is
 * byte-identical on every run (groups/README.md). The assertions live in
 * tests/m7-print-planner*.test.ts.
 */

const CURRENCY = '164';

const PRINT_STYLES = `${XML_HEADER}
<styleSheet xmlns="${SPREADSHEET_NS}">
  <numFmts count="1"><numFmt numFmtId="${CURRENCY}" formatCode="#,##0.00"/></numFmts>
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/><family val="2"/></font>
    <font><b/><sz val="11"/><name val="Calibri"/><family val="2"/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>
      <left style="thick"><color rgb="FF000000"/></left>
      <right style="thick"><color rgb="FF000000"/></right>
      <top style="thick"><color rgb="FF000000"/></top>
      <bottom style="thick"><color rgb="FF000000"/></bottom>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="5">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="${CURRENCY}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;

const PRINT_WORKBOOK = `${XML_HEADER}
<workbook xmlns="${SPREADSHEET_NS}" xmlns:r="${REL_NS}">
  <workbookPr date1904="0"/>
  <bookViews><workbookView activeTab="0"/></bookViews>
  <sheets><sheet name="Sales" sheetId="1" r:id="rId1"/></sheets>
  <definedNames>
    <definedName name="_xlnm.Print_Area" localSheetId="0">'Sales'!$A$1:$C$30,'Sales'!$E$1:$F$6</definedName>
  </definedNames>
  <calcPr calcId="191029" fullCalcOnLoad="0"/>
</workbook>`;

const PRINT_WORKBOOK_RELS = `${XML_HEADER}
<Relationships xmlns="${PACKAGE_REL_NS}">
  <Relationship Id="rId1" Type="${REL_NS}/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="${REL_NS}/styles" Target="styles.xml"/>
</Relationships>`;

function row(index, cells) {
  return `    <row r="${index}">${cells}</row>`;
}

function textCell(ref, value, style = '') {
  const styleAttribute = style === '' ? '' : ` s="${style}"`;
  return `<c r="${ref}"${styleAttribute} t="inlineStr"><is><t>${value}</t></is></c>`;
}

function numberCell(ref, value, style = '') {
  const styleAttribute = style === '' ? '' : ` s="${style}"`;
  return `<c r="${ref}"${styleAttribute}><v>${value}</v></c>`;
}

const PRINT_ROWS = [
  row(1, [textCell('A1', 'Region'), textCell('B1', 'Quarter'), textCell('C1', 'Amount')].join('')),
  row(2, [textCell('A2', 'North', '4'), numberCell('C2', 1200, '3')].join('')),
  row(3, [textCell('A3', 'South'), textCell('B3', 'Q1'), numberCell('C3', 900, '3')].join('')),
  row(4, [textCell('A4', 'East'), textCell('B4', 'Q2'), numberCell('C4', 750.5, '3')].join('')),
  row(5, [textCell('A5', 'West'), textCell('B5', 'Q3'), numberCell('C5', 1500, '3')].join('')),
  row(6, [textCell('A6', 'Total', '1'), textCell('B6', '', '1'), numberCell('C6', 4350.5, '1')].join('')),
  row(7, [textCell('A7', 'Date'), numberCell('C7', 46346, '2')].join('')),
];
// The first print area is 30 rows deep: the empty rows 8-30 are part of the
// area, so the planner must carry them even though only rows 1-7 hold cells.
const PRINT_SHEET = `${XML_HEADER}
<worksheet xmlns="${SPREADSHEET_NS}">
  <dimension ref="A1:F30"/>
  <sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="15" defaultColWidth="11"/>
  <sheetData>
${PRINT_ROWS.join('\n')}
  </sheetData>
  <mergeCells count="1"><mergeCell ref="A2:B2"/></mergeCells>
  <printOptions gridLines="1" gridLinesSet="1" headings="1"/>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
  <pageSetup paperSize="9" orientation="landscape" scale="100" fitToWidth="1" fitToHeight="0"/>
</worksheet>`;

function printAreasPackage() {
  return packageEntries({
    workbook: PRINT_WORKBOOK,
    workbookRels: PRINT_WORKBOOK_RELS,
    sheets: [['xl/worksheets/sheet1.xml', PRINT_SHEET]],
    extra: { 'xl/styles.xml': PRINT_STYLES },
  });
}

export function build({ compatDir }) {
  const fixtures = [['print-areas.xlsx', printAreasPackage()]];
  for (const [name, bytes] of fixtures) {
    const path = join(compatDir, name);
    writeFileSync(path, bytes);
    console.log(`Wrote ${path}`);
  }
}
