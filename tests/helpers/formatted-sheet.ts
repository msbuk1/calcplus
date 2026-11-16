import { buildXlsx } from './build-xlsx';

/**
 * The M4A formatted-sheet fixture.
 *
 * Every imported formatting feature the renderer owns appears at least once:
 * font weight/style/underline/strike/colour, solid and pattern fills, border
 * weights and styles, horizontal and vertical alignment, indent, wrap,
 * shrink-to-fit, rotation, number-format colours, merged ranges, custom row
 * heights and column widths, hidden rows and columns, frozen panes, and a
 * narrow column that must paint the `#` overflow marker.
 *
 * It is built at test time so the visual baseline does not depend on a binary
 * fixture whose styles would drift from the style importer.
 */
export function formattedSheetXlsx(): Uint8Array {
  const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="4">
    <numFmt numFmtId="164" formatCode="#,##0.00"/>
    <numFmt numFmtId="165" formatCode="yyyy-mm-dd"/>
    <numFmt numFmtId="166" formatCode="[Red]0.00"/>
    <numFmt numFmtId="167" formatCode="0.0%"/>
  </numFmts>
  <fonts count="5">
    <font><sz val="11"/><name val="Calibri"/><family val="2"/></font>
    <font><b/><sz val="12"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>
    <font><i/><u/><sz val="11"/><color rgb="FFB00020"/><name val="Calibri"/><family val="2"/></font>
    <font><strike/><u val="double"/><sz val="14"/><name val="Calibri"/><family val="2"/></font>
    <font><sz val="10"/><color rgb="FF1F4E78"/><name val="Calibri"/><family val="2"/></font>
  </fonts>
  <fills count="4">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="lightGrid"><fgColor rgb="FFFFC000"/><bgColor rgb="FFFFFFFF"/></patternFill></fill>
  </fills>
  <borders count="3">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>
      <left style="thin"><color rgb="FFBFBFBF"/></left>
      <right style="thin"><color rgb="FFBFBFBF"/></right>
      <top style="thin"><color rgb="FFBFBFBF"/></top>
      <bottom style="thin"><color rgb="FFBFBFBF"/></bottom>
      <diagonal/>
    </border>
    <border>
      <left style="thin"><color rgb="FF808080"/></left>
      <right style="thin"><color rgb="FF808080"/></right>
      <top style="double"><color rgb="FF202020"/></top>
      <bottom style="medium"><color rgb="FF202020"/></bottom>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="13">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right"/></xf>
    <xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" textRotation="45"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1" indent="1"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="left" shrinkToFit="1"/></xf>
    <xf numFmtId="0" fontId="0" fillId="3" borderId="1" xfId="0" applyFill="1" applyBorder="1"/>
    <xf numFmtId="166" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="167" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="right"/></xf>
    <xf numFmtId="0" fontId="4" fillId="0" borderId="2" xfId="0" applyFont="1" applyBorder="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;

  const worksheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:F12"/>
  <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0">
      <pane xSplit="1" ySplit="2" topLeftCell="B3" activePane="bottomRight" state="frozen"/>
    </sheetView>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15" defaultColWidth="9"/>
  <cols>
    <col min="1" max="1" width="24" customWidth="1"/>
    <col min="2" max="2" width="16" customWidth="1"/>
    <col min="3" max="3" width="14" customWidth="1"/>
    <col min="4" max="4" width="34" customWidth="1"/>
    <col min="5" max="5" width="3" customWidth="1"/>
    <col min="6" max="6" width="12" hidden="1" customWidth="1"/>
  </cols>
  <sheetData>
    <row r="1" ht="30" customHeight="1">
      <c r="A1" t="inlineStr" s="1"><is><t>Quarterly report - formatted</t></is></c>
      <c r="B1" s="1"/><c r="C1" s="1"/><c r="D1" s="1"/>
    </row>
    <row r="2">
      <c r="A2" t="inlineStr" s="1"><is><t>Item</t></is></c>
      <c r="B2" t="inlineStr" s="1"><is><t>Amount</t></is></c>
      <c r="C2" t="inlineStr" s="1"><is><t>Date</t></is></c>
      <c r="D2" t="inlineStr" s="1"><is><t>Share</t></is></c>
    </row>
    <row r="3">
      <c r="A3" t="inlineStr"><is><t>Revenue</t></is></c>
      <c r="B3" s="4"><v>1234.5</v></c>
      <c r="C3" s="5"><v>46000</v></c>
      <c r="D3" s="11"><v>0.256</v></c>
    </row>
    <row r="4" ht="52" customHeight="1">
      <c r="A4" t="inlineStr" s="7"><is><t>Wrapped text breaks across several lines inside the cell.</t></is></c>
      <c r="B4" t="inlineStr" s="7"><is><t>Short</t></is></c>
      <c r="C4" s="10"><v>-1234.5</v></c>
      <c r="D4" t="inlineStr" s="7"><is><t>Indented, wrapped and top-aligned.</t></is></c>
    </row>
    <row r="5">
      <c r="A5" t="inlineStr" s="8"><is><t>Shrink-to-fit text that is much too long for this column</t></is></c>
      <c r="B5" t="inlineStr" s="6"><is><t>Rotate</t></is></c>
      <c r="C5" t="inlineStr" s="2"><is><t>Underline</t></is></c>
      <c r="D5" t="inlineStr" s="3"><is><t>Strike double</t></is></c>
    </row>
    <row r="6" hidden="1">
      <c r="A6" t="inlineStr"><is><t>Hidden row must not paint</t></is></c>
      <c r="B6"><v>999</v></c>
    </row>
    <row r="7">
      <c r="A7" t="inlineStr" s="9"><is><t>Pattern fill</t></is></c>
    </row>
    <row r="8">
      <c r="A8" t="inlineStr"><is><t>Underline</t></is></c>
      <c r="B8" t="inlineStr" s="2"><is><t>Italic red</t></is></c>
      <c r="C8" t="inlineStr" s="3"><is><t>Double</t></is></c>
    </row>
    <row r="9">
      <c r="A9" t="inlineStr"><is><t>Red negative</t></is></c>
      <c r="B9" s="10"><v>-0.5</v></c>
    </row>
    <row r="10">
      <c r="A10" t="inlineStr"><is><t>Narrow column</t></is></c>
      <c r="E10"><v>123456.78</v></c>
    </row>
    <row r="11">
      <c r="A11" t="inlineStr" s="1"><is><t>Merged title clipped at the merged range edge with extra text</t></is></c>
      <c r="B11" s="1"/><c r="C11" s="1"/><c r="D11" s="1"/>
    </row>
    <row r="12">
      <c r="A12" t="inlineStr"><is><t>Hidden column</t></is></c>
      <c r="F12" t="inlineStr"><is><t>must not paint</t></is></c>
    </row>
  </sheetData>
  <mergeCells count="2"><mergeCell ref="A1:D1"/><mergeCell ref="A11:D11"/></mergeCells>
</worksheet>`;

  return buildXlsx(worksheetXml, { sheetName: 'Formatted', stylesXml });
}
