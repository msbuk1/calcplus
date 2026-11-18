import { describe, expect, it } from 'vitest';
import { openWorkbook } from '@ooxml/open';
import { buildXlsx } from './helpers/build-xlsx';

/**
 * OOXML xsd:boolean attributes accept `0`, `1`, `true` and `false`. Excel
 * omits false-valued attributes, but Java writers (Apache POI and friends)
 * spell every boolean out — `hidden="false"` on every column of a template.
 * Reading those as truthy hid whole sheets of columns: only columns without a
 * `<col>` entry rendered.
 */

const WORKSHEET = `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews><sheetView showGridLines="false" showHeadings="false" workbookViewId="0"/></sheetViews>
  <sheetFormatPr defaultColWidth="11.53515625" defaultRowHeight="15"/>
  <cols>
    <col collapsed="false" customWidth="true" hidden="false" max="1" min="1" style="0" width="9.78"/>
    <col collapsed="false" customWidth="false" hidden="true" max="2" min="2" style="0" width="4.22"/>
    <col collapsed="false" customWidth="true" hidden="1" max="3" min="3" style="0" width="8.95"/>
    <col collapsed="false" customWidth="true" hidden="false" max="4" min="4" style="0" width="4.49"/>
  </cols>
  <sheetData>
    <row r="1" customFormat="false" ht="15" hidden="false" customHeight="false"><c r="A1" t="inlineStr"><is><t>header</t></is></c></row>
    <row r="2" customFormat="false" ht="15" hidden="true" customHeight="true"><c r="A2" t="inlineStr"><is><t>hidden-true</t></is></c></row>
    <row r="3" customFormat="false" ht="15" hidden="1" customHeight="true"><c r="A3" t="inlineStr"><is><t>hidden-1</t></is></c></row>
  </sheetData>
</worksheet>`;

describe('xsd:boolean attributes on open', () => {
  const opened = openWorkbook(buildXlsx(WORKSHEET));
  const sheet = opened.workbook.sheets[0]!;

  it('does not hide columns whose hidden attribute spells out false', () => {
    expect(sheet.columnDimensions.get(0)?.hidden ?? false).toBe(false);
    expect(sheet.columnDimensions.get(0)?.width).toBe(9.78);
    expect(sheet.columnDimensions.get(3)?.hidden ?? false).toBe(false);
  });

  it('still hides columns declared hidden in either lexical form', () => {
    expect(sheet.columnDimensions.get(1)?.hidden).toBe(true);
    expect(sheet.columnDimensions.get(2)?.hidden).toBe(true);
  });

  it('reads customWidth=false as unset', () => {
    expect(sheet.columnDimensions.get(1)?.customWidth ?? false).toBe(false);
    expect(sheet.columnDimensions.get(0)?.customWidth).toBe(true);
  });

  it('reads row hidden in both true spellings and leaves hidden=false visible', () => {
    // Model rows are zero-based: file row 1 is key 0.
    expect(sheet.rowDimensions.get(0)?.hidden ?? false).toBe(false);
    expect(sheet.rowDimensions.get(1)?.hidden).toBe(true);
    expect(sheet.rowDimensions.get(2)?.hidden).toBe(true);
  });

  it('reads customHeight=false as unset', () => {
    expect(sheet.rowDimensions.get(0)?.customHeight ?? false).toBe(false);
    expect(sheet.rowDimensions.get(1)?.customHeight).toBe(true);
  });

  it('reads the sheet view toggles spelled out', () => {
    expect(sheet.view.showGridLines).toBe(false);
    expect(sheet.view.showHeadings).toBe(false);
  });
});
