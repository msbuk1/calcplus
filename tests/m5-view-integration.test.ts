import { describe, expect, it } from 'vitest';
import { openWorkbook } from '@ooxml/open';
import { createSheetProjection } from '@view/projection';
import type { AutoFilter } from '@model/types';
import { textValue } from '@model/value';
import { buildXlsx } from './helpers/build-xlsx';

/**
 * The projection seams that complete conditional formatting and filtering:
 * a `dxfId` rule resolves its differential format from the styles part and
 * paints, and a filter-hidden row hides in the projection (sort-autofilter
 * §3.3), not only in the document state.
 */

const STYLES = `<?xml version="1.0"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf/></cellStyleXfs>
  <cellXfs count="1"><xf/></cellXfs>
  <dxfs count="1"><dxf><fill><patternFill><bgColor rgb="FFFF0000"/></patternFill></fill></dxf></dxfs>
</styleSheet>`;

describe('projection integration seams', () => {
  it('resolves a dxfId rule through the styles table and paints it', () => {
    const bytes = buildXlsx(
      `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>10</v></c></row>
    <row r="2"><c r="A2"><v>1</v></c></row>
  </sheetData>
  <conditionalFormatting sqref="A1:A2">
    <cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>5</formula></cfRule>
  </conditionalFormatting>
</worksheet>`,
      { stylesXml: STYLES },
    );
    const opened = openWorkbook(bytes);
    const sheet = opened.workbook.sheets[0]!;
    const projection = createSheetProjection(opened.workbook, sheet.id, { version: 1 });

    const painted = projection.conditionalFormatting?.evaluate(0, 0);
    expect(painted).not.toBeNull();
    // Excel's dxf fills carry their colour in bgColor; the painter folds
    // color ?? backgroundColor (cf-paint.ts), so the overlay must too.
    expect(painted?.overlay.fill?.color ?? painted?.overlay.fill?.backgroundColor).toBe('#FF0000');
    // Below the threshold nothing paints.
    expect(projection.conditionalFormatting?.evaluate(1, 0)?.painted ?? true).toBe(false);
  });

  it('hides filter-derived rows in the projection and keeps manual rows apart', () => {
    const bytes = buildXlsx(
      `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>keep</t></is></c></row>
    <row r="3"><c r="A3" t="inlineStr"><is><t>drop</t></is></c></row>
    <row r="4" hidden="1"><c r="A4" t="inlineStr"><is><t>keep</t></is></c></row>
  </sheetData>
</worksheet>`,
    );
    const opened = openWorkbook(bytes);
    const sheet = opened.workbook.sheets[0]!;
    // A command-written filter keeps rowDimensions manual-only; the filter
    // layer is derived at paint time (sort-autofilter §3.3).
    const filter: AutoFilter = {
      sheetId: sheet.id,
      range: { startRow: 0, startCol: 0, endRow: 3, endCol: 0 },
      columns: new Map([
        [0, { kind: 'values', selected: [textValue('keep')], includeBlanks: false }],
      ]),
    };
    sheet.autoFilter = filter;
    const projection = createSheetProjection(opened.workbook, sheet.id, { version: 1 });

    expect(projection.isRowHidden(2)).toBe(true); // 'drop' fails the criterion.
    expect(projection.isRowHidden(1)).toBe(false); // 'keep' passes.
    expect(projection.isRowHidden(3)).toBe(true); // manually hidden stays hidden.

    // Clearing the filter reveals only the derived layer.
    sheet.autoFilter = null;
    const cleared = createSheetProjection(opened.workbook, sheet.id, { version: 2 });
    expect(cleared.isRowHidden(2)).toBe(false);
    expect(cleared.isRowHidden(3)).toBe(true);
  });
});
