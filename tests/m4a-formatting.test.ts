import { describe, expect, it } from 'vitest';
import { DocumentSession } from '@commands/session';
import type { GridRange } from '@model/types';
import { createCell, getCell, setCellSnapshot } from '@model/types';
import {
  builtinNumberFormat,
  builtinNumberFormatId,
  isLocaleBuiltinNumberFormat,
} from '@model/style';
import { numberValue, textValue } from '@model/value';
import { createEmptyWorkbook, createSheet } from '@model/workbook';
import { openBlankWorkbook } from '@ooxml/blank';
import { openWorkbook } from '@ooxml/open';
import { saveWorkbook } from '@ooxml/save';
import { createSheetProjection } from '@view/projection';
import {
  formatForLayout,
  formatValue,
  numberFormatColor,
  numberFormatColorForText,
  normalizeFormatCode,
} from '@view/number-format';
import { buildXlsx, unzipPart } from './helpers/build-xlsx';
import { formattedSheetXlsx } from './helpers/formatted-sheet';

const XL = { widthPx: 200, measure: (text: string) => text.length * 7, date1904: false };

function singleCell(): GridRange {
  return { startRow: 0, startCol: 0, endRow: 0, endCol: 0 };
}

function blankSession(): {
  opened: ReturnType<typeof openBlankWorkbook>;
  session: DocumentSession;
  sheetId: string;
} {
  const opened = openBlankWorkbook();
  const session = new DocumentSession(opened.workbook);
  return { opened, session, sheetId: opened.workbook.sheets[0]!.id };
}

describe('M4A style model', () => {
  it('reads underline variants, pattern backgrounds, border weights and shrink-to-fit', () => {
    const opened = openWorkbook(formattedSheetXlsx());
    const styles = opened.workbook.styles;
    expect(styles[2]!.font.underlineStyle).toBe('single');
    expect(styles[2]!.font.color).toBe('#B00020');
    expect(styles[3]!.font.underlineStyle).toBe('double');
    expect(styles[3]!.font.strike).toBe(true);
    expect(styles[9]!.fill.pattern).toBe('lightGrid');
    expect(styles[9]!.fill.color).toBe('#FFC000');
    expect(styles[9]!.fill.backgroundColor).toBe('#FFFFFF');
    expect(styles[1]!.border.top.color).toBe('#BFBFBF');
    expect(styles[1]!.border.top.weight).toBe('thin');
    expect(styles[12]!.border.top.weight).toBe('thick');
    expect(styles[12]!.border.bottom.weight).toBe('medium');
    expect(styles[8]!.alignment.shrinkToFit).toBe(true);
    expect(styles[7]!.alignment.wrapText).toBe(true);
  });

  it('distinguishes built-in, custom and unknown number-format references', () => {
    const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="${ns}">
  <numFmts count="1"><numFmt numFmtId="164" formatCode="0.000"/></numFmts>
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="4">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="200" fontId="0" fillId="0" borderId="0" xfId="0"/>
  </cellXfs>
</styleSheet>`;
    const worksheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="${ns}"><sheetData><row r="1"><c r="A1" s="1"><v>45000</v></c><c r="B1" s="2"><v>1.5</v></c><c r="C1" s="3"><v>1</v></c></row></sheetData></worksheet>`;
    const opened = openWorkbook(buildXlsx(worksheet, { stylesXml }));
    expect(opened.workbook.styles[1]!.numFmtKind).toBe('builtin');
    expect(opened.workbook.styles[1]!.numFmtCode).toBe('m/d/yy');
    expect(opened.workbook.styles[2]!.numFmtKind).toBe('custom');
    expect(opened.workbook.styles[2]!.numFmtCode).toBe('0.000');
    expect(opened.workbook.styles[3]!.numFmtKind).toBe('unknown');
    expect(opened.workbook.styles[3]!.numFmtId).toBe(200);
  });

  it('exposes imported merges, dimensions, defaults and hidden state to the projection', () => {
    const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    const worksheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="${ns}">
  <sheetFormatPr defaultRowHeight="15" defaultColWidth="9"/>
  <cols>
    <col min="2" max="3" width="18" hidden="1" bestFit="1" customWidth="1"/>
  </cols>
  <sheetData>
    <row r="1" ht="22" customHeight="1"><c r="A1"><v>1</v></c></row>
    <row r="2" hidden="1"><c r="A2"><v>2</v></c></row>
  </sheetData>
  <mergeCells count="1"><mergeCell ref="A1:B2"/></mergeCells>
</worksheet>`;
    const opened = openWorkbook(buildXlsx(worksheet));
    const projection = createSheetProjection(opened.workbook, opened.initialSheetId, {
      version: 1,
    });
    expect(projection.defaultRowHeightPoints).toBe(15);
    expect(projection.defaultColumnWidthChars).toBe(9);
    expect(projection.rowHeight(0)).toBe(29);
    expect(projection.rowDimension(0)?.customHeight).toBe(true);
    expect(projection.isRowHidden(1)).toBe(true);
    expect(projection.rowHeight(1)).toBe(0);
    expect(projection.isColumnHidden(1)).toBe(true);
    expect(projection.columnWidth(1)).toBe(0);
    expect(projection.columnDimension(1)).toMatchObject({
      width: 18,
      hidden: true,
      bestFit: true,
      customWidth: true,
    });
    expect(projection.mergeAt(0, 0)).toEqual({
      anchorRow: 0,
      anchorCol: 0,
      rowSpan: 2,
      colSpan: 2,
    });
  });
});

describe('M4A formatting commands', () => {
  it('applies outline and inside border presets per cell', () => {
    const { session, sheetId } = blankSession();
    const range: GridRange = { startRow: 0, startCol: 0, endRow: 1, endCol: 1 };
    session.dispatch(
      'format.border',
      { sheetId, range, preset: 'outline', style: 'medium', color: '#FF0000' },
      { activeSheetId: sheetId, selection: range, activeCell: { row: 0, col: 0 } },
    );
    const sheet = session.workbook.sheets[0]!;
    const a1 = session.workbook.styles[getCell(sheet, 0, 0)!.styleId]!;
    const b1 = session.workbook.styles[getCell(sheet, 0, 1)!.styleId]!;
    const b2 = session.workbook.styles[getCell(sheet, 1, 1)!.styleId]!;
    expect(a1.border.top.style).toBe('medium');
    expect(a1.border.left.weight).toBe('medium');
    expect(a1.border.right.style).toBe('none');
    expect(a1.border.bottom.style).toBe('none');
    expect(b1.border.top.style).toBe('medium');
    expect(b1.border.right.style).toBe('medium');
    expect(b1.border.left.style).toBe('none');
    expect(b2.border.bottom.style).toBe('medium');
    expect(b2.border.right.style).toBe('medium');

    const insideSession = blankSession();
    insideSession.session.dispatch(
      'format.border',
      { sheetId: insideSession.sheetId, range, preset: 'inside' },
      {
        activeSheetId: insideSession.sheetId,
        selection: range,
        activeCell: { row: 0, col: 0 },
      },
    );
    const insideSheet = insideSession.session.workbook.sheets[0]!;
    const insideA1 = insideSession.session.workbook.styles[getCell(insideSheet, 0, 0)!.styleId]!;
    expect(insideA1.border.right.style).toBe('thin');
    expect(insideA1.border.bottom.style).toBe('thin');
    expect(insideA1.border.top.style).toBe('none');
    expect(insideA1.border.left.style).toBe('none');
  });

  it('toggles each cell in a mixed range from its own current style', () => {
    const workbook = createEmptyWorkbook();
    const sheet = createSheet('Sheet1', 'xl/worksheets/sheet1.xml');
    workbook.sheets.push(sheet);
    sheet.cells.set('0:0', createCell(textValue('plain')));
    workbook.styles.push({
      id: 1,
      numFmtId: 0,
      fontId: 0,
      fillId: 0,
      borderId: 0,
      xfId: 0,
      rawXml: '',
      font: {
        name: 'Calibri',
        size: 11,
        bold: true,
        italic: false,
        underline: false,
        underlineStyle: 'none',
        strike: false,
        color: null,
      },
      fill: { pattern: 'none', color: null, backgroundColor: null },
      border: {
        left: { style: 'none', color: null, weight: null },
        right: { style: 'none', color: null, weight: null },
        top: { style: 'none', color: null, weight: null },
        bottom: { style: 'none', color: null, weight: null },
        diagonal: { style: 'none', color: null, weight: null },
        diagonalUp: false,
        diagonalDown: false,
      },
      alignment: {
        horizontal: 'general',
        vertical: 'bottom',
        wrapText: false,
        indent: 0,
        textRotation: 0,
        shrinkToFit: false,
      },
      numFmtCode: null,
      numFmtKind: 'general',
    });
    sheet.cells.set('0:1', { ...createCell(textValue('bold')), styleId: 1 });
    const session = new DocumentSession(workbook);
    const range: GridRange = { startRow: 0, startCol: 0, endRow: 0, endCol: 1 };
    session.dispatch(
      'format.bold',
      { sheetId: sheet.id, range },
      { activeSheetId: sheet.id, selection: range, activeCell: { row: 0, col: 0 } },
    );
    const first = workbook.styles[getCell(sheet, 0, 0)!.styleId]!;
    const second = workbook.styles[getCell(sheet, 0, 1)!.styleId]!;
    expect(first.font.bold).toBe(true);
    // The bold cell toggles off instead of being overwritten by the active cell.
    expect(second.font.bold).toBe(false);
  });
});

describe('M4A styles.xml round-trip', () => {
  it('appends fonts and cellXfs without renumbering existing styles', () => {
    const { opened, session, sheetId } = blankSession();
    const originalStyles = unzipPart(opened.package.originalBytes!, 'xl/styles.xml');
    const range = singleCell();
    session.dispatch(
      'format.bold',
      { sheetId, range },
      { activeSheetId: sheetId, selection: range, activeCell: { row: 0, col: 0 } },
    );
    const saved = saveWorkbook({
      workbook: session.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
    });
    expect(saved.noOp).toBe(false);
    const styles = unzipPart(saved.bytes, 'xl/styles.xml');
    expect(styles).toContain('<fonts count="2">');
    expect(styles).toContain('<cellXfs count="2">');
    expect(styles).toContain('<font><b/><sz val="11"/><name val="Calibri"/></font>');
    expect(styles).toContain(
      '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>',
    );
    // The first font keeps its position and bytes in the append-only table.
    const firstFontStart = originalStyles.indexOf('<font>');
    const firstFontEnd = originalStyles.indexOf('</font>', firstFontStart) + '</font>'.length;
    expect(styles).toContain(originalStyles.slice(firstFontStart, firstFontEnd));
    const reopened = openWorkbook(saved.bytes);
    expect(reopened.workbook.styles).toHaveLength(2);
    expect(reopened.workbook.styles[1]!.font.bold).toBe(true);
  });

  it('appends and reopens fills, border styles and alignment', () => {
    const { opened, session, sheetId } = blankSession();
    const range = singleCell();
    const context = {
      activeSheetId: sheetId,
      selection: range,
      activeCell: { row: 0, col: 0 },
    };
    session.dispatch(
      'format.fill',
      {
        sheetId,
        range,
        pattern: 'lightGrid',
        color: '#FFC000',
        backgroundColor: '#FFFFFF',
      },
      context,
    );
    session.dispatch(
      'format.border',
      { sheetId, range, style: 'medium', color: '#FF0000' },
      context,
    );
    session.dispatch(
      'format.alignment',
      {
        sheetId,
        range,
        horizontal: 'center',
        vertical: 'top',
        wrapText: true,
        indent: 2,
        textRotation: 45,
        shrinkToFit: true,
      },
      context,
    );
    const saved = saveWorkbook({
      workbook: session.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
    });
    expect(saved.noOp).toBe(false);
    const styles = unzipPart(saved.bytes, 'xl/styles.xml');
    expect(styles).toContain('<patternFill patternType="lightGrid">');
    expect(styles).toContain('<fgColor rgb="FFFFC000"/>');
    expect(styles).toContain('<bgColor rgb="FFFFFFFF"/>');
    expect(styles).toContain('<left style="medium"><color rgb="FFFF0000"/></left>');
    expect(styles).toContain('<right style="medium"><color rgb="FFFF0000"/></right>');
    expect(styles).toContain('<top style="medium"><color rgb="FFFF0000"/></top>');
    expect(styles).toContain('<bottom style="medium"><color rgb="FFFF0000"/></bottom>');
    expect(styles).toContain('applyFill="1"');
    expect(styles).toContain('applyBorder="1"');
    expect(styles).toContain('applyAlignment="1"');
    expect(styles).toContain('horizontal="center"');
    expect(styles).toContain('vertical="top"');
    expect(styles).toContain('wrapText="1"');
    expect(styles).toContain('shrinkToFit="1"');
    expect(styles).toContain('indent="2"');
    expect(styles).toContain('textRotation="45"');

    const reopened = openWorkbook(saved.bytes);
    const reopenedSheet = reopened.workbook.sheets[0]!;
    const styleId = getCell(reopenedSheet, 0, 0)!.styleId;
    const style = reopened.workbook.styles.find((candidate) => candidate.id === styleId)!;
    expect(style.fill).toEqual({
      pattern: 'lightGrid',
      color: '#FFC000',
      backgroundColor: '#FFFFFF',
    });
    expect(style.border).toEqual({
      left: { style: 'medium', color: '#FF0000', weight: 'medium' },
      right: { style: 'medium', color: '#FF0000', weight: 'medium' },
      top: { style: 'medium', color: '#FF0000', weight: 'medium' },
      bottom: { style: 'medium', color: '#FF0000', weight: 'medium' },
      diagonal: { style: 'none', color: null, weight: null },
      diagonalUp: false,
      diagonalDown: false,
    });
    expect(style.alignment).toEqual({
      horizontal: 'center',
      vertical: 'top',
      wrapText: true,
      indent: 2,
      textRotation: 45,
      shrinkToFit: true,
    });
  });

  it('appends a custom numFmt above 163 and keeps built-in ids built in', () => {
    const custom = blankSession();
    const range = singleCell();
    custom.session.dispatch(
      'format.numberFormat',
      { sheetId: custom.sheetId, range, code: '0.000' },
      { activeSheetId: custom.sheetId, selection: range, activeCell: { row: 0, col: 0 } },
    );
    const customSaved = saveWorkbook({
      workbook: custom.session.workbook,
      package: custom.opened.package,
      dirtyCells: custom.session.dirtyCellMap(),
    });
    const customStyles = unzipPart(customSaved.bytes, 'xl/styles.xml');
    expect(customStyles).toContain('<numFmts count="1">');
    expect(customStyles).toContain('<numFmt numFmtId="164" formatCode="0.000"/>');
    const customReopened = openWorkbook(customSaved.bytes);
    expect(customReopened.workbook.styles[1]!.numFmtId).toBe(164);
    expect(customReopened.workbook.styles[1]!.numFmtKind).toBe('custom');
    expect(customReopened.workbook.styles[1]!.numFmtCode).toBe('0.000');

    const builtin = blankSession();
    builtin.session.dispatch(
      'format.numberFormat',
      { sheetId: builtin.sheetId, range, code: '#,##0.00' },
      { activeSheetId: builtin.sheetId, selection: range, activeCell: { row: 0, col: 0 } },
    );
    const builtinSaved = saveWorkbook({
      workbook: builtin.session.workbook,
      package: builtin.opened.package,
      dirtyCells: builtin.session.dirtyCellMap(),
    });
    const builtinStyles = unzipPart(builtinSaved.bytes, 'xl/styles.xml');
    expect(builtinStyles).not.toContain('<numFmts');
    expect(builtinStyles).toContain('numFmtId="4"');
    const builtinReopened = openWorkbook(builtinSaved.bytes);
    expect(builtinReopened.workbook.styles[1]!.numFmtId).toBe(4);
    expect(builtinReopened.workbook.styles[1]!.numFmtKind).toBe('builtin');
  });

  it('preserves merges, dimensions and hidden state in an edited save', () => {
    const opened = openWorkbook(formattedSheetXlsx());
    const originalStyles = unzipPart(opened.package.originalBytes!, 'xl/styles.xml');
    const sheet = opened.workbook.sheets[0]!;
    setCellSnapshot(sheet, 2, 1, createCell(numberValue(9876.5)));
    const saved = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: new Map([[sheet.id, new Set(['2:1'])]]),
    });
    const worksheet = unzipPart(saved.bytes, 'xl/worksheets/sheet1.xml');
    expect(worksheet).toContain('<mergeCell ref="A1:D1"/>');
    expect(worksheet).toContain('<mergeCell ref="A11:D11"/>');
    expect(worksheet).toContain('width="24" customWidth="1"');
    expect(worksheet).toContain('hidden="1" customWidth="1"');
    expect(worksheet).toContain('<row r="1" ht="30" customHeight="1">');
    expect(worksheet).toContain('<row r="6" hidden="1">');
    expect(unzipPart(saved.bytes, 'xl/styles.xml')).toBe(originalStyles);
  });

  it('leaves styles.xml byte-identical for no-op and value-only saves', () => {
    const opened = openBlankWorkbook();
    const originalStyles = unzipPart(opened.package.originalBytes!, 'xl/styles.xml');
    const noOp = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: new Map(),
    });
    expect(noOp.noOp).toBe(true);
    expect(Buffer.from(noOp.bytes)).toEqual(Buffer.from(opened.package.originalBytes!));

    const sheet = opened.workbook.sheets[0]!;
    setCellSnapshot(sheet, 0, 0, createCell(textValue('hello')));
    const saved = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: new Map([[sheet.id, new Set(['0:0'])]]),
    });
    expect(unzipPart(saved.bytes, 'xl/styles.xml')).toBe(originalStyles);
    expect(unzipPart(saved.bytes, 'xl/worksheets/sheet1.xml')).toContain('hello');
  });
});

describe('M4A number formats', () => {
  it('renders built-in, custom, currency and date formats', () => {
    expect(formatValue(numberValue(1234.5), '0.00')).toBe('1234.50');
    expect(formatValue(numberValue(1234.5), '#,##0.00')).toBe('1,234.50');
    expect(formatValue(numberValue(-1234.5), '$#,##0.00_);($#,##0.00)')).toBe('($1,234.50)');
    expect(formatValue(numberValue(45_000), 'm/d/yy')).toBe('3/15/23');
    expect(formatValue(numberValue(45_000.5), 'm/d/yy h:mm')).toBe('3/15/23 12:00');
    expect(builtinNumberFormat(14)).toBe('m/d/yy');
    expect(builtinNumberFormat(59)).toBe('0');
    expect(builtinNumberFormat(39)).toBe('#,##0.00_);(#,##0.00)');
    expect(builtinNumberFormat(40)).toBe('#,##0.00_);[Red](#,##0.00)');
    expect(formatValue(numberValue(1234.56), '#,##0.00_);(#,##0.00)')).toBe('1,234.56\u0020');
    expect(isLocaleBuiltinNumberFormat(59)).toBe(true);
    expect(builtinNumberFormatId('#,##0.00')).toBe(4);
    expect(builtinNumberFormatId('#,##0.00_);(#,##0.00)')).toBe(39);
    expect(builtinNumberFormatId('#,##0.00_);[Red](#,##0.00)')).toBe(40);
  });

  it('resolves every built-in id in the ECMA table', () => {
    const ids = [
      0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29,
      30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52,
      53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79,
      80, 81,
    ];
    for (const id of ids) {
      const code = builtinNumberFormat(id);
      expect(code, `numFmtId ${id}`).not.toBeNull();
      const text = formatValue(numberValue(id >= 45 && id <= 48 ? 45_000.5 : 1234.5), code);
      expect(text.length, `numFmtId ${id}`).toBeGreaterThan(0);
    }
  });

  it('normalises lowercase e+ and a/p and strips NatNum/DBNum', () => {
    expect(formatValue(numberValue(12_200_000), '0.00e+00')).toBe('1.22E+07');
    expect(formatValue(numberValue(0.5222625), 'h:mm a/p')).toBe('12:32 P');
    expect(formatValue(numberValue(1234), '[NatNum1]0')).toBe('1234');
    expect(formatValue(numberValue(1234), '[DBNum1]0')).toBe('1234');
    expect(normalizeFormatCode('"e+"0.00e+00')).toBe('"e+"0.00E+00');
  });

  it('applies the 1904 adapter to calendar formats but not elapsed time', () => {
    expect(formatValue(numberValue(0), 'yyyy-mm-dd', true)).toBe('1904-01-01');
    expect(formatValue(numberValue(60), 'm/d/yy', false)).toBe('2/29/00');
    expect(formatValue(numberValue(60), 'm/d/yy', true)).toBe('3/1/04');
    expect(formatValue(numberValue(1.5), '[h]:mm:ss', true)).toBe('36:00:00');
  });

  it('falls back to General for an invalid code and never throws', () => {
    expect(formatValue(numberValue(123.4), '0.00;0.00;0.00;@;garbage')).toBe('123.4');
    expect(formatValue(numberValue(123.4), '[NatNum1]')).toBe('123.4');
  });

  it('resolves named and indexed colours for numeric and text sections', () => {
    expect(numberFormatColor('[Red]0.00')).toBe('#FF0000');
    expect(numberFormatColor('[Blue]0.00;[Red]-0.00')).toBe('#0000FF');
    expect(numberFormatColor('[Color 10]0.00')).toBe('#FF0000');
    expect(numberFormatColor('0.00')).toBeNull();
    expect(numberFormatColorForText('[Red]@')).toBe('#FF0000');
    expect(numberFormatColorForText('[Red]@;[Blue]0')).toBe('#FF0000');
    expect(numberFormatColorForText('[Blue]0;[Red]@')).toBe('#FF0000');
  });

  it('applies the text-section colour through the projection', () => {
    const { session, sheetId } = blankSession();
    const sheet = session.workbook.sheets[0]!;
    setCellSnapshot(sheet, 0, 0, createCell(textValue('red text')));
    const range = singleCell();
    session.dispatch(
      'format.numberFormat',
      { sheetId, range, code: '[Red]@;[Blue]0' },
      { activeSheetId: sheetId, selection: range, activeCell: { row: 0, col: 0 } },
    );
    const projection = createSheetProjection(session.workbook, sheetId, { version: 2 });
    expect(projection.cellAt(0, 0)?.style.numberFormatColor).toBe('#FF0000');
  });

  it('groups with the Indian 3/2 convention', () => {
    expect(formatValue(numberValue(1_234_567), '[$-4009]#,##0')).toBe('12,34,567');
  });

  it('renders localized format codes with English separators and month names', () => {
    expect(formatValue(numberValue(1234.5), '[$-407]#,##0.00')).toBe('1,234.50');
    expect(formatValue(numberValue(45_000), '[$-411]mmmm')).toBe('March');
    expect(formatValue(numberValue(1234.5), '[$€-407]#,##0.00')).toBe('€1,234.50');
  });

  it('exposes the width-aware display through the projection', () => {
    const { session, sheetId } = blankSession();
    const sheet = session.workbook.sheets[0]!;
    setCellSnapshot(sheet, 0, 0, createCell(numberValue(123_456.78)));
    setCellSnapshot(sheet, 1, 0, createCell(numberValue(-42)));
    const first = singleCell();
    const second: GridRange = { startRow: 1, startCol: 0, endRow: 1, endCol: 0 };
    const context = {
      activeSheetId: sheetId,
      selection: first,
      activeCell: { row: 0, col: 0 },
    };
    session.dispatch('format.numberFormat', { sheetId, range: first, code: '#,##0.00' }, context);
    session.dispatch(
      'format.numberFormat',
      {
        sheetId,
        range: second,
        code: '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)',
      },
      { ...context, selection: second, activeCell: { row: 1, col: 0 } },
    );
    const projection = createSheetProjection(session.workbook, sheetId, { version: 1 });

    const wide = projection.layoutTextAt(0, 0, 200, XL.measure);
    expect(wide).toEqual({ text: '123,456.78', overflow: false });
    const narrow = projection.layoutTextAt(0, 0, 20, XL.measure);
    expect(narrow?.overflow).toBe(true);
    expect(narrow?.text.startsWith('#')).toBe(true);
    expect(projection.cellAt(0, 0)?.layoutText(20, XL.measure)).toEqual(narrow);

    const accounting = projection.layoutTextAt(1, 0, 140, XL.measure);
    expect(accounting?.overflow).toBe(false);
    expect(accounting?.text.startsWith(' ')).toBe(true);
    expect(accounting?.text.length).toBeGreaterThan(6);
    expect(projection.layoutTextAt(9, 9, 100, XL.measure)).toBeNull();
  });

  it('lays out padding and overflow to the cell width', () => {
    const wide = formatForLayout(numberValue(123_456.78), '#,##0.00', {
      ...XL,
      widthPx: 200,
    });
    expect(wide.text).toBe('123,456.78');
    expect(wide.overflow).toBe(false);
    const narrow = formatForLayout(numberValue(123_456.78), '#,##0.00', {
      ...XL,
      widthPx: 20,
    });
    expect(narrow.overflow).toBe(true);
    expect(narrow.text.startsWith('#')).toBe(true);
    const accounting = formatForLayout(
      numberValue(-42),
      '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)',
      {
        ...XL,
        widthPx: 140,
      },
    );
    expect(accounting.overflow).toBe(false);
    expect(accounting.text.length).toBeGreaterThan(6);
  });
});
