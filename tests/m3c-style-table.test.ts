import { describe, expect, it } from 'vitest';
import { DEFAULT_ALIGNMENT, DEFAULT_FILL } from '@model/style';
import { findOrAppendStyle, patchedStyleInput } from '@model/style-table';
import { createDefaultStyle, createEmptyWorkbook } from '@model/workbook';
import { parseStylesXml } from '@ooxml/styles-xml';

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

describe('M3C style table', () => {
  it('appends new styles and reuses an identical record without renumbering', () => {
    const workbook = createEmptyWorkbook();
    const base = createDefaultStyle(0);
    const bold = patchedStyleInput(base, { font: { bold: true } });
    const first = findOrAppendStyle(workbook, bold);
    expect(first).toBe(1);
    expect(workbook.styles).toHaveLength(2);
    expect(workbook.styles[0]).toEqual(base);

    const again = findOrAppendStyle(workbook, bold);
    expect(again).toBe(1);
    expect(workbook.styles).toHaveLength(2);

    const red = findOrAppendStyle(
      workbook,
      patchedStyleInput(base, { font: { color: '#C00000' } }),
    );
    expect(red).toBe(2);
    expect(workbook.styles.map((style) => style.id)).toEqual([0, 1, 2]);
  });

  it('patches one part while keeping fill, alignment and number format', () => {
    const base = createDefaultStyle(0);
    const styled = {
      ...base,
      fill: { pattern: 'solid', color: '#FFF2CC' },
      alignment: { ...DEFAULT_ALIGNMENT, horizontal: 'center' as const },
      numFmtCode: '#,##0.00',
      numFmtId: 4,
    };
    const patched = patchedStyleInput(styled, { font: { italic: true } });
    expect(patched.font?.italic).toBe(true);
    expect(patched.fill).toEqual(styled.fill);
    expect(patched.alignment?.horizontal).toBe('center');
    expect(patched.numFmtCode).toBe('#,##0.00');
    expect(patched.numFmtId).toBe(4);
  });

  it('reads fonts, fills, borders, alignment and number-format references', () => {
    const stylesXml = `${XML_HEADER}
<styleSheet xmlns="${NS}">
  <numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0.00"/></numFmts>
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="12"/><color rgb="FFC00000"/><name val="Calibri"/></font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border diagonalUp="1"><left style="thin"><color rgb="FF0000FF"/></left><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="164" fontId="1" fillId="1" borderId="1" xfId="0" applyNumberFormat="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center" wrapText="1" indent="2" textRotation="45"/>
    </xf>
  </cellXfs>
</styleSheet>`;
    const parsed = parseStylesXml(new TextEncoder().encode(stylesXml), 'xl/styles.xml');
    expect(parsed.styles).toHaveLength(2);
    const style = parsed.styles[1]!;
    expect(style.font.bold).toBe(true);
    expect(style.font.size).toBe(12);
    expect(style.font.color).toBe('#C00000');
    expect(style.fill).toEqual({ pattern: 'solid', color: '#FFF2CC' });
    expect(style.border.left.style).toBe('thin');
    expect(style.border.left.color).toBe('#0000FF');
    expect(style.border.diagonalUp).toBe(true);
    expect(style.alignment.horizontal).toBe('center');
    expect(style.alignment.wrapText).toBe(true);
    expect(style.alignment.indent).toBe(2);
    expect(style.numFmtCode).toBe('#,##0.00');
    expect(style.numFmtId).toBe(164);
    expect(DEFAULT_FILL.pattern).toBe('none');
  });
});
