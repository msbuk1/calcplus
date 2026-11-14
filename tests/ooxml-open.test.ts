import { readFileSync } from 'node:fs';
import { zipSync, unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { openWorkbook } from '@ooxml/open';
import { WorkbookOpenError } from '@ooxml/errors';
import { buildXlsx } from './helpers/build-xlsx';

const fixture = (name: string): Uint8Array =>
  new Uint8Array(readFileSync(new URL(`../fixtures/compat/${name}`, import.meta.url)));

describe('open a minimal workbook', () => {
  it('maps the workbook, sparse cells, styles, merges and view state', () => {
    const opened = openWorkbook(fixture('m1-chart-unknown.xlsx'));
    const workbook = opened.workbook;
    expect(workbook.sheets).toHaveLength(1);
    const sheet = workbook.sheets[0]!;
    expect(sheet.name).toBe('Sheet1');
    expect(sheet.cells.size).toBe(10);
    expect(sheet.merges).toHaveLength(1);
    expect(sheet.merges[0]!.range).toEqual({
      startRow: 0,
      startCol: 0,
      endRow: 0,
      endCol: 1,
    });
    expect(sheet.view.freezeRows).toBe(1);
    expect(sheet.view.zoom).toBe(100);
    expect(sheet.rowDimensions.get(0)?.height).toBe(22);
    expect(sheet.columnDimensions.get(0)?.width).toBe(18.5);
    expect(workbook.date1904).toBe(false);
    expect(workbook.styles).toHaveLength(2);
    expect(workbook.styles[1]!.numFmtId).toBe(164);
    expect(workbook.names.map((name) => name.name)).toContain('FixtureRange');
  });

  it('reads shared strings, numbers, booleans, errors, inline strings and cached formula results', () => {
    const { workbook } = openWorkbook(fixture('m1-chart-unknown.xlsx'));
    const sheet = workbook.sheets[0]!;
    const cell = (key: string) => sheet.cells.get(key);
    expect(cell('0:0')?.value).toEqual({ type: 'text', value: 'Name' });
    expect(cell('0:1')?.value).toEqual({ type: 'text', value: 'Value' });
    expect(cell('1:1')?.value).toEqual({ type: 'number', value: 42.5 });
    expect(cell('1:2')?.value).toEqual({ type: 'text', value: 'inline text' });
    expect(cell('2:1')?.formula).toBe('B2*2');
    expect(cell('2:1')?.savedResult).toEqual({ type: 'number', value: 85 });
    expect(cell('3:0')?.value).toEqual({ type: 'boolean', value: true });
    expect(cell('3:1')?.value).toEqual({ type: 'error', value: '#DIV/0!' });
    expect(cell('0:0')?.styleId).toBe(1);
  });

  it('keeps Tier B parts opaque and reports them', () => {
    const opened = openWorkbook(fixture('m1-chart-unknown.xlsx'));
    expect(opened.workbook.opaqueParts.has('xl/charts/chart1.xml')).toBe(true);
    expect(opened.workbook.opaqueParts.has('xl/drawings/drawing1.xml')).toBe(true);
    expect(opened.workbook.opaqueParts.has('xl/drawings/vmlDrawing1.vml')).toBe(true);
    expect(opened.workbook.sourceParts.has('xl/worksheets/sheet1.xml')).toBe(true);
    expect(opened.workbook.sourceParts.has('xl/sharedStrings.xml')).toBe(true);
    expect(opened.compatibility.notes.join(' ')).toMatch(/Charts are preserved/);
    expect(opened.compatibility.externalLinkParts).toEqual([]);
  });

  it('opens the identity fixture and exposes the active sheet', () => {
    const opened = openWorkbook(fixture('identity-minimal.xlsx'));
    expect(opened.initialSheetId).toBe(opened.workbook.sheets[0]!.id);
    expect(opened.workbook.sheets[0]!.cells.size).toBe(3);
  });

  it('detects macro and external-link parts without executing or refreshing them', () => {
    const bytes = buildXlsx(
      '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>',
      {
        extraParts: {
          'xl/vbaProject.bin': 'fake macro bytes',
          'xl/externalLinks/externalLink1.xml':
            '<?xml version="1.0"?><externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>',
        },
      },
    );
    const opened = openWorkbook(bytes);
    expect(opened.compatibility.hasMacros).toBe(true);
    expect(opened.compatibility.externalLinkParts).toContain('xl/externalLinks/externalLink1.xml');
    expect(opened.compatibility.notes.join(' ')).toMatch(/never executed/);
    expect(opened.compatibility.notes.join(' ')).toMatch(/never refreshed/);
    expect(
      new TextDecoder().decode(opened.workbook.opaqueParts.get('xl/vbaProject.bin')!.bytes),
    ).toBe('fake macro bytes');
  });
});

describe('Tier C rejection', () => {
  it('rejects encrypted compound files with the contract message', () => {
    const bytes = new Uint8Array(512);
    bytes.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
    expect(() => openWorkbook(bytes)).toThrow(WorkbookOpenError);
    expect(() => openWorkbook(bytes)).toThrow(/encrypted or password-protected/);
  });

  it('rejects BIFF/ODS extensions before container parsing', () => {
    expect(() => openWorkbook(new Uint8Array(16), 'book.xls')).toThrow(
      /file format is not supported/,
    );
    expect(() => openWorkbook(new Uint8Array(16), 'book.ods')).toThrow(
      /file format is not supported/,
    );
  });

  it('rejects non-ZIP and truncated input without partial import', () => {
    expect(() => openWorkbook(new TextEncoder().encode('not a workbook'))).toThrow(
      /not supported, or the file is damaged/,
    );
  });

  it('rejects Strict OOXML with a clear message', () => {
    const bytes = buildXlsx(
      '<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"><sheetData/></worksheet>',
      { strict: true },
    );
    expect(() => openWorkbook(bytes)).toThrow(/OOXML variant/);
  });

  it('rejects DTD and entity declarations before parsing', () => {
    const bytes = buildXlsx(
      `<?xml version="1.0"?><!DOCTYPE worksheet [<!ENTITY xxe SYSTEM "file:///etc/passwd">]><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>&xxe;</v></c></row></sheetData></worksheet>`,
    );
    expect(() => openWorkbook(bytes)).toThrow(/DTD or entity declaration/);
  });

  it('rejects a zip-bomb ratio before decompressing it', () => {
    const bytes = buildXlsx(
      '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>',
      {
        extraParts: {
          'xl/media/bomb.bin': '0'.repeat(1_000_000),
        },
      },
    );
    expect(() => openWorkbook(bytes)).toThrow(/too large or unusually compressed/);
  });

  it('fails safely when a referenced worksheet part is missing', () => {
    const bytes = buildXlsx(
      '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>',
    );
    const parts = new Map(Object.entries(unzipSync(bytes)));
    parts.delete('xl/worksheets/sheet1.xml');
    const broken = zipSync(Object.fromEntries(parts));
    expect(() => openWorkbook(broken)).toThrow(/worksheet part .* is missing/);
  });
});
