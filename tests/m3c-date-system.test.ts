import { describe, expect, it } from 'vitest';
import { DocumentSession } from '@commands/session';
import {
  FAKE_1900_LEAP_DAY_SERIAL,
  datePartsToSerial,
  isFake1900LeapDay,
  serialToDateParts,
} from '@model/date-system';
import { numberValue } from '@model/value';
import { openWorkbook } from '@ooxml/open';
import { saveWorkbook } from '@ooxml/save';
import { formatValue } from '@view/number-format';
import { buildXlsx, unzipPart } from './helpers/build-xlsx';

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const worksheet = `${XML_HEADER}<worksheet xmlns="${NS}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>`;

describe('M3C date systems', () => {
  it('reproduces the 1900 fake leap day and the pre-March mapping', () => {
    expect(isFake1900LeapDay(FAKE_1900_LEAP_DAY_SERIAL)).toBe(true);
    expect(serialToDateParts(59, false)).toEqual({
      year: 1900,
      month: 2,
      day: 28,
      hours: 0,
      minutes: 0,
      seconds: 0,
    });
    expect(serialToDateParts(60, false)).toEqual({
      year: 1900,
      month: 2,
      day: 29,
      hours: 0,
      minutes: 0,
      seconds: 0,
    });
    expect(serialToDateParts(61, false)).toEqual({
      year: 1900,
      month: 3,
      day: 1,
      hours: 0,
      minutes: 0,
      seconds: 0,
    });
    expect(serialToDateParts(1, false)?.month).toBe(1);
    expect(serialToDateParts(1, false)?.day).toBe(1);
    expect(serialToDateParts(0, false)?.day).toBe(0);
  });

  it('maps serials back to dates in both systems', () => {
    expect(datePartsToSerial(1900, 2, 28, false)).toBe(59);
    expect(datePartsToSerial(1900, 2, 29, false)).toBe(60);
    expect(datePartsToSerial(1900, 3, 1, false)).toBe(61);
    expect(datePartsToSerial(1904, 1, 1, true)).toBe(0);
    expect(datePartsToSerial(1904, 2, 29, true)).toBe(59);
    expect(datePartsToSerial(1904, 3, 1, true)).toBe(60);
  });

  it('formats serials with the workbook date system', () => {
    expect(formatValue(numberValue(59), 'm/d/yy', false)).toBe('2/28/00');
    expect(formatValue(numberValue(60), 'm/d/yy', false)).toBe('2/29/00');
    expect(formatValue(numberValue(61), 'm/d/yy', false)).toBe('3/1/00');
    expect(formatValue(numberValue(60), 'm/d/yy', true)).toBe('3/1/04');
  });

  it('saves and reopens the workbook date-system flag', () => {
    const opened = openWorkbook(buildXlsx(worksheet));
    const session = new DocumentSession(opened.workbook);
    expect(opened.workbook.date1904).toBe(false);

    const change = session.dispatch('workbook.dateSystem', { date1904: true });
    expect(change.changes).toHaveLength(1);
    expect(opened.workbook.date1904).toBe(true);
    expect(session.hasUnsavedChanges).toBe(true);

    const saved = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
    });
    expect(saved.noOp).toBe(false);
    expect(unzipPart(saved.bytes, 'xl/workbook.xml')).toContain('date1904="1"');
    expect(openWorkbook(saved.bytes).workbook.date1904).toBe(true);

    session.dispatch('edit.undo', undefined);
    expect(opened.workbook.date1904).toBe(false);
    session.dispatch('edit.redo', undefined);
    expect(opened.workbook.date1904).toBe(true);
  });
});
