import { describe, expect, it } from 'vitest';
import { DocumentSession } from '@commands/session';
import {
  FUNCTION_ARITY,
  PUBLISHED_FUNCTIONS,
  type PublishedFunction,
} from '@engine/function-registry';
import { prepareFormula } from '@engine/formula-text';
import { openWorkbook } from '@ooxml/open';
import { saveWorkbook } from '@ooxml/save';
import { BLANK } from '@model/value';
import type { Cell } from '@model/types';
import { buildXlsx, unzipPart } from './helpers/build-xlsx';

/**
 * Argument-count screening (formula semantics §2 and §9).
 *
 * Excel refuses a formula that gives a function the wrong number of arguments
 * ("There's a problem with this formula") and keeps the cell's previous value.
 * calcplus used to accept `SUM()`, evaluate it to `#N/A` and write it out;
 * Excel then stripped the formula from the saved file. These tests pin the
 * accepted ranges, the argument counting rules and the entry refusal.
 *
 * The expected ranges in `FUNCTION_ARITY` were verified against Excel 365's
 * parser: one count below the minimum, both boundaries and one above the
 * maximum for every published function.
 */

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const ws = (rows: string): string =>
  `${XML_HEADER}<worksheet xmlns="${NS}"><sheetData>${rows}</sheetData></worksheet>`;

function call(name: string, argCount: number): string {
  return argCount === 0
    ? `${name}()`
    : `${name}(${Array.from({ length: argCount }, () => '1').join(',')})`;
}

function cellFor(sheet: { readonly cells: Map<string, Cell> }, key: string): Cell {
  const cell = sheet.cells.get(key);
  if (cell === undefined) throw new Error(`Missing cell ${key}`);
  return cell;
}

describe('every published function accepts its minimum and maximum argument count', () => {
  for (const name of PUBLISHED_FUNCTIONS) {
    const arity = FUNCTION_ARITY[name as PublishedFunction];
    it(`${name} accepts ${arity.min} to ${arity.max} arguments`, () => {
      const minimum = prepareFormula(call(name, arity.min));
      expect(minimum.skip, call(name, arity.min)).toBeNull();
      expect(minimum.text, call(name, arity.min)).not.toBeNull();

      const maximum = prepareFormula(call(name, arity.max));
      expect(maximum.skip, call(name, arity.max)).toBeNull();
      expect(maximum.text, call(name, arity.max)).not.toBeNull();

      const above = prepareFormula(call(name, arity.max + 1));
      expect(above.skip, call(name, arity.max + 1)).toBe('invalid-arity');
      expect(above.text).toBeNull();

      if (arity.min > 0) {
        const below = prepareFormula(call(name, arity.min - 1));
        expect(below.skip, call(name, arity.min - 1)).toBe('invalid-arity');
        expect(below.text).toBeNull();
      }
    });
  }
});

describe('zero and empty arguments are counted as Excel counts them', () => {
  it('treats FN() as zero arguments and FN(,) as two empty ones', () => {
    // Excel refuses `=SUM()` and `=SUM( )` but accepts `=SUM(,)` and `=SUM(1,)`.
    expect(prepareFormula('SUM()').skip).toBe('invalid-arity');
    expect(prepareFormula('SUM()').arityErrors[0]).toEqual({
      name: 'SUM',
      argCount: 0,
      min: 1,
      max: 255,
    });
    expect(prepareFormula('SUM( )').skip).toBe('invalid-arity');
    expect(prepareFormula('SUM( )').arityErrors[0]?.argCount).toBe(0);
    expect(prepareFormula('SUM(,)').skip).toBeNull();
    expect(prepareFormula('SUM(,)').arityErrors).toEqual([]);
    expect(prepareFormula('SUM(1,)').skip).toBeNull();
  });

  it('counts empty arguments toward the total', () => {
    // `TODAY(,)` is two empty arguments, so it is refused like `TODAY(1,1)`.
    expect(prepareFormula('TODAY(,)').skip).toBe('invalid-arity');
    expect(prepareFormula('TODAY(,)').arityErrors[0]?.argCount).toBe(2);
    expect(prepareFormula('NOT(,)').skip).toBe('invalid-arity');
    expect(prepareFormula('LEFT(,)').skip).toBeNull();
  });
});

describe('only top-level commas count toward a call argument count', () => {
  it('ignores commas inside nested calls', () => {
    expect(prepareFormula('IF(A1>0, SUM(B1:B2), MAX(C1, C2))').skip).toBeNull();
    // One comma for IF; the two inside SUM do not make five arguments.
    expect(prepareFormula('IF(SUM(1,2), 3, 4)').skip).toBeNull();
    expect(prepareFormula('MAX(SUM(1, 2, 3), MIN(4, 5))').skip).toBeNull();
    expect(prepareFormula('CHOOSE(1, SUM(1,2), MAX(3,4), MIN(5,6))').skip).toBeNull();
  });

  it('refuses the outer call when its own count is wrong and surfaces inner violations', () => {
    const outer = prepareFormula('IF(A1, SUM(B1:B2), MAX(C1,C2), D1)');
    expect(outer.skip).toBe('invalid-arity');
    expect(outer.arityErrors[0]).toMatchObject({ name: 'IF', argCount: 4 });

    // SUM is fine at two arguments; ROUND is not, and the formula is refused.
    const inner = prepareFormula('SUM(A1, ROUND(1))');
    expect(inner.skip).toBe('invalid-arity');
    expect(inner.arityErrors[0]).toMatchObject({ name: 'ROUND', argCount: 1 });

    const both = prepareFormula('IF(ABS(1,2), TODAY(1), 1)');
    expect(both.skip).toBe('invalid-arity');
    expect(both.arityErrors.map((error) => error.name)).toEqual(['ABS', 'TODAY']);
  });

  it('ignores commas inside array constants', () => {
    // Excel accepts `=SUM({1,2},3)` and `=IF({1,2},3,4)`.
    expect(prepareFormula('SUM({1,2},3)').skip).toBeNull();
    expect(prepareFormula('IF({1,2}, 3, 4)').skip).toBeNull();
  });
});

describe('names outside the published set carry no arity', () => {
  it('classifies a zero-argument unknown call as an unsupported function', () => {
    const prepared = prepareFormula('_xlfn.XLOOKUP()');
    expect(prepared.skip).toBe('unsupported-function');
    expect(prepared.arityErrors).toEqual([]);
  });

  it('does not arity-check TRUE() and FALSE() calls', () => {
    expect(prepareFormula('TRUE()').skip).toBeNull();
    expect(prepareFormula('FALSE()').skip).toBeNull();
    expect(prepareFormula('IF(TRUE(), 1, 2)').skip).toBeNull();
  });
});

describe('entry refuses a wrong argument count and keeps the cell', () => {
  const sheet = ws('<row r="1"><c r="A1"><v>2</v></c><c r="B1"><f>A1*2</f><v>4</v></c></row>');

  it('leaves a formula cell untouched and explains the problem', () => {
    const opened = openWorkbook(buildXlsx(sheet));
    const session = new DocumentSession(opened.workbook);
    const model = opened.workbook.sheets[0]!;
    const result = session.dispatch('cell.set', {
      sheetId: model.id,
      row: 0,
      col: 1,
      value: BLANK,
      formula: 'SUM()',
    });
    expect(result.changes).toEqual([]);
    expect(result.ui).toMatchObject({ type: 'toast' });
    expect(result.ui?.message).toContain('SUM takes 1 to 255 arguments');
    expect(result.ui?.message).toContain('The cell was not changed.');

    const cell = cellFor(model, '0:1');
    expect(cell.formula).toBe('A1*2');
    expect(cell.savedResult).toEqual({ type: 'number', value: 4 });
    expect(session.hasUnsavedChanges).toBe(false);
  });

  it('leaves a value cell untouched', () => {
    const opened = openWorkbook(buildXlsx(sheet));
    const session = new DocumentSession(opened.workbook);
    const model = opened.workbook.sheets[0]!;
    const result = session.dispatch('cell.set', {
      sheetId: model.id,
      row: 0,
      col: 0,
      value: BLANK,
      formula: 'TODAY(1)',
    });
    expect(result.changes).toEqual([]);
    expect(result.ui?.message).toContain('TODAY takes no arguments');
    expect(cellFor(model, '0:0').value).toEqual({ type: 'number', value: 2 });
  });

  it('refuses every invalid form and accepts the valid ones', () => {
    const opened = openWorkbook(buildXlsx(sheet));
    const session = new DocumentSession(opened.workbook);
    const model = opened.workbook.sheets[0]!;
    const refused = [
      'SUM()',
      'SUM( )',
      'IF(A1)',
      'ROUND(1)',
      'INDEX(A1:A5)',
      'INDEX(A1:B2,1,1,1)',
      'MATCH(1,A1:A2,0,1)',
      'TODAY(1,1)',
    ];
    for (const formula of refused) {
      const result = session.dispatch('cell.set', {
        sheetId: model.id,
        row: 0,
        col: 0,
        value: BLANK,
        formula,
      });
      expect(result.changes, formula).toEqual([]);
      expect(result.ui?.message, formula).toContain('The cell was not changed.');
    }
    expect(cellFor(model, '0:0').value).toEqual({ type: 'number', value: 2 });

    const accepted = session.dispatch('cell.set', {
      sheetId: model.id,
      row: 0,
      col: 0,
      value: BLANK,
      formula: 'SUM(1,2)',
    });
    expect(accepted.changes).toHaveLength(1);
    expect(cellFor(model, '0:0').formula).toBe('SUM(1,2)');
  });
});

describe('a saved workbook holds no formula outside the accepted arities', () => {
  it('writes every accepted formula and none of the refused ones', () => {
    const opened = openWorkbook(buildXlsx(ws('<row r="1"><c r="A1"><v>2</v></c></row>')));
    const session = new DocumentSession(opened.workbook);
    const model = opened.workbook.sheets[0]!;

    const accepted = [
      'SUM(A1:A2)',
      'SUM(,)',
      'SUM({1,2},3)',
      'IF(A1>0, SUM(A1:A2), MAX(A1, A2))',
      'TODAY()',
      'ROW()',
      'COLUMN(A1)',
      'INDEX(A1:A2, 1)',
      'CHOOSE(1, 2, 3)',
      'LEFT("ab", 1)',
      'MID("abc", 1, 2)',
      'VLOOKUP(A1, A1:B2, 2, FALSE)',
    ];
    const refused = [
      'SUM()',
      'SUM( )',
      'IF(A1)',
      'IF(1,2,3,4)',
      'TODAY(1)',
      'INDEX(A1:A5)',
      'INDEX(A1:B2,1,1,1)',
      'MATCH(1,A1:A2,0,1)',
      'ROUND(1)',
    ];

    let row = 1;
    for (const formula of accepted) {
      const result = session.dispatch('cell.set', {
        sheetId: model.id,
        row,
        col: 0,
        value: BLANK,
        formula,
      });
      expect(result.changes, formula).toHaveLength(1);
      row += 1;
    }
    for (const formula of refused) {
      const result = session.dispatch('cell.set', {
        sheetId: model.id,
        row,
        col: 1,
        value: BLANK,
        formula,
      });
      expect(result.changes, formula).toEqual([]);
      row += 1;
    }

    const saved = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
    });
    const xml = unzipPart(saved.bytes, 'xl/worksheets/sheet1.xml');
    const written = [...xml.matchAll(/<f(?:\s[^>]*)?>([\s\S]*?)<\/f>/g)].map((match) =>
      decodeXml(match[1] ?? ''),
    );

    for (const formula of refused) {
      expect(written, formula).not.toContain(formula);
    }
    expect(written).toEqual(accepted);
    // The property Excel was enforcing for us: every written formula is within
    // the arities calcplus accepts.
    for (const formula of written) {
      expect(prepareFormula(formula).skip, formula).toBeNull();
    }
  });
});

function decodeXml(text: string): string {
  return text
    .replaceAll('&lt;', '<')
    .replaceAll('&gt;', '>')
    .replaceAll('&quot;', '"')
    .replaceAll('&apos;', "'")
    .replaceAll('&amp;', '&');
}
