import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import type { EngineResultCell, EngineCellInput, EngineBatchResult } from '@engine/port';
import { HyperFormulaEngine } from '@engine/hyperformula-engine';
import { prepareFormula } from '@engine/formula-text';
import { DocumentSession } from '@commands/session';
import {
  BLANK,
  booleanValue,
  errorValue,
  numberValue,
  textValue,
  type CellValue,
} from '@model/value';
import type { Cell } from '@model/types';
import { openWorkbook } from '@ooxml/open';
import { saveWorkbook } from '@ooxml/save';
import { cellValueMatchesGolden, describeValue, type GoldenCell } from '@testing/compare';
import { buildXlsx, unzipPart } from './helpers/build-xlsx';

interface M3BCase {
  readonly id: string;
  readonly group?: string;
  readonly formula: string;
  readonly expected: GoldenCell;
  readonly dynamic?: boolean;
}

interface M3BFixture {
  readonly schemaVersion: number;
  readonly provenance: {
    readonly kind: string;
    readonly oracle: null;
    readonly note: string;
  };
  readonly data: {
    readonly sheet: string;
    readonly rows: readonly (readonly (string | number | boolean)[])[];
  };
  readonly functions: readonly M3BCase[];
  readonly errors: readonly M3BCase[];
  readonly dates: readonly M3BCase[];
  readonly references: readonly M3BCase[];
  readonly lookups: readonly M3BCase[];
  readonly shims: readonly M3BCase[];
}

const fixture = JSON.parse(
  readFileSync(new URL('../fixtures/expected/m3b-formulas.json', import.meta.url), 'utf8'),
) as M3BFixture;

interface CaseResults {
  readonly result: EngineBatchResult;
  readonly byId: ReadonlyMap<string, EngineResultCell>;
}

const FORMULA_ROW = 30;

function toCellValue(value: string | number | boolean): CellValue {
  if (typeof value === 'number') return numberValue(value);
  if (typeof value === 'boolean') return booleanValue(value);
  return textValue(value);
}

function calculateCases(
  cases: readonly M3BCase[],
  options: { readonly formulaRow?: number; readonly includeData?: boolean } = {},
): CaseResults {
  const formulaRow = options.formulaRow ?? FORMULA_ROW;
  const includeData = options.includeData ?? true;
  const engine = new HyperFormulaEngine();
  engine.start({ date1904: false, sheetNames: [fixture.data.sheet] });
  const inputs: EngineCellInput[] = [];

  if (includeData) {
    for (let row = 0; row < fixture.data.rows.length; row += 1) {
      const values = fixture.data.rows[row] ?? [];
      for (let col = 0; col < values.length; col += 1) {
        inputs.push({
          sheet: 0,
          row,
          col,
          formula: null,
          value: toCellValue(values[col] ?? ''),
        });
      }
    }
  }

  for (let col = 0; col < cases.length; col += 1) {
    inputs.push({
      sheet: 0,
      row: formulaRow,
      col,
      formula: cases[col]!.formula,
      value: BLANK,
    });
  }

  engine.load(inputs);
  const result = engine.finishLoad();
  const byCol = new Map(
    result.cells.filter((cell) => cell.row === formulaRow).map((cell) => [cell.col, cell]),
  );
  const byId = new Map<string, EngineResultCell>();
  for (let col = 0; col < cases.length; col += 1) {
    const cell = byCol.get(col);
    if (cell === undefined) throw new Error(`No engine result for ${cases[col]!.id}`);
    byId.set(cases[col]!.id, cell);
  }
  engine.dispose();
  return { result, byId };
}

const functionResults = calculateCases(fixture.functions);
const errorResults = calculateCases(fixture.errors);
const dateResults = calculateCases(fixture.dates);
const referenceResults = calculateCases(fixture.references);
const lookupResults = calculateCases(fixture.lookups);
const shimResults = calculateCases(fixture.shims);
const circularResults = calculateCases(
  [
    {
      id: 'dead-branch',
      formula: 'IF(FALSE(),A1,0)',
      expected: { kind: 'error', value: '#CIRC!' },
    },
  ],
  { formulaRow: 0, includeData: false },
);

function expectCase(results: CaseResults, golden: M3BCase): void {
  const cell = results.byId.get(golden.id);
  expect(cell, `${golden.id} produced no result`).toBeDefined();
  if (cell === undefined) return;
  if (golden.dynamic === true) {
    expect(cell.value.type).toBe('number');
    return;
  }
  const matches = cellValueMatchesGolden(cell.value, golden.expected);
  expect(matches, `${golden.id} was ${describeValue(cell.value)}`).toBe(true);
}

function collectErrorValues(results: readonly CaseResults[]): string[] {
  const values: string[] = [];
  for (const group of results) {
    for (const cell of group.result.cells) {
      if (cell.value.type === 'error') values.push(cell.value.value);
    }
  }
  return values;
}

describe('M3B fixture provenance', () => {
  it('declares hand-authored provenance, not Excel-oracle provenance', () => {
    expect(fixture.schemaVersion).toBe(1);
    expect(fixture.provenance.kind).toBe('hand-authored');
    expect(fixture.provenance.oracle).toBeNull();
    expect(fixture.provenance.note).toMatch(/not Excel 365 oracle goldens/);
  });

  it('covers exactly the 52 published functions and skips none', () => {
    expect(fixture.functions).toHaveLength(52);
    expect(new Set(fixture.functions.map((golden) => golden.id)).size).toBe(52);
    expect(functionResults.result.cells).toHaveLength(52);
    expect(functionResults.result.skipped).toHaveLength(0);
  });
});

describe('M3B 52-function golden subset', () => {
  for (const golden of fixture.functions) {
    it(`${golden.group ?? 'Function'} ${golden.id} matches its golden`, () => {
      expectCase(functionResults, golden);
    });
  }
});

describe('M3B error, coercion and operator table', () => {
  it('produces every expected error, coercion and operator result', () => {
    expect(errorResults.result.skipped).toHaveLength(0);
    for (const golden of fixture.errors) expectCase(errorResults, golden);
  });

  it('never produces #ERROR!, #SPILL!, #CALC! or #LIC!', () => {
    const banned = new Set(['#ERROR!', '#SPILL!', '#CALC!', '#LIC!']);
    const values = collectErrorValues([
      functionResults,
      errorResults,
      dateResults,
      referenceResults,
      lookupResults,
      shimResults,
      circularResults,
    ]);
    for (const value of values) expect(banned.has(value)).toBe(false);
  });
});

describe('M3B date, reference, lookup and circular-reference suites', () => {
  it('passes the date-function cases', () => {
    expect(dateResults.result.skipped).toHaveLength(0);
    for (const golden of fixture.dates) expectCase(dateResults, golden);
  });

  it('passes the reference cases', () => {
    expect(referenceResults.result.skipped).toHaveLength(0);
    for (const golden of fixture.references) expectCase(referenceResults, golden);
  });

  it('passes the lookup cases', () => {
    expect(lookupResults.result.skipped).toHaveLength(0);
    for (const golden of fixture.lookups) expectCase(lookupResults, golden);
  });

  it('maps a statically dead self-reference to #CIRC! and reports the cycle', () => {
    expect(circularResults.result.skipped).toHaveLength(0);
    const cell = circularResults.byId.get('dead-branch');
    expect(cell?.status).toBe('circular');
    expect(cell?.value).toEqual(errorValue('#CIRC!'));
    expect(circularResults.result.circular).toHaveLength(1);
    expect(circularResults.result.cells.some((result) => result.value.type === 'error')).toBe(true);
  });
});

describe('M3B shim cases', () => {
  it('does not skip any shim formula', () => {
    expect(shimResults.result.skipped).toHaveLength(0);
  });

  for (const golden of fixture.shims) {
    it(`${golden.id} matches its shim golden`, () => {
      expectCase(shimResults, golden);
    });
  }

  it('reports a fractional CHOOSE selector through the shim and truncates it', () => {
    const cell = shimResults.byId.get('choose-fraction-low');
    expect(cell?.value).toEqual(textValue('a'));
    const high = shimResults.byId.get('choose-fraction-high');
    expect(high?.value).toEqual(textValue('b'));
  });

  it('matches numeric and numeric-looking text criteria for COUNTIF and SUMIF', () => {
    expect(shimResults.byId.get('criteria-numeric-text')?.value).toEqual(numberValue(2));
    expect(shimResults.byId.get('criteria-numeric-text-not')?.value).toEqual(numberValue(2));
    expect(shimResults.byId.get('sumif-numeric-text')?.value).toEqual(numberValue(5));
    expect(shimResults.byId.get('sumif-numeric-text-not')?.value).toEqual(numberValue(17));
  });
});

describe('M3B authored-formula pre-screen', () => {
  it('keeps malformed and v2 constructs out of the engine', () => {
    for (const formula of [
      '1+',
      '1+*2',
      'A1#',
      'A1#B1',
      'SUM(1+)',
      'SUM((A1:A2,B1:B2))',
      'A1 B1',
      '@A1',
      '_xlfn.CONCAT("a")',
    ]) {
      const prepared = prepareFormula(formula);
      expect(prepared.text, formula).toBeNull();
      expect(prepared.skip, formula).toBe('unparseable');
    }
  });

  it('still sends supported formulas and valid error literals', () => {
    expect(prepareFormula('SUM(A1:A2)').text).toBe('SUM(A1:A2)');
    expect(prepareFormula('#N/A').text).toBe('#N/A');
    expect(prepareFormula('INT(-8.9)').text).toBe('CALCPLUS_INT(-8.9)');
  });

  it('reports unsupported functions separately from unparseable text', () => {
    const prepared = prepareFormula('COLUMNS(A1:A2)');
    expect(prepared.text).toBeNull();
    expect(prepared.skip).toBe('unsupported-function');
    expect(prepared.unsupportedFunctions).toContain('COLUMNS');
  });
});

describe('M3A incremental, volatile and circular behaviour', () => {
  it('recalculates dependants incrementally and includes volatile cells in the batch', () => {
    const engine = new HyperFormulaEngine();
    engine.start({ date1904: false, sheetNames: ['Sheet1'] });
    engine.load([
      { sheet: 0, row: 0, col: 0, formula: null, value: numberValue(2) },
      { sheet: 0, row: 0, col: 1, formula: 'A1*2', value: BLANK },
      { sheet: 0, row: 0, col: 2, formula: 'NOW()', value: BLANK },
    ]);
    engine.finishLoad();
    const batch = engine.apply([
      { sheet: 0, row: 0, col: 0, formula: null, value: numberValue(5) },
    ]);
    const byCol = new Map(batch.cells.map((cell) => [cell.col, cell]));
    expect(byCol.get(1)?.value).toEqual(numberValue(10));
    expect(byCol.get(2)?.value.type).toBe('number');
    engine.dispose();
  });
});

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const ws = (rows: string): string =>
  `${XML_HEADER}<worksheet xmlns="${NS}"><sheetData>${rows}</sheetData></worksheet>`;

function worksheetWithFormulaCache(): Uint8Array {
  return buildXlsx(
    ws(
      '<row r="1">' +
        '<c r="A1"><v>2</v></c>' +
        '<c r="B1"><f>A1*2</f><v>4</v></c>' +
        '<c r="C1"><f>COLUMNS(A1:B1)</f><v>7</v></c>' +
        '<c r="D1"><f>C1+1</f><v>8</v></c>' +
        '<c r="E1"><f>1+</f><v>5</v></c>' +
        '<c r="F1"><f>A1#</f><v>9</v></c>' +
        '<c r="G1"><f>IF(FALSE(),G1,0)</f><v>0</v></c>' +
        '</row>',
    ),
  );
}

function cellFor(sheet: { readonly cells: Map<string, Cell> }, key: string): Cell {
  const cell = sheet.cells.get(key);
  if (cell === undefined) throw new Error(`Missing cell ${key}`);
  return cell;
}

describe('M3B authored text, cached results and stale state', () => {
  it('keeps authored text and cached results while marking dependants stale', () => {
    const opened = openWorkbook(worksheetWithFormulaCache());
    const session = new DocumentSession(opened.workbook);
    const sheet = opened.workbook.sheets[0]!;
    expect(cellFor(sheet, '0:1').formula).toBe('A1*2');
    expect(cellFor(sheet, '0:2').formula).toBe('COLUMNS(A1:B1)');
    expect(cellFor(sheet, '0:3').formula).toBe('C1+1');
    expect(cellFor(sheet, '0:4').formula).toBe('1+');
    expect(cellFor(sheet, '0:5').formula).toBe('A1#');
    expect(cellFor(sheet, '0:6').formula).toBe('IF(FALSE(),G1,0)');

    session.dispatch('cell.set', {
      sheetId: sheet.id,
      row: 0,
      col: 0,
      value: numberValue(5),
      formula: null,
    });
    session.applyCalculationResults(
      [
        {
          sheetId: sheet.id,
          row: 0,
          col: 1,
          value: numberValue(10),
          status: 'value',
          stale: false,
        },
        { sheetId: sheet.id, row: 0, col: 2, value: numberValue(7), status: 'value', stale: true },
        { sheetId: sheet.id, row: 0, col: 3, value: numberValue(8), status: 'value', stale: true },
        { sheetId: sheet.id, row: 0, col: 4, value: numberValue(5), status: 'value', stale: true },
        { sheetId: sheet.id, row: 0, col: 5, value: numberValue(9), status: 'value', stale: true },
        {
          sheetId: sheet.id,
          row: 0,
          col: 6,
          value: errorValue('#CIRC!'),
          status: 'circular',
          stale: true,
        },
      ],
      { markDirty: true },
    );
    session.dispatch('cell.set', {
      sheetId: sheet.id,
      row: 0,
      col: 7,
      value: BLANK,
      formula: 'SOMETHING(1)',
    });
    session.applyCalculationResults(
      [
        {
          sheetId: sheet.id,
          row: 0,
          col: 7,
          value: errorValue('#NAME?'),
          status: 'value',
          stale: true,
        },
      ],
      { markDirty: true },
    );

    expect(cellFor(sheet, '0:1').savedResult).toEqual(numberValue(10));
    expect(cellFor(sheet, '0:1').stale).toBe(false);
    expect(cellFor(sheet, '0:2').savedResult).toEqual(numberValue(7));
    expect(cellFor(sheet, '0:2').stale).toBe(true);
    expect(cellFor(sheet, '0:3').value).toEqual(numberValue(8));
    expect(cellFor(sheet, '0:3').savedResult).toEqual(numberValue(8));
    expect(cellFor(sheet, '0:3').stale).toBe(true);
    expect(cellFor(sheet, '0:4').savedResult).toEqual(numberValue(5));
    expect(cellFor(sheet, '0:4').stale).toBe(true);
    expect(cellFor(sheet, '0:5').savedResult).toEqual(numberValue(9));
    expect(cellFor(sheet, '0:5').stale).toBe(true);
    expect(cellFor(sheet, '0:6').value).toEqual(errorValue('#CIRC!'));
    expect(cellFor(sheet, '0:6').savedResult).toEqual(numberValue(0));
    expect(cellFor(sheet, '0:6').stale).toBe(true);
    expect(cellFor(sheet, '0:7').savedResult).toBeNull();
    expect(cellFor(sheet, '0:7').value).toEqual(errorValue('#NAME?'));
    expect(cellFor(sheet, '0:7').stale).toBe(true);
    expect(session.staleCellCount()).toBe(6);

    const saved = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
    });
    const xml = unzipPart(saved.bytes, 'xl/worksheets/sheet1.xml');
    expect(xml).toContain('<c r="B1"><f>A1*2</f><v>10</v></c>');
    expect(xml).toContain('<c r="C1"><f>COLUMNS(A1:B1)</f><v>7</v></c>');
    expect(xml).toContain('<c r="D1"><f>C1+1</f><v>8</v></c>');
    expect(xml).toContain('<c r="E1"><f>1+</f><v>5</v></c>');
    expect(xml).toContain('<c r="F1"><f>A1#</f><v>9</v></c>');
    expect(xml).toContain('<c r="G1"><f>IF(FALSE(),G1,0)</f><v>0</v></c>');
    expect(xml).toContain('<c r="H1"><f>SOMETHING(1)</f></c>');
    expect(unzipPart(saved.bytes, 'xl/workbook.xml')).toContain('fullCalcOnLoad="1"');
    for (const banned of ['#CIRC!', '#NAME?', '#ERROR!', '#SPILL!', '#CALC!', '#LIC!']) {
      expect(xml).not.toContain(banned);
    }
  });
});
