import { describe, expect, it } from 'vitest';
import fc from 'fast-check';
import type { GridRange } from '@model/types';
import type {
  CfCalendarDate,
  CfCellOperator,
  CfCellSource,
  CfDateOperator,
  CfColorScaleOverlay,
  CfDataBarOverlay,
  CfDifferentialFormat,
  CfIconSetId,
  CfIconSetOverlay,
  CfOperand,
  CfTextOperator,
  CfThreshold,
  CfTimePeriod,
  ConditionalFormatRule,
} from '@model/conditional-format';
import {
  CF_ICON_SETS,
  createConditionalFormatIndex,
  orderRulesByPriority,
} from '@model/conditional-format';
import type { CellValue } from '@model/value';
import {
  BLANK,
  booleanValue,
  errorValue,
  formatCellValue,
  numberValue,
  textValue,
} from '@model/value';
import { datePartsToSerial } from '@model/date-system';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

const SHEET = 'sheet-1';
const TODAY: CfCalendarDate = { year: 2024, month: 3, day: 15 };

function range(startRow: number, startCol: number, endRow: number, endCol: number): GridRange {
  return { startRow, startCol, endRow, endCol };
}

const COLUMN_A: GridRange = range(0, 0, 99, 0);

interface RuleBaseFields {
  readonly id: string;
  readonly sheetId: string;
  readonly ranges: readonly GridRange[];
  readonly priority?: number;
  readonly stopIfTrue: boolean;
  readonly raw: unknown;
}

function base(
  id: string,
  ranges: GridRange | readonly GridRange[],
  priority?: number,
  stopIfTrue = false,
): RuleBaseFields {
  return {
    id,
    sheetId: SHEET,
    ranges: Array.isArray(ranges) ? ranges : [ranges],
    priority,
    stopIfTrue,
    raw: { id },
  };
}

function cellIsRule(
  id: string,
  operator: CfCellOperator,
  value1: CfOperand,
  value2?: CfOperand,
  options: {
    ranges?: GridRange | readonly GridRange[];
    priority?: number;
    stopIfTrue?: boolean;
    dxfId?: number;
  } = {},
): ConditionalFormatRule {
  return {
    ...base(id, options.ranges ?? COLUMN_A, options.priority, options.stopIfTrue),
    ...(options.dxfId === undefined ? {} : { dxfId: options.dxfId }),
    kind: 'cellIs',
    operator,
    value1,
    ...(value2 === undefined ? {} : { value2 }),
  } as ConditionalFormatRule;
}

function numberOperand(value: number): CfOperand {
  return { type: 'number', value };
}

function textOperand(value: string): CfOperand {
  return { type: 'text', value };
}

function booleanOperand(value: boolean): CfOperand {
  return { type: 'boolean', value };
}

function perCellIndex(
  rules: readonly ConditionalFormatRule[],
  options: { today?: CfCalendarDate; date1904?: boolean; cells?: CfCellSource } = {},
) {
  return createConditionalFormatIndex({
    sheetId: SHEET,
    rules,
    date1904: options.date1904 ?? false,
    today: options.today ?? TODAY,
    ...(options.cells === undefined ? {} : { cells: options.cells }),
  });
}

function sourceFrom(
  values: Record<string, CellValue>,
  displays: Record<string, string> = {},
): CfCellSource {
  return {
    valueAt: (row, col) => values[`${row}:${col}`] ?? BLANK,
    displayTextAt: (row, col) => {
      const key = `${row}:${col}`;
      return displays[key] ?? formatCellValue(values[key] ?? BLANK);
    },
  };
}

function columnSource(values: readonly CellValue[], col = 0): CfCellSource {
  const map: Record<string, CellValue> = {};
  values.forEach((value, row) => {
    map[`${row}:${col}`] = value;
  });
  return sourceFrom(map);
}

function serial(year: number, month: number, day: number, date1904 = false): number {
  return datePartsToSerial(year, month, day, date1904);
}

function datePeriodRule(id: string, period: CfTimePeriod): ConditionalFormatRule {
  return {
    ...base(id, COLUMN_A, 1),
    kind: 'timePeriod',
    mode: 'period',
    period,
  } as ConditionalFormatRule;
}

function dateComparisonRule(
  id: string,
  operator: 'equal' | 'notEqual' | 'before' | 'after' | 'between' | 'notBetween',
  value1: number,
  value2?: number,
): ConditionalFormatRule {
  return {
    ...base(id, COLUMN_A, 1),
    kind: 'timePeriod',
    mode: 'comparison',
    operator,
    value1,
    ...(value2 === undefined ? {} : { value2 }),
  } as ConditionalFormatRule;
}

function textRule(
  id: string,
  operator: CfTextOperator,
  text: string,
  ranges: GridRange = COLUMN_A,
): ConditionalFormatRule {
  return {
    ...base(id, ranges, 1),
    kind: 'text',
    operator,
    text,
  } as ConditionalFormatRule;
}

function threshold(rule: CfThreshold): CfThreshold {
  return rule;
}

const RED = '#FF0000';
const BLUE = '#0000FF';
const GREEN = '#00FF00';
const YELLOW = '#FFFF00';
const BLACK = '#000000';
const WHITE = '#FFFFFF';

// ---------------------------------------------------------------------------
// CF1 — cell values
// ---------------------------------------------------------------------------

describe('CF1 cf-cell-value', () => {
  const cases: Array<{
    operator: CfCellOperator;
    value1: number;
    value2?: number;
    matching: number[];
    failing: number[];
  }> = [
    { operator: 'lessThan', value1: 20, matching: [10, 19], failing: [20, 30] },
    { operator: 'lessThanOrEqual', value1: 20, matching: [20, 19], failing: [21] },
    { operator: 'equal', value1: 20, matching: [20], failing: [19, 21] },
    { operator: 'notEqual', value1: 20, matching: [19, 21], failing: [20] },
    { operator: 'greaterThanOrEqual', value1: 20, matching: [20, 21], failing: [19] },
    { operator: 'greaterThan', value1: 20, matching: [21, 30], failing: [20] },
    {
      operator: 'between',
      value1: 10,
      value2: 20,
      matching: [10, 15, 20],
      failing: [9, 21],
    },
    {
      operator: 'notBetween',
      value1: 10,
      value2: 20,
      matching: [9, 21],
      failing: [10, 15, 20],
    },
  ];

  for (const fixture of cases) {
    it(`evaluates ${fixture.operator}`, () => {
      const index = perCellIndex([
        cellIsRule(
          'r1',
          fixture.operator,
          numberOperand(fixture.value1),
          fixture.value2 === undefined ? undefined : numberOperand(fixture.value2),
        ),
      ]);
      for (const value of fixture.matching) {
        expect(index.evaluateCell(0, 0, { value: numberValue(value) }).ruleIds, `${value}`).toEqual(
          ['r1'],
        );
      }
      for (const value of fixture.failing) {
        expect(index.evaluateCell(0, 0, { value: numberValue(value) }).ruleIds, `${value}`).toEqual(
          [],
        );
      }
    });
  }

  it('coerces a blank to 0 in numeric comparisons', () => {
    const lessThan = perCellIndex([cellIsRule('lt', 'lessThan', numberOperand(20))]);
    expect(lessThan.evaluateCell(0, 0).ruleIds).toEqual(['lt']);

    const greaterThanMinusOne = perCellIndex([cellIsRule('gt', 'greaterThan', numberOperand(-1))]);
    expect(greaterThanMinusOne.evaluateCell(0, 0).ruleIds).toEqual(['gt']);

    const equalZero = perCellIndex([cellIsRule('eq', 'equal', numberOperand(0))]);
    expect(equalZero.evaluateCell(0, 0).ruleIds).toEqual(['eq']);

    const between = perCellIndex([
      cellIsRule('between', 'between', numberOperand(-5), numberOperand(5)),
    ]);
    expect(between.evaluateCell(0, 0).ruleIds).toEqual(['between']);

    const greaterThanZero = perCellIndex([cellIsRule('gt0', 'greaterThan', numberOperand(0))]);
    expect(greaterThanZero.evaluateCell(0, 0).ruleIds).toEqual([]);
  });

  it('compares text case-insensitively and orders booleans FALSE < TRUE', () => {
    const equalText = perCellIndex([cellIsRule('t', 'equal', textOperand('abc'))]);
    expect(equalText.evaluateCell(0, 0, { value: textValue('ABC') }).ruleIds).toEqual(['t']);
    expect(equalText.evaluateCell(0, 0, { value: textValue('abd') }).ruleIds).toEqual([]);

    const falseBeforeTrue = perCellIndex([
      cellIsRule('gtFalse', 'greaterThan', booleanOperand(false)),
    ]);
    expect(falseBeforeTrue.evaluateCell(0, 0, { value: booleanValue(true) }).ruleIds).toEqual([
      'gtFalse',
    ]);
    expect(falseBeforeTrue.evaluateCell(0, 0, { value: booleanValue(false) }).ruleIds).toEqual([]);

    const lessThanTrue = perCellIndex([cellIsRule('ltTrue', 'lessThan', booleanOperand(true))]);
    expect(lessThanTrue.evaluateCell(0, 0, { value: booleanValue(false) }).ruleIds).toEqual([
      'ltTrue',
    ]);
    expect(lessThanTrue.evaluateCell(0, 0, { value: booleanValue(true) }).ruleIds).toEqual([]);
  });

  it('never matches an error, even for not-equal and not-between', () => {
    const notEqual = perCellIndex([cellIsRule('ne', 'notEqual', numberOperand(5))]);
    const notBetween = perCellIndex([
      cellIsRule('nb', 'notBetween', numberOperand(1), numberOperand(10)),
    ]);
    const lessThan = perCellIndex([cellIsRule('lt', 'lessThan', numberOperand(20))]);
    const error = errorValue('#DIV/0!');
    expect(notEqual.evaluateCell(0, 0, { value: error }).ruleIds).toEqual([]);
    expect(notBetween.evaluateCell(0, 0, { value: error }).ruleIds).toEqual([]);
    expect(lessThan.evaluateCell(0, 0, { value: error }).ruleIds).toEqual([]);
  });

  it('compares a blank against a text operand as the empty string', () => {
    const equalEmpty = perCellIndex([cellIsRule('empty', 'equal', textOperand(''))]);
    expect(equalEmpty.evaluateCell(0, 0).ruleIds).toEqual(['empty']);
    const equalText = perCellIndex([cellIsRule('text', 'equal', textOperand('abc'))]);
    expect(equalText.evaluateCell(0, 0).ruleIds).toEqual([]);
  });
});

// ---------------------------------------------------------------------------
// CF2 — text
// ---------------------------------------------------------------------------

describe('CF2 cf-text', () => {
  function matches(operator: CfTextOperator, pattern: string, value: string): string[] {
    const index = perCellIndex([textRule('r1', operator, pattern)]);
    return [...index.evaluateCell(0, 0, { value: textValue(value) }).ruleIds];
  }

  it('covers contains, not contains, begins with and ends with', () => {
    expect(matches('contains', 'world', 'Hello World')).toEqual(['r1']);
    expect(matches('contains', 'world', 'Hello')).toEqual([]);
    expect(matches('notContains', 'world', 'Hello')).toEqual(['r1']);
    expect(matches('notContains', 'world', 'Hello World')).toEqual([]);
    expect(matches('beginsWith', 'hell', 'Hello World')).toEqual(['r1']);
    expect(matches('beginsWith', 'world', 'Hello World')).toEqual([]);
    expect(matches('endsWith', 'world', 'Hello World')).toEqual(['r1']);
    expect(matches('endsWith', 'hello', 'Hello World')).toEqual([]);
  });

  it('is case-insensitive', () => {
    expect(matches('contains', 'WORLD', 'hello world')).toEqual(['r1']);
    expect(matches('beginsWith', 'HEL', 'hello world')).toEqual(['r1']);
    expect(matches('endsWith', 'ORLD', 'hello world')).toEqual(['r1']);
  });

  it('supports *, ? and ~ wildcards', () => {
    expect(matches('contains', 'h*o', 'hello')).toEqual(['r1']);
    expect(matches('contains', 'h?llo', 'hello')).toEqual(['r1']);
    expect(matches('contains', 'h?llo', 'hllo')).toEqual([]);
    expect(matches('contains', 'h~*llo', 'h*llo')).toEqual(['r1']);
    expect(matches('contains', 'h~*llo', 'hello')).toEqual([]);
    expect(matches('contains', 'h~?llo', 'h?llo')).toEqual(['r1']);
    expect(matches('contains', 'h~?llo', 'hello')).toEqual([]);
    expect(matches('contains', 'a~~b', 'a~b')).toEqual(['r1']);
    expect(matches('beginsWith', 'h*o', 'hello')).toEqual(['r1']);
    expect(matches('endsWith', 'l*o', 'hello')).toEqual(['r1']);
  });

  it('matches the formatted display text, not the raw number', () => {
    const index = perCellIndex([textRule('comma', 'contains', ',')]);
    expect(index.evaluateCell(0, 0, { value: numberValue(1234) }).ruleIds).toEqual([]);
    expect(
      index.evaluateCell(0, 0, { value: numberValue(1234), displayText: '1,234' }).ruleIds,
    ).toEqual(['comma']);

    const source = sourceFrom({ '0:0': numberValue(1234567) }, { '0:0': '1,234,567' });
    const fromSource = perCellIndex([textRule('comma', 'contains', ',')], { cells: source });
    expect(fromSource.evaluateCell(0, 0).ruleIds).toEqual(['comma']);
  });

  it('never matches a blank against a non-empty rule, including not-contains', () => {
    const contains = perCellIndex([textRule('c', 'contains', 'a')]);
    const notContains = perCellIndex([textRule('n', 'notContains', 'a')]);
    expect(contains.evaluateCell(0, 0).ruleIds).toEqual([]);
    expect(notContains.evaluateCell(0, 0).ruleIds).toEqual([]);
  });

  it('never matches an error, including not-contains', () => {
    const notContains = perCellIndex([textRule('n', 'notContains', 'anything')]);
    expect(
      notContains.evaluateCell(0, 0, { value: errorValue('#N/A'), displayText: '#N/A' }).ruleIds,
    ).toEqual([]);
  });
});

// ---------------------------------------------------------------------------
// CF3 — dates
// ---------------------------------------------------------------------------

describe('CF3 cf-dates', () => {
  const comparisonCases: Array<{
    operator: CfDateOperator;
    value1: [number, number, number];
    value2?: [number, number, number];
    inside: Array<[number, number, number]>;
    outside: Array<[number, number, number]>;
  }> = [
    {
      operator: 'equal',
      value1: [2024, 3, 15],
      inside: [[2024, 3, 15]],
      outside: [
        [2024, 3, 14],
        [2024, 3, 16],
      ],
    },
    {
      operator: 'notEqual',
      value1: [2024, 3, 15],
      inside: [[2024, 3, 14]],
      outside: [[2024, 3, 15]],
    },
    {
      operator: 'before',
      value1: [2024, 3, 15],
      inside: [
        [2024, 3, 14],
        [2020, 1, 1],
      ],
      outside: [
        [2024, 3, 15],
        [2024, 3, 16],
      ],
    },
    {
      operator: 'after',
      value1: [2024, 3, 15],
      inside: [
        [2024, 3, 16],
        [2030, 1, 1],
      ],
      outside: [
        [2024, 3, 15],
        [2024, 3, 14],
      ],
    },
    {
      operator: 'between',
      value1: [2024, 3, 1],
      value2: [2024, 3, 31],
      inside: [
        [2024, 3, 1],
        [2024, 3, 15],
        [2024, 3, 31],
      ],
      outside: [
        [2024, 2, 29],
        [2024, 4, 1],
      ],
    },
    {
      operator: 'notBetween',
      value1: [2024, 3, 1],
      value2: [2024, 3, 31],
      inside: [
        [2024, 2, 29],
        [2024, 4, 1],
      ],
      outside: [
        [2024, 3, 1],
        [2024, 3, 31],
      ],
    },
  ];

  for (const date1904 of [false, true]) {
    it(`compares date serials in the ${date1904 ? '1904' : '1900'} system`, () => {
      for (const fixture of comparisonCases) {
        const index = perCellIndex(
          [
            dateComparisonRule(
              'r1',
              fixture.operator,
              serial(...fixture.value1, date1904),
              fixture.value2 === undefined ? undefined : serial(...fixture.value2, date1904),
            ),
          ],
          { date1904 },
        );
        for (const date of fixture.inside) {
          expect(
            index.evaluateCell(0, 0, { value: numberValue(serial(...date, date1904)) }).ruleIds,
          ).toEqual(['r1']);
        }
        for (const date of fixture.outside) {
          expect(
            index.evaluateCell(0, 0, { value: numberValue(serial(...date, date1904)) }).ruleIds,
          ).toEqual([]);
        }
      }
    });
  }

  const presetCases: Array<{
    period: CfTimePeriod;
    inside: CfCalendarDate;
    outside: CfCalendarDate;
  }> = [
    {
      period: 'yesterday',
      inside: { year: 2024, month: 3, day: 14 },
      outside: { year: 2024, month: 3, day: 13 },
    },
    {
      period: 'today',
      inside: { year: 2024, month: 3, day: 15 },
      outside: { year: 2024, month: 3, day: 14 },
    },
    {
      period: 'tomorrow',
      inside: { year: 2024, month: 3, day: 16 },
      outside: { year: 2024, month: 3, day: 15 },
    },
    {
      period: 'last7Days',
      inside: { year: 2024, month: 3, day: 9 },
      outside: { year: 2024, month: 3, day: 8 },
    },
    {
      period: 'lastWeek',
      inside: { year: 2024, month: 3, day: 5 },
      outside: { year: 2024, month: 3, day: 10 },
    },
    {
      period: 'thisWeek',
      inside: { year: 2024, month: 3, day: 10 },
      outside: { year: 2024, month: 3, day: 9 },
    },
    {
      period: 'nextWeek',
      inside: { year: 2024, month: 3, day: 17 },
      outside: { year: 2024, month: 3, day: 24 },
    },
    {
      period: 'lastMonth',
      inside: { year: 2024, month: 2, day: 10 },
      outside: { year: 2024, month: 3, day: 1 },
    },
    {
      period: 'thisMonth',
      inside: { year: 2024, month: 3, day: 31 },
      outside: { year: 2024, month: 4, day: 1 },
    },
    {
      period: 'nextMonth',
      inside: { year: 2024, month: 4, day: 30 },
      outside: { year: 2024, month: 5, day: 1 },
    },
  ];

  for (const date1904 of [false, true]) {
    it(`covers all ten time-period presets in the ${date1904 ? '1904' : '1900'} system`, () => {
      for (const fixture of presetCases) {
        const index = perCellIndex([datePeriodRule('r1', fixture.period)], {
          date1904,
          today: TODAY,
        });
        expect(
          index.evaluateCell(0, 0, {
            value: numberValue(
              serial(fixture.inside.year, fixture.inside.month, fixture.inside.day, date1904),
            ),
          }).ruleIds,
          fixture.period,
        ).toEqual(['r1']);
        expect(
          index.evaluateCell(0, 0, {
            value: numberValue(
              serial(fixture.outside.year, fixture.outside.month, fixture.outside.day, date1904),
            ),
          }).ruleIds,
          fixture.period,
        ).toEqual([]);
      }
    });
  }

  it('uses Sunday as the first day of a week', () => {
    const sunday: CfCalendarDate = { year: 2024, month: 3, day: 10 };
    const thisWeek = perCellIndex([datePeriodRule('w', 'thisWeek')], { today: sunday });
    expect(
      thisWeek.evaluateCell(0, 0, { value: numberValue(serial(2024, 3, 10)) }).ruleIds,
    ).toEqual(['w']);
    expect(thisWeek.evaluateCell(0, 0, { value: numberValue(serial(2024, 3, 9)) }).ruleIds).toEqual(
      [],
    );
    const lastWeek = perCellIndex([datePeriodRule('l', 'lastWeek')], { today: sunday });
    expect(lastWeek.evaluateCell(0, 0, { value: numberValue(serial(2024, 3, 9)) }).ruleIds).toEqual(
      ['l'],
    );
    expect(lastWeek.evaluateCell(0, 0, { value: numberValue(serial(2024, 3, 3)) }).ruleIds).toEqual(
      ['l'],
    );
  });

  it('crosses year boundaries for week and month periods', () => {
    const newYear: CfCalendarDate = { year: 2025, month: 1, day: 1 };
    const lastMonth = perCellIndex([datePeriodRule('lm', 'lastMonth')], { today: newYear });
    expect(
      lastMonth.evaluateCell(0, 0, { value: numberValue(serial(2024, 12, 31)) }).ruleIds,
    ).toEqual(['lm']);
    const thisWeek = perCellIndex([datePeriodRule('tw', 'thisWeek')], { today: newYear });
    expect(
      thisWeek.evaluateCell(0, 0, { value: numberValue(serial(2024, 12, 29)) }).ruleIds,
    ).toEqual(['tw']);
    const nextMonth = perCellIndex([datePeriodRule('nm', 'nextMonth')], { today: newYear });
    expect(
      nextMonth.evaluateCell(0, 0, { value: numberValue(serial(2025, 2, 28)) }).ruleIds,
    ).toEqual(['nm']);
  });

  it('includes a time fraction within the day and the leap day in the month', () => {
    const today = perCellIndex([datePeriodRule('today', 'today')]);
    expect(
      today.evaluateCell(0, 0, { value: numberValue(serial(2024, 3, 15) + 0.5) }).ruleIds,
    ).toEqual(['today']);
    const lastMonth = perCellIndex([datePeriodRule('lm', 'lastMonth')], { today: TODAY });
    expect(
      lastMonth.evaluateCell(0, 0, { value: numberValue(serial(2024, 2, 29)) }).ruleIds,
    ).toEqual(['lm']);
  });

  it('treats a blank as serial 0 for comparisons but not as a date occurrence', () => {
    const before = perCellIndex([dateComparisonRule('b', 'before', serial(2024, 1, 1))]);
    expect(before.evaluateCell(0, 0).ruleIds).toEqual(['b']);
    const todayRule = perCellIndex([datePeriodRule('t', 'today')]);
    expect(todayRule.evaluateCell(0, 0).ruleIds).toEqual([]);
    expect(todayRule.evaluateCell(0, 0, { value: errorValue('#VALUE!') }).ruleIds).toEqual([]);
  });

  it('re-evaluates when today changes', () => {
    const index = perCellIndex([datePeriodRule('today', 'today')]);
    expect(index.evaluateCell(0, 0, { value: numberValue(serial(2024, 3, 15)) }).ruleIds).toEqual([
      'today',
    ]);
    index.setToday({ year: 2024, month: 3, day: 16 });
    expect(index.evaluateCell(0, 0, { value: numberValue(serial(2024, 3, 15)) }).ruleIds).toEqual(
      [],
    );
  });
});

// ---------------------------------------------------------------------------
// CF4 — top/bottom
// ---------------------------------------------------------------------------

describe('CF4 cf-top-bottom', () => {
  function topRule(options: {
    rank?: number;
    top?: boolean;
    percent?: boolean;
    ranges?: GridRange | readonly GridRange[];
  }): ConditionalFormatRule {
    return {
      ...base('top', options.ranges ?? COLUMN_A, 1),
      kind: 'top10',
      ...(options.rank === undefined ? {} : { rank: options.rank }),
      top: options.top ?? true,
      percent: options.percent ?? false,
    } as ConditionalFormatRule;
  }

  const mixed: CellValue[] = [
    numberValue(10),
    numberValue(20),
    numberValue(30),
    numberValue(40),
    numberValue(50),
    BLANK,
    errorValue('#DIV/0!'),
    textValue('not a number'),
  ];

  it('ranks top N and includes ties, excluding blanks and errors', () => {
    const top = perCellIndex([topRule({ rank: 3 })], { cells: columnSource(mixed) });
    expect(top.statisticsFor('top', 0)).toMatchObject({ kind: 'top10', count: 5, threshold: 30 });
    for (const [row, expected] of [
      [0, false],
      [1, false],
      [2, true],
      [3, true],
      [4, true],
      [5, false],
      [6, false],
      [7, false],
    ] as const) {
      expect(top.evaluateCell(row, 0).ruleIds, `row ${row}`).toEqual(expected ? ['top'] : []);
    }

    const ties = perCellIndex([topRule({ rank: 3 })], {
      cells: columnSource([
        numberValue(10),
        numberValue(20),
        numberValue(30),
        numberValue(30),
        numberValue(40),
      ]),
    });
    const painted = [0, 1, 2, 3, 4].filter((row) => ties.evaluateCell(row, 0).ruleIds.length === 1);
    expect(painted).toEqual([2, 3, 4]);
  });

  it('ranks bottom N', () => {
    const bottom = perCellIndex([topRule({ rank: 2, top: false })], {
      cells: columnSource(mixed),
    });
    expect(bottom.statisticsFor('top', 0)).toMatchObject({ threshold: 20 });
    expect(bottom.evaluateCell(0, 0).ruleIds).toEqual(['top']);
    expect(bottom.evaluateCell(1, 0).ruleIds).toEqual(['top']);
    expect(bottom.evaluateCell(2, 0).ruleIds).toEqual([]);
  });

  it('supports N% with rounding up and at least one cell', () => {
    const values = Array.from({ length: 10 }, (_, index) => numberValue(index + 1));
    const tenPercent = perCellIndex([topRule({ rank: 10, percent: true })], {
      cells: columnSource(values),
    });
    expect(tenPercent.statisticsFor('top', 0)).toMatchObject({ count: 10, threshold: 10 });
    expect(tenPercent.evaluateCell(9, 0).ruleIds).toEqual(['top']);
    expect(tenPercent.evaluateCell(8, 0).ruleIds).toEqual([]);

    const thirtyPercent = perCellIndex([topRule({ rank: 30, percent: true })], {
      cells: columnSource(values),
    });
    expect(thirtyPercent.statisticsFor('top', 0)).toMatchObject({ threshold: 8 });
    expect(thirtyPercent.evaluateCell(7, 0).ruleIds).toEqual(['top']);
    expect(thirtyPercent.evaluateCell(6, 0).ruleIds).toEqual([]);
  });

  it('defaults to N=10', () => {
    const values = Array.from({ length: 12 }, (_, index) => numberValue(index + 1));
    const index = perCellIndex([topRule({})], { cells: columnSource(values) });
    expect(index.statisticsFor('top', 0)).toMatchObject({ rank: 10, threshold: 3 });
    expect(index.evaluateCell(2, 0).ruleIds).toEqual(['top']);
    expect(index.evaluateCell(1, 0).ruleIds).toEqual([]);
  });

  it('ranks each contiguous range as its own domain', () => {
    const values: Record<string, CellValue> = {};
    for (let row = 0; row < 5; row += 1) {
      values[`${row}:0`] = numberValue(row + 1);
      values[`${row}:2`] = numberValue((row + 1) * 100);
    }
    const rule = topRule({
      rank: 2,
      ranges: [range(0, 0, 4, 0), range(0, 2, 4, 2)],
    });
    const index = perCellIndex([rule], { cells: sourceFrom(values) });
    expect(index.evaluateCell(3, 0).ruleIds).toEqual(['top']);
    expect(index.evaluateCell(4, 0).ruleIds).toEqual(['top']);
    expect(index.evaluateCell(2, 0).ruleIds).toEqual([]);
    expect(index.evaluateCell(3, 2).ruleIds).toEqual(['top']);
    expect(index.evaluateCell(4, 2).ruleIds).toEqual(['top']);
    expect(index.evaluateCell(2, 2).ruleIds).toEqual([]);
    expect(index.statisticsFor('top', 1)).toMatchObject({ count: 5, threshold: 400 });
  });
});

// ---------------------------------------------------------------------------
// CF5 — duplicates and unique
// ---------------------------------------------------------------------------

describe('CF5 cf-duplicates', () => {
  const values: CellValue[] = [
    numberValue(1),
    numberValue(1),
    textValue('1'),
    BLANK,
    errorValue('#N/A'),
    errorValue('#N/A'),
    textValue('ABC'),
    textValue('abc'),
  ];

  function duplicateRule(kind: 'duplicate' | 'unique'): ConditionalFormatRule {
    return { ...base(kind, COLUMN_A, 1), kind } as ConditionalFormatRule;
  }

  it('paints duplicates by typed value and ignores blanks', () => {
    const index = perCellIndex([duplicateRule('duplicate')], { cells: columnSource(values) });
    expect(index.evaluateCell(0, 0).ruleIds).toEqual(['duplicate']);
    expect(index.evaluateCell(1, 0).ruleIds).toEqual(['duplicate']);
    expect(index.evaluateCell(2, 0).ruleIds).toEqual([]); // text "1" is not the number 1
    expect(index.evaluateCell(3, 0).ruleIds).toEqual([]); // blanks are never duplicates
    expect(index.evaluateCell(4, 0).ruleIds).toEqual(['duplicate']); // errors can be duplicates
    expect(index.evaluateCell(5, 0).ruleIds).toEqual(['duplicate']);
    expect(index.evaluateCell(6, 0).ruleIds).toEqual(['duplicate']);
    expect(index.evaluateCell(7, 0).ruleIds).toEqual(['duplicate']);
  });

  it('paints only the unique values', () => {
    const index = perCellIndex([duplicateRule('unique')], { cells: columnSource(values) });
    expect(index.evaluateCell(0, 0).ruleIds).toEqual([]);
    expect(index.evaluateCell(2, 0).ruleIds).toEqual(['unique']);
    expect(index.evaluateCell(3, 0).ruleIds).toEqual([]);
    expect(index.evaluateCell(4, 0).ruleIds).toEqual([]);
    expect(index.evaluateCell(6, 0).ruleIds).toEqual([]);
  });

  it('uses each contiguous range as its own comparison domain', () => {
    const rule = {
      ...base('duplicate', [range(0, 0, 2, 0), range(0, 2, 2, 2)], 1),
      kind: 'duplicate',
    } as ConditionalFormatRule;
    const index = perCellIndex([rule], {
      cells: sourceFrom({
        '0:0': numberValue(5),
        '1:0': numberValue(5),
        '2:0': numberValue(6),
        '0:2': numberValue(5),
        '1:2': numberValue(6),
        '2:2': numberValue(6),
      }),
    });
    expect([0, 1, 2].map((row) => index.evaluateCell(row, 0).ruleIds.length)).toEqual([1, 1, 0]);
    expect([0, 1, 2].map((row) => index.evaluateCell(row, 2).ruleIds.length)).toEqual([0, 1, 1]);
  });
});

// ---------------------------------------------------------------------------
// CF6 / CF7 — colour scales
// ---------------------------------------------------------------------------

describe('CF6 cf-colour-scale-2', () => {
  function scaleRule(
    thresholds: readonly CfThreshold[],
    colors: readonly string[],
    id = 'scale',
  ): ConditionalFormatRule {
    return {
      ...base(id, COLUMN_A, 1),
      kind: 'colorScale',
      thresholds,
      colors,
    } as ConditionalFormatRule;
  }

  function scaleIndex(rule: ConditionalFormatRule, values: readonly number[]) {
    return perCellIndex([rule], {
      cells: columnSource(values.map((value) => numberValue(value))),
    });
  }

  it('interpolates linearly in sRGB between automatic min and max', () => {
    const index = scaleIndex(
      scaleRule(
        [threshold({ type: 'automatic' }), threshold({ type: 'automatic' })],
        [BLACK, WHITE],
      ),
      [0, 25, 50, 75, 100],
    );
    const mid = index.evaluateCell(2, 0).colorScale!;
    expect(mid.position).toBeCloseTo(0.5, 12);
    expect(mid.color).toBe('#808080');
    const quarter = index.evaluateCell(1, 0).colorScale!;
    expect(quarter.position).toBeCloseTo(0.25, 12);
    expect(quarter.color).toBe('#404040');
    expect(index.evaluateCell(0, 0).colorScale).toMatchObject({ position: 0, color: BLACK });
    expect(index.evaluateCell(4, 0).colorScale).toMatchObject({ position: 1, color: WHITE });
  });

  it('accepts OOXML ARGB colours and drops the alpha byte', () => {
    const index = scaleIndex(
      scaleRule(
        [threshold({ type: 'number', value: 0 }), threshold({ type: 'number', value: 10 })],
        ['FFFF0000', 'FF0000FF'],
      ),
      [0, 5, 10],
    );
    expect(index.statisticsFor('scale', 0)).toMatchObject({
      points: [
        { value: 0, color: RED },
        { value: 10, color: BLUE },
      ],
    });
  });

  it('uses number thresholds as fixed points', () => {
    const index = scaleIndex(
      scaleRule(
        [threshold({ type: 'number', value: 10 }), threshold({ type: 'number', value: 20 })],
        [RED, BLUE],
      ),
      [5, 10, 15, 20, 25],
    );
    expect(index.evaluateCell(0, 0).colorScale).toMatchObject({ position: 0, color: RED });
    expect(index.evaluateCell(2, 0).colorScale!.position).toBeCloseTo(0.5, 12);
    expect(index.evaluateCell(2, 0).colorScale!.color).toBe('#800080');
    expect(index.evaluateCell(4, 0).colorScale).toMatchObject({ position: 1, color: BLUE });
  });

  it('resolves percent thresholds against the range width', () => {
    const index = scaleIndex(
      scaleRule(
        [threshold({ type: 'percent', value: 25 }), threshold({ type: 'percent', value: 75 })],
        [RED, BLUE],
      ),
      [0, 25, 50, 75, 100],
    );
    expect(index.statisticsFor('scale', 0)).toMatchObject({
      kind: 'colorScale',
      points: [
        { value: 25, color: RED },
        { value: 75, color: BLUE },
      ],
    });
    expect(index.evaluateCell(2, 0).colorScale!.position).toBeCloseTo(0.5, 12);
  });

  it('resolves percentile thresholds from the ordered values', () => {
    const index = scaleIndex(
      scaleRule(
        [
          threshold({ type: 'percentile', value: 25 }),
          threshold({ type: 'percentile', value: 75 }),
        ],
        [RED, BLUE],
      ),
      [1, 2, 3, 4, 5, 6, 7, 8, 9],
    );
    expect(index.statisticsFor('scale', 0)).toMatchObject({
      points: [{ value: 3 }, { value: 7 }],
    });
    expect(index.evaluateCell(4, 0).colorScale!.position).toBeCloseTo(0.5, 12);
  });

  it('does not paint blanks, errors or text and excludes them from the statistics', () => {
    const index = perCellIndex(
      [
        scaleRule(
          [threshold({ type: 'automatic' }), threshold({ type: 'automatic' })],
          [BLACK, WHITE],
        ),
      ],
      {
        cells: columnSource([
          numberValue(10),
          BLANK,
          errorValue('#N/A'),
          textValue('x'),
          numberValue(20),
        ]),
      },
    );
    expect(index.statisticsFor('scale', 0)).toMatchObject({
      points: [{ value: 10 }, { value: 20 }],
    });
    expect(index.evaluateCell(1, 0).colorScale).toBeNull();
    expect(index.evaluateCell(2, 0).colorScale).toBeNull();
    expect(index.evaluateCell(3, 0).colorScale).toBeNull();
  });
});

describe('CF7 cf-colour-scale-3', () => {
  function threeColourRule(thresholds: readonly CfThreshold[]): ConditionalFormatRule {
    return {
      ...base('scale3', COLUMN_A, 1),
      kind: 'colorScale',
      thresholds,
      colors: [RED, YELLOW, GREEN],
    } as ConditionalFormatRule;
  }

  it('interpolates each segment and places points across the scale', () => {
    const index = perCellIndex(
      [
        threeColourRule([
          threshold({ type: 'percent', value: 0 }),
          threshold({ type: 'percent', value: 50 }),
          threshold({ type: 'percent', value: 100 }),
        ]),
      ],
      { cells: columnSource([0, 25, 50, 75, 100].map((value) => numberValue(value))) },
    );
    expect(index.evaluateCell(0, 0).colorScale).toMatchObject({ position: 0, color: RED });
    const quarter = index.evaluateCell(1, 0).colorScale!;
    expect(quarter.position).toBeCloseTo(0.25, 12);
    expect(quarter.color).toBe('#FF8000');
    expect(index.evaluateCell(2, 0).colorScale).toMatchObject({ position: 0.5, color: YELLOW });
    const threeQuarters = index.evaluateCell(3, 0).colorScale!;
    expect(threeQuarters.position).toBeCloseTo(0.75, 12);
    expect(threeQuarters.color).toBe('#80FF00');
    expect(index.evaluateCell(4, 0).colorScale).toMatchObject({ position: 1, color: GREEN });
  });

  it('defaults the midpoint to the 50th percentile', () => {
    const index = perCellIndex(
      [
        threeColourRule([
          threshold({ type: 'automatic' }),
          threshold({ type: 'automatic' }),
          threshold({ type: 'automatic' }),
        ]),
      ],
      { cells: columnSource([0, 25, 50, 75, 100].map((value) => numberValue(value))) },
    );
    expect(index.statisticsFor('scale3', 0)).toMatchObject({
      points: [
        { value: 0, color: RED },
        { value: 50, color: YELLOW },
        { value: 100, color: GREEN },
      ],
    });
    expect(index.evaluateCell(1, 0).colorScale!.color).toBe('#FF8000');
  });

  it('keeps positions at the bounds for values outside the scale', () => {
    const index = perCellIndex(
      [
        threeColourRule([
          threshold({ type: 'number', value: 10 }),
          threshold({ type: 'number', value: 50 }),
          threshold({ type: 'number', value: 90 }),
        ]),
      ],
      { cells: columnSource([-100, 10, 50, 90, 1000].map((value) => numberValue(value))) },
    );
    expect(index.evaluateCell(0, 0).colorScale).toMatchObject({ position: 0, color: RED });
    expect(index.evaluateCell(4, 0).colorScale).toMatchObject({ position: 1, color: GREEN });
  });
});

// ---------------------------------------------------------------------------
// CF8 / CF9 — data bars
// ---------------------------------------------------------------------------

function dataBarRule(options: {
  min?: CfThreshold;
  max?: CfThreshold;
  fill?: 'gradient' | 'solid';
  showValue?: boolean;
  border?: boolean;
  direction?: 'leftToRight' | 'rightToLeft';
  negativeColor?: string;
  id?: string;
}): ConditionalFormatRule {
  return {
    ...base(options.id ?? 'bar', COLUMN_A, 1),
    kind: 'dataBar',
    min: options.min ?? threshold({ type: 'automatic' }),
    max: options.max ?? threshold({ type: 'automatic' }),
    color: RED,
    ...(options.negativeColor === undefined ? {} : { negativeColor: options.negativeColor }),
    ...(options.fill === undefined ? {} : { fill: options.fill }),
    ...(options.showValue === undefined ? {} : { showValue: options.showValue }),
    ...(options.border === undefined ? {} : { border: options.border }),
    ...(options.direction === undefined ? {} : { direction: options.direction }),
  } as ConditionalFormatRule;
}

function dataBarIndex(rule: ConditionalFormatRule, values: readonly number[]) {
  return perCellIndex([rule], {
    cells: columnSource(values.map((value) => numberValue(value))),
  });
}

describe('CF8 cf-data-bar-positive', () => {
  it('uses automatic min 0 and max from the range, with a gradient by default', () => {
    const index = dataBarIndex(dataBarRule({}), [10, 50, 100]);
    expect(index.statisticsFor('bar', 0)).toMatchObject({ kind: 'dataBar', min: 0, max: 100 });
    const bar = index.evaluateCell(1, 0).dataBar!;
    expect(bar.start).toBeCloseTo(0, 12);
    expect(bar.end).toBeCloseTo(0.5, 12);
    expect(bar.axis).toBeCloseTo(0, 12);
    expect(bar.negative).toBe(false);
    expect(bar.fill).toBe('gradient');
    expect(bar.showValue).toBe(true);
    expect(bar.color).toBe(RED);
  });

  it('passes solid fill and show/hide value through', () => {
    const index = dataBarIndex(dataBarRule({ fill: 'solid', showValue: false }), [10, 50, 100]);
    const bar = index.evaluateCell(1, 0).dataBar!;
    expect(bar.fill).toBe('solid');
    expect(bar.showValue).toBe(false);
  });

  it('resolves fixed number thresholds', () => {
    const index = dataBarIndex(
      dataBarRule({
        min: threshold({ type: 'number', value: 0 }),
        max: threshold({ type: 'number', value: 200 }),
      }),
      [0, 50, 200],
    );
    expect(index.statisticsFor('bar', 0)).toMatchObject({ min: 0, max: 200 });
    expect(index.evaluateCell(1, 0).dataBar!.end).toBeCloseTo(0.25, 12);
  });

  it('clamps the axis when a fixed domain excludes zero', () => {
    const index = dataBarIndex(
      dataBarRule({
        min: threshold({ type: 'number', value: 10 }),
        max: threshold({ type: 'number', value: 110 }),
      }),
      [10, 60, 110],
    );
    const bar = index.evaluateCell(1, 0).dataBar!;
    expect(bar.start).toBeCloseTo(0, 12);
    expect(bar.end).toBeCloseTo(0.5, 12);
    expect(bar.axis).toBeCloseTo(0, 12);
  });

  it('resolves percent and percentile thresholds', () => {
    const percent = dataBarIndex(
      dataBarRule({
        min: threshold({ type: 'percent', value: 20 }),
        max: threshold({ type: 'percent', value: 80 }),
      }),
      [0, 50, 100],
    );
    expect(percent.statisticsFor('bar', 0)).toMatchObject({ min: 20, max: 80 });
    expect(percent.evaluateCell(1, 0).dataBar!.end).toBeCloseTo(0.5, 12);

    const percentile = dataBarIndex(
      dataBarRule({
        min: threshold({ type: 'percentile', value: 20 }),
        max: threshold({ type: 'percentile', value: 80 }),
      }),
      [1, 2, 3, 4, 5, 6, 7, 8, 9, 10],
    );
    const statistics = percentile.statisticsFor('bar', 0);
    expect(statistics).toMatchObject({ kind: 'dataBar' });
    if (statistics?.kind !== 'dataBar') throw new Error('expected data-bar statistics');
    expect(statistics.min).toBeCloseTo(2.8, 12);
    expect(statistics.max).toBeCloseTo(8.2, 12);
    const bar = percentile.evaluateCell(4, 0).dataBar!;
    expect(bar.end).toBeCloseTo((5 - 2.8) / (8.2 - 2.8), 12);
  });

  it('does not paint blanks, errors or text and excludes them from the statistics', () => {
    const index = perCellIndex([dataBarRule({})], {
      cells: columnSource([
        numberValue(10),
        BLANK,
        errorValue('#N/A'),
        textValue('x'),
        numberValue(20),
      ]),
    });
    expect(index.statisticsFor('bar', 0)).toMatchObject({ min: 0, max: 20 });
    expect(index.evaluateCell(1, 0).dataBar).toBeNull();
    expect(index.evaluateCell(2, 0).dataBar).toBeNull();
    expect(index.evaluateCell(3, 0).dataBar).toBeNull();
  });
});

describe('CF9 cf-data-bar-negative', () => {
  it('puts the axis at zero for mixed signs and uses the negative colour', () => {
    const index = dataBarIndex(dataBarRule({ negativeColor: BLUE }), [-50, -25, 0, 25, 50]);
    expect(index.statisticsFor('bar', 0)).toMatchObject({ min: -50, max: 50 });
    const positive = index.evaluateCell(3, 0).dataBar!;
    expect(positive.axis).toBeCloseTo(0.5, 12);
    expect(positive.start).toBeCloseTo(0.5, 12);
    expect(positive.end).toBeCloseTo(0.75, 12);
    expect(positive.negative).toBe(false);
    expect(positive.color).toBe(RED);
    const negative = index.evaluateCell(1, 0).dataBar!;
    expect(negative.start).toBeCloseTo(0.25, 12);
    expect(negative.end).toBeCloseTo(0.5, 12);
    expect(negative.negative).toBe(true);
    expect(negative.color).toBe(BLUE);
    const zero = index.evaluateCell(2, 0).dataBar!;
    expect(zero.start).toBeCloseTo(0.5, 12);
    expect(zero.end).toBeCloseTo(0.5, 12);
    expect(zero.negative).toBe(false);
  });

  it('keeps the axis meaningful when only negative values exist', () => {
    const index = dataBarIndex(dataBarRule({}), [-50, -10]);
    expect(index.statisticsFor('bar', 0)).toMatchObject({ min: -50, max: 0 });
    const bar = index.evaluateCell(1, 0).dataBar!;
    expect(bar.axis).toBeCloseTo(1, 12);
    expect(bar.start).toBeCloseTo(0.8, 12);
    expect(bar.end).toBeCloseTo(1, 12);
  });

  it('passes border and direction through', () => {
    const index = dataBarIndex(
      dataBarRule({ border: true, direction: 'rightToLeft', negativeColor: BLUE }),
      [-50, 50],
    );
    const bar = index.evaluateCell(0, 0).dataBar!;
    expect(bar.border).toBe(true);
    expect(bar.direction).toBe('rightToLeft');
    expect(bar.color).toBe(BLUE);
  });
});

// ---------------------------------------------------------------------------
// CF10 — icon sets
// ---------------------------------------------------------------------------

function iconSetRule(
  iconSet: CfIconSetId,
  thresholds: readonly CfThreshold[],
  options: { reverse?: boolean; showValue?: boolean; iconOnly?: boolean; id?: string } = {},
): ConditionalFormatRule {
  return {
    ...base(options.id ?? 'icons', COLUMN_A, 1),
    kind: 'iconSet',
    iconSet,
    thresholds,
    ...(options.reverse === undefined ? {} : { reverse: options.reverse }),
    ...(options.showValue === undefined ? {} : { showValue: options.showValue }),
    ...(options.iconOnly === undefined ? {} : { iconOnly: options.iconOnly }),
  } as ConditionalFormatRule;
}

function iconIndexFor(rule: ConditionalFormatRule, values: readonly number[], row: number) {
  const index = perCellIndex([rule], {
    cells: columnSource(values.map((value) => numberValue(value))),
  });
  return index.evaluateCell(row, 0).iconSet;
}

describe('CF10 cf-icon-sets', () => {
  it('knows every built-in set and its icon count', () => {
    expect(CF_ICON_SETS).toHaveLength(17);
    const ids = new Set(CF_ICON_SETS.map((definition) => definition.id));
    expect(ids.size).toBe(17);
    for (const definition of CF_ICON_SETS) {
      const thresholds: CfThreshold[] = Array.from(
        { length: definition.iconCount - 1 },
        (_, index) => ({ type: 'number', value: (index + 1) * 20 }),
      );
      const overlay = iconIndexFor(
        iconSetRule(definition.id, thresholds),
        [0, 10, 30, 50, 70, 100],
        5,
      )!;
      expect(overlay.iconSet).toBe(definition.id);
      expect(overlay.iconCount).toBe(definition.iconCount);
      expect(overlay.iconIndex).toBeGreaterThanOrEqual(0);
      expect(overlay.iconIndex).toBeLessThan(definition.iconCount);
    }
  });

  it('selects icons by number thresholds at the boundaries', () => {
    const rule = iconSetRule('3Arrows', [
      threshold({ type: 'number', value: 20 }),
      threshold({ type: 'number', value: 80 }),
    ]);
    expect(iconIndexFor(rule, [0, 10, 50, 90, 100], 0)!.iconIndex).toBe(0);
    expect(iconIndexFor(rule, [0, 10, 50, 90, 100], 2)!.iconIndex).toBe(1);
    expect(iconIndexFor(rule, [0, 10, 50, 90, 100], 3)!.iconIndex).toBe(2);
    // The threshold itself belongs to the higher icon.
    const boundary = iconIndexFor(rule, [0, 10, 20, 79, 80], 2)!;
    expect(boundary.iconIndex).toBe(1);
    expect(iconIndexFor(rule, [0, 10, 20, 79, 80], 4)!.iconIndex).toBe(2);
  });

  it('resolves percent and percentile thresholds', () => {
    const percent = iconIndexFor(
      iconSetRule('3TrafficLights1', [
        threshold({ type: 'percent', value: 33 }),
        threshold({ type: 'percent', value: 67 }),
      ]),
      [0, 20, 50, 80, 100],
      1,
    )!;
    expect(percent.thresholds).toEqual([33, 67]);

    const percentile = iconIndexFor(
      iconSetRule('5Quarters', [
        threshold({ type: 'percentile', value: 20 }),
        threshold({ type: 'percentile', value: 40 }),
        threshold({ type: 'percentile', value: 60 }),
        threshold({ type: 'percentile', value: 80 }),
      ]),
      [1, 2, 3, 4, 5, 6, 7, 8, 9, 10],
      4,
    )!;
    const expected = [2.8, 4.6, 6.4, 8.2];
    expect(percentile.thresholds).toHaveLength(expected.length);
    percentile.thresholds.forEach((value, position) => {
      expect(value).toBeCloseTo(expected[position]!, 12);
    });
    expect(percentile.iconCount).toBe(5);
  });

  it('applies reverse to the chosen icon index', () => {
    const rule = iconSetRule(
      '3Arrows',
      [threshold({ type: 'number', value: 20 }), threshold({ type: 'number', value: 80 })],
      { reverse: true },
    );
    expect(iconIndexFor(rule, [0, 10, 50, 90, 100], 3)!.iconIndex).toBe(0);
    expect(iconIndexFor(rule, [0, 10, 50, 90, 100], 0)!.iconIndex).toBe(2);
  });

  it('carries show-value and icon-only', () => {
    const rule = iconSetRule('3Flags', [threshold({ type: 'number', value: 50 })]);
    expect(iconIndexFor(rule, [10, 90], 0)).toMatchObject({ showValue: true, iconOnly: false });

    const hidden = iconSetRule('3Flags', [threshold({ type: 'number', value: 50 })], {
      showValue: false,
    });
    expect(iconIndexFor(hidden, [10, 90], 1)).toMatchObject({ showValue: false, iconOnly: false });

    const iconOnly = iconSetRule('3Flags', [threshold({ type: 'number', value: 50 })], {
      iconOnly: true,
    });
    expect(iconIndexFor(iconOnly, [10, 90], 1)).toMatchObject({
      showValue: false,
      iconOnly: true,
    });
  });

  it('does not paint blanks, errors or text', () => {
    const rule = iconSetRule('3Signs', [threshold({ type: 'number', value: 50 })]);
    const index = perCellIndex([rule], {
      cells: columnSource([BLANK, errorValue('#N/A'), textValue('x')]),
    });
    expect(index.evaluateCell(0, 0).iconSet).toBeNull();
    expect(index.evaluateCell(1, 0).iconSet).toBeNull();
    expect(index.evaluateCell(2, 0).iconSet).toBeNull();
  });
});

// ---------------------------------------------------------------------------
// CF11 — priority, merge and stop-if-true
// ---------------------------------------------------------------------------

describe('CF11 cf-priority', () => {
  const dxfs = new Map<number, CfDifferentialFormat>([
    [1, { fill: { pattern: 'solid', color: RED } }],
    [2, { font: { bold: true }, numberFormat: { code: '0.00' } }],
    [3, { border: { top: { style: 'thin', color: BLUE } } }],
    [4, { fill: { pattern: 'solid', color: BLUE } }],
  ]);
  const dxfFor = (dxfId: number): CfDifferentialFormat | null => dxfs.get(dxfId) ?? null;

  function priorityIndex(
    rules: readonly ConditionalFormatRule[],
    value: CellValue = numberValue(5),
  ) {
    const index = createConditionalFormatIndex({
      sheetId: SHEET,
      rules,
      date1904: false,
      today: TODAY,
      dxfFor,
    });
    return index.evaluateCell(0, 0, { value });
  }

  it('merges property by property with the first true rule winning each property', () => {
    const evaluation = priorityIndex([
      cellIsRule('r1', 'greaterThan', numberOperand(0), undefined, { priority: 1, dxfId: 1 }),
      cellIsRule('r2', 'greaterThan', numberOperand(0), undefined, { priority: 2, dxfId: 2 }),
    ]);
    expect(evaluation.ruleIds).toEqual(['r1', 'r2']);
    expect(evaluation.overlay.fill).toEqual({ pattern: 'solid', color: RED });
    expect(evaluation.overlay.font).toEqual({ bold: true });
    expect(evaluation.overlay.numberFormat).toEqual({ code: '0.00' });
    expect(evaluation.overlay.border).toBeNull();
    expect(evaluation.painted).toBe(true);
  });

  it('keeps the first writer of a shared property', () => {
    const evaluation = priorityIndex([
      cellIsRule('blue', 'greaterThan', numberOperand(0), undefined, {
        priority: 1,
        dxfId: 4,
      }),
      cellIsRule('red', 'greaterThan', numberOperand(0), undefined, { priority: 2, dxfId: 1 }),
    ]);
    expect(evaluation.overlay.fill).toEqual({ pattern: 'solid', color: BLUE });
  });

  it('lets a lower-priority rule fill a property the higher rule does not set', () => {
    const evaluation = priorityIndex([
      cellIsRule('high', 'lessThan', numberOperand(0), undefined, { priority: 1, dxfId: 1 }),
      cellIsRule('low', 'greaterThan', numberOperand(0), undefined, { priority: 2, dxfId: 3 }),
    ]);
    expect(evaluation.ruleIds).toEqual(['low']);
    expect(evaluation.overlay.fill).toBeNull();
    expect(evaluation.overlay.border).toEqual({ top: { style: 'thin', color: BLUE } });
  });

  it('stops lower rules after a true stop-if-true rule', () => {
    const evaluation = priorityIndex([
      cellIsRule('stop', 'greaterThan', numberOperand(0), undefined, {
        priority: 1,
        dxfId: 1,
        stopIfTrue: true,
      }),
      cellIsRule('later', 'greaterThan', numberOperand(0), undefined, {
        priority: 2,
        dxfId: 2,
      }),
    ]);
    expect(evaluation.stoppedBy).toBe('stop');
    expect(evaluation.ruleIds).toEqual(['stop']);
    expect(evaluation.overlay.font).toBeNull();
    expect(evaluation.overlay.fill).toEqual({ pattern: 'solid', color: RED });
  });

  it('orders by priority and falls back to document order for duplicates and missing values', () => {
    const rules = [
      cellIsRule('third', 'greaterThan', numberOperand(0), undefined, { priority: 3 }),
      cellIsRule('first', 'greaterThan', numberOperand(0), undefined, { priority: 1 }),
      cellIsRule('second', 'greaterThan', numberOperand(0), undefined, { priority: 2 }),
    ];
    expect(priorityIndex(rules).ruleIds).toEqual(['first', 'second', 'third']);
    expect(orderRulesByPriority(rules).map((rule) => rule.id)).toEqual([
      'first',
      'second',
      'third',
    ]);

    const duplicate = [
      cellIsRule('a', 'greaterThan', numberOperand(0), undefined, { priority: 1 }),
      cellIsRule('b', 'greaterThan', numberOperand(0), undefined, { priority: 1 }),
      cellIsRule('c', 'greaterThan', numberOperand(0), undefined, { priority: 0 }),
    ];
    expect(priorityIndex(duplicate).ruleIds).toEqual(['c', 'a', 'b']);

    const missing = [
      {
        ...base('missing', COLUMN_A, undefined),
        kind: 'cellIs',
        operator: 'greaterThan',
        value1: numberOperand(0),
      } as ConditionalFormatRule,
      cellIsRule('known', 'greaterThan', numberOperand(0), undefined, { priority: 1 }),
    ];
    expect(priorityIndex(missing).ruleIds).toEqual(['known', 'missing']);
  });

  it('never paints an unsupported rule and never lets it stop lower rules', () => {
    const unsupported = {
      ...base('formula', COLUMN_A, 1, true),
      kind: 'unsupported',
      reason: 'formula',
      ruleType: 'expression',
    } as ConditionalFormatRule;
    const evaluation = priorityIndex([
      unsupported,
      cellIsRule('paint', 'greaterThan', numberOperand(0), undefined, { priority: 2, dxfId: 1 }),
    ]);
    expect(evaluation.ruleIds).toEqual(['paint']);
    expect(evaluation.unsupportedRuleIds).toEqual(['formula']);
    expect(evaluation.stoppedBy).toBeNull();
    expect(evaluation.overlay.fill).toEqual({ pattern: 'solid', color: RED });
  });

  it('shows manual formatting through when no rule is true', () => {
    const evaluation = priorityIndex([
      cellIsRule('never', 'greaterThan', numberOperand(1000), undefined, { priority: 1, dxfId: 1 }),
    ]);
    expect(evaluation.ruleIds).toEqual([]);
    expect(evaluation.overlay).toEqual({
      fill: null,
      font: null,
      border: null,
      numberFormat: null,
    });
    expect(evaluation.painted).toBe(false);
  });

  it('keeps the highest-priority visual overlay and ignores dxf-free duplicates', () => {
    const firstScale = {
      ...base('scale-a', COLUMN_A, 1),
      kind: 'colorScale',
      thresholds: [
        threshold({ type: 'number', value: 0 }),
        threshold({ type: 'number', value: 10 }),
      ],
      colors: [RED, BLUE],
    } as ConditionalFormatRule;
    const secondScale = {
      ...base('scale-b', COLUMN_A, 2),
      kind: 'colorScale',
      thresholds: [
        threshold({ type: 'number', value: 0 }),
        threshold({ type: 'number', value: 10 }),
      ],
      colors: [GREEN, BLACK],
    } as ConditionalFormatRule;
    const index = perCellIndex([firstScale, secondScale], {
      cells: columnSource([numberValue(5)]),
    });
    const evaluation = index.evaluateCell(0, 0);
    expect(evaluation.ruleIds).toEqual(['scale-a', 'scale-b']);
    expect(evaluation.colorScale!.ruleId).toBe('scale-a');
  });
});

// ---------------------------------------------------------------------------
// §6 caching and invalidation
// ---------------------------------------------------------------------------

describe('CF evaluation caching and invalidation', () => {
  it('reuses per-cell results until the cell is invalidated', () => {
    const values: Record<string, CellValue> = { '0:0': numberValue(5) };
    const source = sourceFrom(values);
    const index = perCellIndex([cellIsRule('r', 'lessThan', numberOperand(20))], {
      cells: source,
    });
    expect(index.evaluateCell(0, 0).ruleIds).toEqual(['r']);
    values['0:0'] = numberValue(50);
    // The per-cell cache is keyed by rule plus value, so a new value misses it.
    expect(index.evaluateCell(0, 0).ruleIds).toEqual([]);
    values['0:0'] = numberValue(5);
    expect(index.evaluateCell(0, 0).ruleIds).toEqual(['r']);
  });

  it('invalidates range statistics for ranges containing the edited cell only', () => {
    const values: Record<string, CellValue> = {
      '0:0': numberValue(10),
      '1:0': numberValue(20),
      '0:2': numberValue(100),
      '1:2': numberValue(200),
    };
    const rule = {
      ...base('top', [range(0, 0, 1, 0), range(0, 2, 1, 2)], 1),
      kind: 'top10',
      rank: 1,
      top: true,
      percent: false,
    } as ConditionalFormatRule;
    const index = perCellIndex([rule], { cells: sourceFrom(values) });
    expect(index.statisticsFor('top', 0)).toMatchObject({ threshold: 20 });
    expect(index.statisticsFor('top', 1)).toMatchObject({ threshold: 200 });

    values['0:0'] = numberValue(1000);
    index.invalidateCell(0, 0);
    expect(index.statisticsFor('top', 0)).toMatchObject({ threshold: 1000 });
    // The other contiguous range was not invalidated and keeps its cached result.
    expect(index.statisticsFor('top', 1)).toMatchObject({ threshold: 200 });
    values['0:2'] = numberValue(5000);
    expect(index.statisticsFor('top', 1)).toMatchObject({ threshold: 200 });
    index.invalidateCell(0, 2);
    expect(index.statisticsFor('top', 1)).toMatchObject({ threshold: 5000 });
  });

  it('rebuilds the index and clears caches when the rules or dxf table change', () => {
    const index = perCellIndex([cellIsRule('r', 'lessThan', numberOperand(20))]);
    expect(index.evaluateCell(0, 0, { value: numberValue(5) }).ruleIds).toEqual(['r']);

    index.setRules([cellIsRule('s', 'greaterThan', numberOperand(0), undefined, { dxfId: 1 })]);
    expect(index.rules.map((rule) => rule.id)).toEqual(['s']);
    expect(index.evaluateCell(0, 0, { value: numberValue(5) }).ruleIds).toEqual(['s']);
    expect(index.evaluateCell(0, 0, { value: numberValue(5) }).overlay.fill).toBeNull();

    index.setDxfResolver(() => ({ fill: { color: RED } }));
    expect(index.evaluateCell(0, 0, { value: numberValue(5) }).overlay.fill).toEqual({
      color: RED,
    });
  });

  it('re-evaluates date rules when the workbook date system changes', () => {
    const index = perCellIndex([datePeriodRule('today', 'today')]);
    const serial1900 = serial(2024, 3, 15, false);
    expect(index.evaluateCell(0, 0, { value: numberValue(serial1900) }).ruleIds).toEqual(['today']);
    index.setDateSystem(true);
    expect(index.evaluateCell(0, 0, { value: numberValue(serial1900) }).ruleIds).toEqual([]);
    expect(
      index.evaluateCell(0, 0, { value: numberValue(serial(2024, 3, 15, true)) }).ruleIds,
    ).toEqual(['today']);
  });

  it('builds a range index and reports the rules covering a cell in priority order', () => {
    const wide = {
      ...base('wide', range(0, 0, 99, 5), 2),
      kind: 'cellIs',
      operator: 'greaterThan',
      value1: numberOperand(0),
    } as ConditionalFormatRule;
    const narrow = {
      ...base('narrow', range(5, 2, 5, 2), 1),
      kind: 'cellIs',
      operator: 'greaterThan',
      value1: numberOperand(0),
    } as ConditionalFormatRule;
    const index = perCellIndex([wide, narrow]);
    expect(index.ruleIdsAt(5, 2)).toEqual(['narrow', 'wide']);
    expect(index.ruleIdsAt(5, 1)).toEqual(['wide']);
    expect(index.ruleIdsAt(100, 0)).toEqual([]);
    expect(index.ruleIdsAt(0, 6)).toEqual([]);
  });
});

// ---------------------------------------------------------------------------
// §10 property invariants
// ---------------------------------------------------------------------------

describe('CF invariants', () => {
  it('keeps data-bar lengths and the axis within [0, 1]', () => {
    fc.assert(
      fc.property(
        fc.array(fc.double({ min: -1e6, max: 1e6, noNaN: true, noDefaultInfinity: true }), {
          minLength: 1,
          maxLength: 25,
        }),
        fc.integer({ min: 0, max: 24 }),
        (values, pick) => {
          const row = pick % values.length;
          const index = perCellIndex([dataBarRule({ negativeColor: BLUE })], {
            cells: columnSource(values.map((value) => numberValue(value))),
          });
          const bar: CfDataBarOverlay | null = index.evaluateCell(row, 0).dataBar;
          if (bar === null) return false;
          return (
            bar.start >= 0 &&
            bar.start <= bar.end &&
            bar.end <= 1 &&
            bar.axis >= 0 &&
            bar.axis <= 1 &&
            Number.isFinite(bar.start) &&
            Number.isFinite(bar.end)
          );
        },
      ),
      { numRuns: 100 },
    );
  });

  it('keeps colour-scale positions within [0, 1]', () => {
    fc.assert(
      fc.property(
        fc.array(fc.double({ min: -1e6, max: 1e6, noNaN: true, noDefaultInfinity: true }), {
          minLength: 2,
          maxLength: 25,
        }),
        fc.integer({ min: 0, max: 24 }),
        (values, pick) => {
          const row = pick % values.length;
          const rule = {
            ...base('scale', COLUMN_A, 1),
            kind: 'colorScale',
            thresholds: [
              threshold({ type: 'percent', value: 0 }),
              threshold({ type: 'percent', value: 50 }),
              threshold({ type: 'percent', value: 100 }),
            ],
            colors: [RED, YELLOW, GREEN],
          } as ConditionalFormatRule;
          const index = perCellIndex([rule], {
            cells: columnSource(values.map((value) => numberValue(value))),
          });
          const scale: CfColorScaleOverlay | null = index.evaluateCell(row, 0).colorScale;
          if (scale === null) return false;
          return scale.position >= 0 && scale.position <= 1 && /^#[0-9A-F]{6}$/.test(scale.color);
        },
      ),
      { numRuns: 100 },
    );
  });

  it('resolves icon thresholds monotonically', () => {
    fc.assert(
      fc.property(
        fc.array(fc.double({ min: -1000, max: 1000, noNaN: true, noDefaultInfinity: true }), {
          minLength: 5,
          maxLength: 5,
        }),
        fc.array(fc.double({ min: 0, max: 100, noNaN: true, noDefaultInfinity: true }), {
          minLength: 2,
          maxLength: 2,
        }),
        (values, percents) => {
          const sortedPercents = [...percents].sort((left, right) => left - right);
          const rules: ConditionalFormatRule[] = [
            iconSetRule('3Arrows', [
              threshold({ type: 'percent', value: sortedPercents[0]! }),
              threshold({ type: 'percent', value: sortedPercents[1]! }),
            ]),
            iconSetRule('4Arrows', [
              threshold({ type: 'number', value: -10 }),
              threshold({ type: 'number', value: 0 }),
              threshold({ type: 'number', value: 10 }),
            ]),
            iconSetRule('5Quarters', [
              threshold({ type: 'percentile', value: 10 }),
              threshold({ type: 'percentile', value: 30 }),
              threshold({ type: 'percentile', value: 60 }),
              threshold({ type: 'percentile', value: 90 }),
            ]),
          ];
          const index = perCellIndex(rules, {
            cells: columnSource(values.map((value) => numberValue(value))),
          });
          return rules.every((rule) => {
            const statistics = index.statisticsFor(rule.id, 0);
            if (statistics === null || statistics.kind !== 'iconSet') return false;
            return statistics.thresholds.every(
              (value, position) => position === 0 || value >= statistics.thresholds[position - 1]!,
            );
          });
        },
      ),
      { numRuns: 100 },
    );
  });

  it('keeps icon indexes within the set', () => {
    fc.assert(
      fc.property(
        fc.array(fc.double({ min: -1000, max: 1000, noNaN: true, noDefaultInfinity: true }), {
          minLength: 1,
          maxLength: 25,
        }),
        fc.integer({ min: 0, max: 24 }),
        (values, pick) => {
          const row = pick % values.length;
          const index = perCellIndex(
            [
              iconSetRule('5Rating', [
                threshold({ type: 'percent', value: 20 }),
                threshold({ type: 'percent', value: 40 }),
                threshold({ type: 'percent', value: 60 }),
                threshold({ type: 'percent', value: 80 }),
              ]),
            ],
            {
              cells: columnSource(values.map((value) => numberValue(value))),
            },
          );
          const icons: CfIconSetOverlay | null = index.evaluateCell(row, 0).iconSet;
          if (icons === null) return false;
          return icons.iconIndex >= 0 && icons.iconIndex < icons.iconCount;
        },
      ),
      { numRuns: 100 },
    );
  });
});
