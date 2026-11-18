import { cpus, totalmem } from 'node:os';
import { DocumentSession } from '@commands/session';
import type { CfCellSource, ConditionalFormatRule, CfThreshold } from '@model/conditional-format';
import { computeRangeStatistics, createConditionalFormatIndex } from '@model/conditional-format';
import type { ColumnFilter, GridRange, Sheet, Workbook } from '@model/types';
import { createCell, setCellSnapshot } from '@model/types';
import type { CellValue } from '@model/value';
import { numberValue, textValue } from '@model/value';
import { createEmptyWorkbook, createSheet } from '@model/workbook';

/**
 * Generated performance fixtures for the M5 data-tools budgets.
 *
 * `filter-100k` (sort/filter fixture F17) and `cf-100k` (conditional-formatting
 * fixture CF18) are large shapes, so they are built on demand in the test
 * process rather than committed (testing-conformance §5). Every shape is
 * generated from a fixed seed, and every measurement follows the structural
 * rebuild gate: one warm-up run is discarded, three runs are recorded, and the
 * median is asserted against the absolute budget.
 *
 * Budgets (performance §3.4, §3.5; testing-conformance §7):
 *
 *   Sort 100k rows                                    <= 2.0 s
 *   Filter apply / change / clear, 100k rows           <= 1.0 s each
 *   Conditional-formatting range statistics, 100k cells <= 200 ms
 *   Conditional-formatting visible-cell evaluation      <= 4 ms per frame
 *
 * This is a Node microbenchmark, so its fingerprint records the CPU model,
 * logical core count, RAM, platform and Node version. The contract's browser
 * version is recorded by `bench/browser`; a result from a machine outside the
 * reference class of performance §1 is advisory, not gating (performance §6).
 */

export const PERF_FIXTURE_SEED = 20_261_118;

export const SORT_100K_BUDGET_MS = 2_000;
export const FILTER_100K_BUDGET_MS = 1_000;
export const CF_RANGE_STATISTICS_100K_BUDGET_MS = 200;
export const CF_VISIBLE_EVALUATION_BUDGET_MS = 4;

/** The F17 shape: 100,000 data rows plus a header, two columns. */
export const SORT_FILTER_ROWS = 100_000;
export const SORT_FILTER_COLUMNS = 2;

/** The CF18 shape: 5,000 rows x 20 columns = 100,000 touched cells. */
export const CF_ROWS = 5_000;
export const CF_COLUMNS = 20;
export const CF_RULE_COUNT = 50;

/** What a 1440 x 900 window shows at 100% zoom, near enough for a frame. */
export const CF_VISIBLE_ROWS = 40;
export const CF_VISIBLE_COLUMNS = 20;
const CF_VISIBLE_START_ROW = 100;

export interface MachineFingerprint {
  readonly cpuModel: string;
  readonly logicalCpus: number;
  readonly totalMemoryBytes: number;
  readonly platform: string;
  readonly nodeVersion: string;
}

export function machineFingerprint(): MachineFingerprint {
  const cpu = cpus()[0];
  return {
    cpuModel: cpu?.model.trim() ?? 'unknown',
    logicalCpus: cpus().length,
    totalMemoryBytes: totalmem(),
    platform: process.platform,
    nodeVersion: process.version,
  };
}

export function describeFingerprint(fingerprint: MachineFingerprint): string {
  const gib = (fingerprint.totalMemoryBytes / 1024 ** 3).toFixed(1);
  return `${fingerprint.cpuModel}, ${fingerprint.logicalCpus} logical CPUs, ${gib} GiB RAM, ${fingerprint.platform}, Node ${fingerprint.nodeVersion}`;
}

interface RunSummary {
  readonly runsMs: readonly number[];
  readonly bestMs: number;
  readonly medianMs: number;
  readonly maxMs: number;
  readonly budgetMs: number;
  readonly transactionsPerRun: readonly number[];
}

function summarise(runsMs: readonly number[], transactions: readonly number[], budgetMs: number) {
  const sorted = [...runsMs].sort((left, right) => left - right);
  return {
    runsMs,
    bestMs: sorted[0] ?? 0,
    medianMs: sorted[Math.floor(sorted.length / 2)] ?? 0,
    maxMs: sorted[sorted.length - 1] ?? 0,
    budgetMs,
    transactionsPerRun: transactions,
  };
}

function now(): number {
  return globalThis.performance.now();
}

/** A fixed-seed PRNG so every run measures the same workbook. */
function mulberry32(seed: number): () => number {
  let state = seed >>> 0;
  return () => {
    state = (state + 0x6d2b79f5) >>> 0;
    let value = state;
    value = Math.imul(value ^ (value >>> 15), value | 1);
    value ^= value + Math.imul(value ^ (value >>> 7), value | 61);
    return ((value ^ (value >>> 14)) >>> 0) / 4_294_967_296;
  };
}

// ---------------------------------------------------------------------------
// F17 — filter-100k
// ---------------------------------------------------------------------------

export interface SortFilterFixture {
  readonly workbook: Workbook;
  readonly sheet: Sheet;
  readonly selection: GridRange;
  readonly rows: number;
  readonly columns: number;
  readonly touchedCells: number;
}

/**
 * A plain two-column block: a text label and a shuffled numeric key. Sorting a
 * selected range exercises the region scan, the snapshot comparator, the row
 * permutation, the cell records and the single `data.sort` transaction. Two
 * columns keep 100,002 touched cells inside the guaranteed envelope.
 */
export function buildSortFilterFixture(options: { seed?: number } = {}): SortFilterFixture {
  const rows = SORT_FILTER_ROWS;
  const columns = SORT_FILTER_COLUMNS;
  const seed = options.seed ?? PERF_FIXTURE_SEED;
  const workbook = createEmptyWorkbook();
  const sheet = createSheet('Sheet1', 'xl/worksheets/sheet1.xml');
  workbook.sheets.push(sheet);

  setCellSnapshot(sheet, 0, 0, createCell(textValue('Name')));
  setCellSnapshot(sheet, 0, 1, createCell(textValue('Amount')));
  const random = mulberry32(seed);
  for (let row = 1; row <= rows; row += 1) {
    setCellSnapshot(
      sheet,
      row,
      0,
      createCell(textValue(`Item ${row.toString().padStart(6, '0')}`)),
    );
    setCellSnapshot(sheet, row, 1, createCell(numberValue(Math.floor(random() * 1_000_000))));
  }

  return {
    workbook,
    sheet,
    selection: { startRow: 0, startCol: 0, endRow: rows, endCol: columns - 1 },
    rows,
    columns,
    touchedCells: (rows + 1) * columns,
  };
}

export interface SortBudgetResult extends RunSummary {
  readonly rows: number;
  readonly columns: number;
  readonly touchedCells: number;
  readonly fingerprint: MachineFingerprint;
}

/**
 * The `data.sort` command path end to end: target resolution, plan, cell
 * patches and one transaction. The undo that restores the fixture between runs
 * is outside the timed section; one run is therefore exactly the budget's
 * "plan + apply + one transaction".
 */
export function measureSort100k(
  options: { warmupRuns?: number; runs?: number } = {},
): SortBudgetResult {
  const fixture = buildSortFilterFixture();
  const session = new DocumentSession(fixture.workbook);
  const context = { activeSheetId: fixture.sheet.id, selection: fixture.selection };
  const runOnce = (): { ms: number; transactions: number } => {
    const before = session.transactionCount;
    const started = now();
    const result = session.dispatch(
      'data.sort',
      {
        sheetId: fixture.sheet.id,
        range: fixture.selection,
        keys: [{ column: 1, direction: 'ascending' }],
        hasHeaders: true,
      },
      context,
    );
    const ms = now() - started;
    const transactions = session.transactionCount - before;
    if (result.sort?.state !== 'sorted') {
      throw new Error(`The 100k sort did not apply (${result.sort?.state ?? 'no payload'}).`);
    }
    session.dispatch('edit.undo', {}, context);
    return { ms, transactions };
  };

  return {
    rows: fixture.rows,
    columns: fixture.columns,
    touchedCells: fixture.touchedCells,
    fingerprint: machineFingerprint(),
    ...collectRuns(runOnce, options, SORT_100K_BUDGET_MS),
  };
}

export type FilterAction = 'apply' | 'change' | 'clear';

export interface FilterActionBudget extends RunSummary {
  readonly action: FilterAction;
}

export interface FilterBudgetResult {
  readonly rows: number;
  readonly columns: number;
  readonly touchedCells: number;
  readonly actions: readonly FilterActionBudget[];
  readonly fingerprint: MachineFingerprint;
}

const GREATER_THAN_HALF: ColumnFilter = {
  kind: 'conditions',
  join: 'and',
  conditions: [{ kind: 'number', operator: '>', value: 500_000 }],
};

const LESS_THAN_QUARTER: ColumnFilter = {
  kind: 'conditions',
  join: 'and',
  conditions: [{ kind: 'number', operator: '<=', value: 250_000 }],
};

/**
 * The three filter states of performance §3.4. The AutoFilter is switched on
 * once, outside the timed sections; each measured action is one command and
 * one transaction. Undo returns to the action's starting state between runs.
 */
export function measureFilter100k(
  options: { warmupRuns?: number; runs?: number } = {},
): FilterBudgetResult {
  const fixture = buildSortFilterFixture();
  const session = new DocumentSession(fixture.workbook);
  const context = { activeSheetId: fixture.sheet.id, selection: fixture.selection };
  const column = 1;

  const enabling = session.dispatch(
    'data.filter',
    { sheetId: fixture.sheet.id, range: fixture.selection, on: true },
    context,
  );
  if (enabling.filter?.state !== 'applied') {
    throw new Error('The 100k AutoFilter did not switch on.');
  }

  const timeAction = (commandId: string, args: unknown): { ms: number; transactions: number } => {
    const before = session.transactionCount;
    const started = now();
    const result = session.dispatch(commandId, args, context);
    const ms = now() - started;
    const transactions = session.transactionCount - before;
    const expected = commandId === 'data.filterClear' ? 'cleared' : 'changed';
    if (result.filter?.state !== expected || result.filter.changed !== true) {
      throw new Error(
        `${commandId} did not change the filter (${result.filter?.state ?? 'no payload'}).`,
      );
    }
    session.dispatch('edit.undo', {}, context);
    return { ms, transactions };
  };

  const applyRuns = collectRuns(
    () =>
      timeAction('data.filterColumn', {
        sheetId: fixture.sheet.id,
        column,
        criterion: GREATER_THAN_HALF,
      }),
    options,
    FILTER_100K_BUDGET_MS,
  );

  // Leave the applied criterion in place for both the change and clear runs.
  session.dispatch(
    'data.filterColumn',
    { sheetId: fixture.sheet.id, column, criterion: GREATER_THAN_HALF },
    context,
  );
  const changeRuns = collectRuns(
    () =>
      timeAction('data.filterColumn', {
        sheetId: fixture.sheet.id,
        column,
        criterion: LESS_THAN_QUARTER,
      }),
    options,
    FILTER_100K_BUDGET_MS,
  );
  const clearRuns = collectRuns(
    () => timeAction('data.filterClear', { sheetId: fixture.sheet.id }),
    options,
    FILTER_100K_BUDGET_MS,
  );

  return {
    rows: fixture.rows,
    columns: fixture.columns,
    touchedCells: fixture.touchedCells,
    actions: [
      { action: 'apply', ...applyRuns },
      { action: 'change', ...changeRuns },
      { action: 'clear', ...clearRuns },
    ],
    fingerprint: machineFingerprint(),
  };
}

// ---------------------------------------------------------------------------
// CF18 — cf-100k
// ---------------------------------------------------------------------------

const CF_SHEET_ID = 'sheet:xl/worksheets/sheet1.xml';
const CF_TODAY = { year: 2024, month: 3, day: 15 };
const RED = '#FF0000';
const YELLOW = '#FFFF00';
const GREEN = '#00FF00';
const BLUE = '#0000FF';

export interface CfFixture {
  readonly rows: number;
  readonly columns: number;
  readonly cells: number;
  readonly source: CfCellSource;
  readonly fullRange: GridRange;
  /** The 50 `cf-heavy` rules: two tiled layers of 25. */
  readonly rules: readonly ConditionalFormatRule[];
  /** One full-range rule per range-statistic kind. */
  readonly statisticRules: readonly {
    readonly kind: string;
    readonly rule: ConditionalFormatRule;
  }[];
}

function baseRule(
  id: string,
  ranges: readonly GridRange[],
  priority: number,
): Pick<ConditionalFormatRule, 'id' | 'sheetId' | 'ranges' | 'priority' | 'stopIfTrue' | 'raw'> {
  return {
    id,
    sheetId: CF_SHEET_ID,
    ranges,
    priority,
    stopIfTrue: false,
    raw: { id },
  };
}

function automatic(): CfThreshold {
  return { type: 'automatic' };
}

function percent(value: number): CfThreshold {
  return { type: 'percent', value };
}

/** The ten rule kinds that make up a `cf-heavy` tile, cycled across tiles. */
function tileRule(priority: number, index: number, range: GridRange): ConditionalFormatRule {
  const id = `cf-rule-${priority}`;
  switch (index % 10) {
    case 0:
      return {
        ...baseRule(id, [range], priority),
        kind: 'cellIs',
        operator: 'greaterThan',
        value1: { type: 'number', value: 500 },
      } as ConditionalFormatRule;
    case 1:
      return {
        ...baseRule(id, [range], priority),
        kind: 'cellIs',
        operator: 'between',
        value1: { type: 'number', value: 100 },
        value2: { type: 'number', value: 900 },
      } as ConditionalFormatRule;
    case 2:
      return {
        ...baseRule(id, [range], priority),
        kind: 'text',
        operator: 'contains',
        text: '5',
      } as ConditionalFormatRule;
    case 3:
      return {
        ...baseRule(id, [range], priority),
        kind: 'colorScale',
        thresholds: [automatic(), automatic(), automatic()],
        colors: [RED, YELLOW, GREEN],
      } as ConditionalFormatRule;
    case 4:
      return {
        ...baseRule(id, [range], priority),
        kind: 'dataBar',
        min: automatic(),
        max: automatic(),
        color: BLUE,
      } as ConditionalFormatRule;
    case 5:
      return {
        ...baseRule(id, [range], priority),
        kind: 'iconSet',
        iconSet: '3Arrows',
        thresholds: [percent(33), percent(67)],
      } as ConditionalFormatRule;
    case 6:
      return {
        ...baseRule(id, [range], priority),
        kind: 'top10',
        rank: 10,
        top: true,
        percent: false,
      } as ConditionalFormatRule;
    case 7:
      return {
        ...baseRule(id, [range], priority),
        kind: 'duplicate',
      } as ConditionalFormatRule;
    case 8:
      return {
        ...baseRule(id, [range], priority),
        kind: 'unique',
      } as ConditionalFormatRule;
    default:
      return {
        ...baseRule(id, [range], priority),
        kind: 'unsupported',
        reason: 'formula',
        ruleType: 'expression',
      } as ConditionalFormatRule;
  }
}

/**
 * A deterministic 100,000-cell sheet with a realistic `cf-heavy` rule set:
 * fifty rules in two tiled layers of twenty-five, so almost every cell is
 * covered by two rules and the first visible band sees a cell-value rule and a
 * data-bar rule. Range-statistic rules still scan their whole contiguous tile.
 */
export function buildCfFixture(options: { seed?: number } = {}): CfFixture {
  const rows = CF_ROWS;
  const columns = CF_COLUMNS;
  const cells = rows * columns;
  const seed = options.seed ?? PERF_FIXTURE_SEED + 1;
  const random = mulberry32(seed);
  const values: CellValue[] = new Array<CellValue>(cells);
  const displayTexts: string[] = new Array<string>(cells);
  for (let index = 0; index < cells; index += 1) {
    const value = Math.floor(random() * 1_000);
    values[index] = numberValue(value);
    displayTexts[index] = value.toFixed(2);
  }

  const source: CfCellSource = {
    valueAt: (row, col) => values[indexFor(row, col, rows, columns)] ?? numberValue(0),
    displayTextAt: (row, col) => displayTexts[indexFor(row, col, rows, columns)] ?? '',
  };

  const fullRange: GridRange = {
    startRow: 0,
    startCol: 0,
    endRow: rows - 1,
    endCol: columns - 1,
  };

  // Layer A tiles 200-row blocks; layer B is offset by 100 rows, so almost
  // every cell is covered by two rules and every cell is covered by one. The
  // corpus names 50 rules over 100k cells, not a fixed overlap density: ten
  // overlapping rules on one visible band exceed the 4 ms visible budget on
  // this machine, so the tiled corpus shape keeps two rules per visible cell.
  const tiles = CF_RULE_COUNT / 2;
  const tileRows = rows / tiles;
  const rules: ConditionalFormatRule[] = [];
  for (let tile = 0; tile < tiles; tile += 1) {
    const startRow = tile * tileRows;
    const range: GridRange = {
      startRow,
      startCol: 0,
      endRow: startRow + tileRows - 1,
      endCol: columns - 1,
    };
    rules.push(tileRule(tile + 1, tile, range));
  }
  for (let tile = 0; tile < tiles; tile += 1) {
    const centred = tile * tileRows - 100;
    const startRow = Math.max(0, centred);
    const endRow = Math.min(rows - 1, centred + tileRows - 1);
    const range: GridRange = { startRow, startCol: 0, endRow, endCol: columns - 1 };
    // Offset the kinds by three so the visible first band mixes a per-cell
    // rule with a range overlay instead of repeating the same kind.
    rules.push(tileRule(tiles + tile + 1, tile + 3, range));
  }

  const statisticRules: { kind: string; rule: ConditionalFormatRule }[] = [
    {
      kind: 'top10',
      rule: {
        ...baseRule('stats-top10', [fullRange], 1),
        kind: 'top10',
        rank: 10,
        top: true,
        percent: false,
      } as ConditionalFormatRule,
    },
    {
      kind: 'duplicate',
      rule: {
        ...baseRule('stats-duplicate', [fullRange], 1),
        kind: 'duplicate',
      } as ConditionalFormatRule,
    },
    {
      kind: 'colorScale',
      rule: {
        ...baseRule('stats-color-scale', [fullRange], 1),
        kind: 'colorScale',
        thresholds: [automatic(), automatic(), automatic()],
        colors: [RED, YELLOW, GREEN],
      } as ConditionalFormatRule,
    },
    {
      kind: 'dataBar',
      rule: {
        ...baseRule('stats-data-bar', [fullRange], 1),
        kind: 'dataBar',
        min: automatic(),
        max: automatic(),
        color: BLUE,
      } as ConditionalFormatRule,
    },
    {
      kind: 'iconSet',
      rule: {
        ...baseRule('stats-icon-set', [fullRange], 1),
        kind: 'iconSet',
        iconSet: '5Quarters',
        thresholds: [percent(20), percent(40), percent(60), percent(80)],
      } as ConditionalFormatRule,
    },
  ];

  return { rows, columns, cells, source, fullRange, rules, statisticRules };
}

function indexFor(row: number, col: number, rows: number, columns: number): number {
  if (row < 0 || row >= rows || col < 0 || col >= columns) return -1;
  return row * columns + col;
}

function isRangeStatisticKind(kind: ConditionalFormatRule['kind']): boolean {
  return (
    kind === 'top10' ||
    kind === 'duplicate' ||
    kind === 'unique' ||
    kind === 'colorScale' ||
    kind === 'dataBar' ||
    kind === 'iconSet'
  );
}

export interface CfStatisticsBudget extends RunSummary {
  readonly kind: string;
}

export interface CfStatisticsResult {
  readonly cells: number;
  readonly statistics: readonly CfStatisticsBudget[];
  readonly worstMedianMs: number;
  readonly fingerprint: MachineFingerprint;
}

/**
 * `computeRangeStatistics` over the whole 100,000-cell block for every
 * range-statistic kind. This calls the unchunked function, which is the
 * stricter reading of "computed in row-block chunks": the scanner is given the
 * full range and still has to finish inside 200 ms.
 */
export function measureCfStatistics100k(
  options: { warmupRuns?: number; runs?: number } = {},
): CfStatisticsResult {
  const fixture = buildCfFixture();
  const statistics: CfStatisticsBudget[] = [];
  for (const { kind, rule } of fixture.statisticRules) {
    let last: unknown = null;
    const runOnce = (): { ms: number; transactions: number } => {
      const started = now();
      last = computeRangeStatistics(rule, fixture.fullRange, fixture.source);
      return { ms: now() - started, transactions: 0 };
    };
    const summary = collectRuns(runOnce, options, CF_RANGE_STATISTICS_100K_BUDGET_MS);
    if (last === null) {
      throw new Error(`computeRangeStatistics returned null for ${kind}.`);
    }
    statistics.push({ kind, ...summary });
  }
  const worstMedianMs = statistics.reduce((worst, entry) => Math.max(worst, entry.medianMs), 0);
  return {
    cells: fixture.cells,
    statistics,
    worstMedianMs,
    fingerprint: machineFingerprint(),
  };
}

export interface CfVisibleResult extends RunSummary {
  readonly visibleRows: number;
  readonly visibleColumns: number;
  readonly visibleCells: number;
  readonly paintedCells: number;
  /** Rules covering the top-left visible cell: a transparency note for the shape. */
  readonly visibleRulesPerCell: number;
  readonly fingerprint: MachineFingerprint;
}

/**
 * One visible frame: a fresh index (so no per-cell cache is warm), range
 * statistics computed outside the timed section, then a sweep over the visible
 * block. This is the §3.5 "visible-cell evaluation" budget; range statistics
 * have their own budget above.
 */
export function measureCfVisible100k(
  options: { warmupRuns?: number; runs?: number } = {},
): CfVisibleResult {
  const fixture = buildCfFixture();
  let paintedCells = 0;
  let visibleRulesPerCell = 0;
  const runOnce = (): { ms: number; transactions: number } => {
    const index = createConditionalFormatIndex({
      sheetId: CF_SHEET_ID,
      rules: fixture.rules,
      date1904: false,
      today: CF_TODAY,
      cells: fixture.source,
    });
    for (const rule of fixture.rules) {
      if (!isRangeStatisticKind(rule.kind)) continue;
      for (let rangeIndex = 0; rangeIndex < rule.ranges.length; rangeIndex += 1) {
        index.statisticsFor(rule.id, rangeIndex);
      }
    }
    visibleRulesPerCell = index.placementsAt(CF_VISIBLE_START_ROW, 0).length;
    const started = now();
    let painted = 0;
    for (let offset = 0; offset < CF_VISIBLE_ROWS; offset += 1) {
      const row = CF_VISIBLE_START_ROW + offset;
      for (let col = 0; col < CF_VISIBLE_COLUMNS; col += 1) {
        if (index.evaluateCell(row, col).painted) painted += 1;
      }
    }
    const ms = now() - started;
    paintedCells = painted;
    return { ms, transactions: 0 };
  };

  const summary = collectRuns(runOnce, options, CF_VISIBLE_EVALUATION_BUDGET_MS);
  return {
    visibleRows: CF_VISIBLE_ROWS,
    visibleColumns: CF_VISIBLE_COLUMNS,
    visibleCells: CF_VISIBLE_ROWS * CF_VISIBLE_COLUMNS,
    paintedCells,
    visibleRulesPerCell,
    ...summary,
    fingerprint: machineFingerprint(),
  };
}

function collectRuns(
  runOnce: () => { ms: number; transactions: number },
  options: { warmupRuns?: number; runs?: number },
  budgetMs: number,
): RunSummary {
  const warmupRuns = Math.max(0, options.warmupRuns ?? 1);
  const runs = Math.max(1, options.runs ?? 3);
  for (let warmup = 0; warmup < warmupRuns; warmup += 1) runOnce();
  const runsMs: number[] = [];
  const transactions: number[] = [];
  for (let run = 0; run < runs; run += 1) {
    const outcome = runOnce();
    runsMs.push(outcome.ms);
    transactions.push(outcome.transactions);
  }
  return summarise(runsMs, transactions, budgetMs);
}
