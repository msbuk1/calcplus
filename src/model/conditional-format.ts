/**
 * Conditional-formatting evaluation (conditional formatting §2, §3 and §6).
 *
 * The module is a pure overlay engine: it reads typed cell values and caller
 * supplied display text and returns descriptors the renderer paints on top of
 * manual formatting. It never mutates a cell, its style or its value, and it
 * never reads a clock — "today" is an input.
 *
 * Rendering (§5), XML round-trip (§7) and the compatibility report belong to
 * other layers. Unsupported rules keep their `raw` payload here so those layers
 * can report them; evaluation only records them as not painted.
 */

import type { GridRange, SheetId, BorderWeight, UnderlineStyle } from './types';
import { rangeContains } from './types';
import type { CellValue } from './value';
import { BLANK, formatCellValue } from './value';
import { datePartsToSerial } from './date-system';

// ---------------------------------------------------------------------------
// Rule model (§2)
// ---------------------------------------------------------------------------

export type CfKind =
  | 'cellIs'
  | 'text'
  | 'timePeriod'
  | 'top10'
  | 'duplicate'
  | 'unique'
  | 'colorScale'
  | 'dataBar'
  | 'iconSet'
  | 'unsupported';

export type CfCellOperator =
  | 'lessThan'
  | 'lessThanOrEqual'
  | 'equal'
  | 'notEqual'
  | 'greaterThanOrEqual'
  | 'greaterThan'
  | 'between'
  | 'notBetween';

export type CfTextOperator = 'contains' | 'notContains' | 'beginsWith' | 'endsWith';

export type CfDateOperator = 'equal' | 'notEqual' | 'before' | 'after' | 'between' | 'notBetween';

/** Excel's ten relative date presets. There is deliberately no year-based preset. */
export type CfTimePeriod =
  | 'yesterday'
  | 'today'
  | 'tomorrow'
  | 'last7Days'
  | 'lastWeek'
  | 'thisWeek'
  | 'nextWeek'
  | 'lastMonth'
  | 'thisMonth'
  | 'nextMonth';

/** A literal comparison operand. Formula operands make the rule unsupported. */
export type CfOperand =
  | { readonly type: 'number'; readonly value: number }
  | { readonly type: 'text'; readonly value: string }
  | { readonly type: 'boolean'; readonly value: boolean };

/** A colour-scale, data-bar or icon-set threshold point. */
export type CfThreshold =
  | { readonly type: 'automatic' }
  | { readonly type: 'number'; readonly value: number }
  | { readonly type: 'percent'; readonly value: number }
  | { readonly type: 'percentile'; readonly value: number };

/** All Excel built-in icon sets in v1 scope. */
export type CfIconSetId =
  | '3Arrows'
  | '3ArrowsGray'
  | '3Flags'
  | '3TrafficLights1'
  | '3TrafficLights2'
  | '3Signs'
  | '3Symbols'
  | '3Symbols2'
  | '4Arrows'
  | '4ArrowsGray'
  | '4RedToBlack'
  | '4Rating'
  | '4TrafficLights'
  | '5Arrows'
  | '5ArrowsGray'
  | '5Rating'
  | '5Quarters';

export interface CfIconSetDefinition {
  readonly id: CfIconSetId;
  readonly iconCount: 3 | 4 | 5;
}

/** The icon count for every built-in set, for the renderer and tests. */
export const CF_ICON_SETS: readonly CfIconSetDefinition[] = Object.freeze([
  { id: '3Arrows', iconCount: 3 },
  { id: '3ArrowsGray', iconCount: 3 },
  { id: '3Flags', iconCount: 3 },
  { id: '3TrafficLights1', iconCount: 3 },
  { id: '3TrafficLights2', iconCount: 3 },
  { id: '3Signs', iconCount: 3 },
  { id: '3Symbols', iconCount: 3 },
  { id: '3Symbols2', iconCount: 3 },
  { id: '4Arrows', iconCount: 4 },
  { id: '4ArrowsGray', iconCount: 4 },
  { id: '4RedToBlack', iconCount: 4 },
  { id: '4Rating', iconCount: 4 },
  { id: '4TrafficLights', iconCount: 4 },
  { id: '5Arrows', iconCount: 5 },
  { id: '5ArrowsGray', iconCount: 5 },
  { id: '5Rating', iconCount: 5 },
  { id: '5Quarters', iconCount: 5 },
]);

export type CfUnsupportedReason =
  | 'formula'
  | 'aboveAverage'
  | 'blankOrError'
  | 'formulaThreshold'
  | 'unknownOperator'
  | 'unknownRuleType';

interface CfRuleBase {
  readonly id: string;
  readonly sheetId: SheetId;
  /** Contiguous ranges from `sqref`; each is its own ranking domain. */
  readonly ranges: readonly GridRange[];
  /** 1 is highest. Missing priorities sort last, in document order (§2). */
  readonly priority?: number;
  readonly stopIfTrue: boolean;
  readonly dxfId?: number;
  /** Original XML and attributes, preserved untouched for reporting and save. */
  readonly raw: unknown;
}

export type ConditionalFormatRule = CfRuleBase &
  (
    | {
        readonly kind: 'cellIs';
        readonly operator: CfCellOperator;
        readonly value1: CfOperand;
        readonly value2?: CfOperand;
      }
    | {
        readonly kind: 'text';
        readonly operator: CfTextOperator;
        readonly text: string;
      }
    | {
        readonly kind: 'timePeriod';
        readonly mode: 'period';
        readonly period: CfTimePeriod;
      }
    | {
        readonly kind: 'timePeriod';
        readonly mode: 'comparison';
        readonly operator: CfDateOperator;
        readonly value1: number;
        readonly value2?: number;
      }
    | {
        readonly kind: 'top10';
        /** N, default 10. */
        readonly rank?: number;
        readonly top: boolean;
        readonly percent: boolean;
      }
    | { readonly kind: 'duplicate' }
    | { readonly kind: 'unique' }
    | {
        readonly kind: 'colorScale';
        /** Two or three points, in ascending order. */
        readonly thresholds: readonly CfThreshold[];
        /** CSS hex colours, one per threshold. */
        readonly colors: readonly string[];
      }
    | {
        readonly kind: 'dataBar';
        readonly min: CfThreshold;
        readonly max: CfThreshold;
        /** Positive bar colour as a CSS hex string. */
        readonly color: string;
        readonly negativeColor?: string;
        readonly fill?: 'gradient' | 'solid';
        readonly showValue?: boolean;
        readonly border?: boolean;
        readonly direction?: 'leftToRight' | 'rightToLeft';
        readonly axisColor?: string;
      }
    | {
        readonly kind: 'iconSet';
        readonly iconSet: CfIconSetId;
        /** Ascending thresholds; the icon count minus one entries. */
        readonly thresholds: readonly CfThreshold[];
        readonly reverse?: boolean;
        readonly showValue?: boolean;
        readonly iconOnly?: boolean;
      }
    | {
        readonly kind: 'unsupported';
        readonly reason: CfUnsupportedReason;
        readonly ruleType?: string;
      }
  );

// ---------------------------------------------------------------------------
// Differential format and overlay descriptors
// ---------------------------------------------------------------------------

/**
 * A dxf holds only the properties its rule sets. Each patch field is optional,
 * so a renderer spreads it over the manual style and the manual value shows
 * through where the dxf is silent (§3.1).
 */
export interface CfFillPatch {
  readonly pattern?: string;
  readonly color?: string | null;
  readonly backgroundColor?: string | null;
}

export interface CfFontPatch {
  readonly name?: string;
  readonly size?: number;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: boolean;
  readonly underlineStyle?: UnderlineStyle;
  readonly strike?: boolean;
  readonly color?: string | null;
}

export interface CfBorderEdgePatch {
  readonly style?: string;
  readonly color?: string | null;
  readonly weight?: BorderWeight | null;
}

export interface CfBorderPatch {
  readonly left?: CfBorderEdgePatch;
  readonly right?: CfBorderEdgePatch;
  readonly top?: CfBorderEdgePatch;
  readonly bottom?: CfBorderEdgePatch;
  readonly diagonal?: CfBorderEdgePatch;
  readonly diagonalUp?: boolean;
  readonly diagonalDown?: boolean;
}

export interface CfNumberFormatPatch {
  readonly numFmtId?: number;
  readonly code?: string;
}

/** The four merge properties. The first true rule that sets one wins it. */
export interface CfDifferentialFormat {
  readonly fill?: CfFillPatch;
  readonly font?: CfFontPatch;
  readonly border?: CfBorderPatch;
  readonly numberFormat?: CfNumberFormatPatch;
  /** Original `<dxf>` XML, untouched. */
  readonly rawXml?: string;
}

export type CfDxfResolver = (dxfId: number) => CfDifferentialFormat | null;

/** The merged CF contribution for one cell. `null` means the manual value shows. */
export interface CfOverlay {
  readonly fill: CfFillPatch | null;
  readonly font: CfFontPatch | null;
  readonly border: CfBorderPatch | null;
  readonly numberFormat: CfNumberFormatPatch | null;
}

export interface CfColorScalePoint {
  readonly value: number;
  /** Normalised `#RRGGBB`. */
  readonly color: string;
}

export interface CfColorScaleOverlay {
  readonly kind: 'colorScale';
  readonly ruleId: string;
  readonly value: number;
  /** Normalised position across the whole scale, 0 at the first point, 1 at the last. */
  readonly position: number;
  /** Interpolated `#RRGGBB`. */
  readonly color: string;
  readonly points: readonly CfColorScalePoint[];
}

/** Geometry inputs for the renderer, never pixels (§3.3). */
export interface CfDataBarOverlay {
  readonly kind: 'dataBar';
  readonly ruleId: string;
  readonly value: number;
  readonly min: number;
  readonly max: number;
  /** Left edge as a fraction of the cell/merged width, 0-1. */
  readonly start: number;
  /** Right edge as a fraction of the cell/merged width, 0-1. */
  readonly end: number;
  /** The zero axis as a fraction of the cell width, 0-1. */
  readonly axis: number;
  readonly negative: boolean;
  /** The colour the bar paints with (positive or negative colour). */
  readonly color: string;
  readonly fill: 'gradient' | 'solid';
  readonly direction: 'leftToRight' | 'rightToLeft';
  readonly border: boolean;
  readonly showValue: boolean;
  readonly axisColor: string | null;
}

export interface CfIconSetOverlay {
  readonly kind: 'iconSet';
  readonly ruleId: string;
  readonly iconSet: CfIconSetId;
  readonly iconCount: number;
  /** Chosen icon index, after applying `reverse`. */
  readonly iconIndex: number;
  readonly value: number;
  /** Resolved ascending thresholds, before applying `reverse`. */
  readonly thresholds: readonly number[];
  readonly showValue: boolean;
  readonly iconOnly: boolean;
}

export interface CfCellEvaluation {
  /** True rule ids, in priority order. */
  readonly ruleIds: readonly string[];
  /** The rule id whose stop-if-true halted lower rules, or null. */
  readonly stoppedBy: string | null;
  readonly overlay: CfOverlay;
  readonly colorScale: CfColorScaleOverlay | null;
  readonly dataBar: CfDataBarOverlay | null;
  readonly iconSet: CfIconSetOverlay | null;
  /** Unsupported rules that cover this cell; preserved and reported, never painted. */
  readonly unsupportedRuleIds: readonly string[];
  /** True when at least one property or visual overlay is set. */
  readonly painted: boolean;
}

// ---------------------------------------------------------------------------
// Range statistics
// ---------------------------------------------------------------------------

export interface CfTop10Statistics {
  readonly kind: 'top10';
  readonly count: number;
  readonly rank: number;
  readonly percent: boolean;
  readonly top: boolean;
  /** The Nth value, or null when no cell can qualify. Ties compare equal. */
  readonly threshold: number | null;
}

export interface CfDuplicateStatistics {
  readonly kind: 'duplicate' | 'unique';
  readonly counts: ReadonlyMap<string, number>;
}

export interface CfColorScaleStatistics {
  readonly kind: 'colorScale';
  readonly points: readonly CfColorScalePoint[];
}

export interface CfDataBarStatistics {
  readonly kind: 'dataBar';
  readonly min: number;
  readonly max: number;
}

export interface CfIconSetStatistics {
  readonly kind: 'iconSet';
  readonly thresholds: readonly number[];
  readonly min: number;
  readonly max: number;
}

export type CfRangeStatistics =
  | CfTop10Statistics
  | CfDuplicateStatistics
  | CfColorScaleStatistics
  | CfDataBarStatistics
  | CfIconSetStatistics;

// ---------------------------------------------------------------------------
// Evaluation inputs
// ---------------------------------------------------------------------------

export interface CfCalendarDate {
  readonly year: number;
  readonly month: number;
  readonly day: number;
}

export interface CfCellSubject {
  /** The cell's current value: the saved result for a formula cell. */
  readonly value: CellValue;
  /** Formatted display text; the haystack for text rules. Defaults to General. */
  readonly displayText?: string;
}

/** Read access to a sheet's values. Kept as a port so this module stays pure. */
export interface CfCellSource {
  valueAt(row: number, col: number): CellValue;
  /** The formatted display text, for text rules. */
  displayTextAt(row: number, col: number): string;
}

export interface ConditionalFormatIndexOptions {
  readonly sheetId: SheetId;
  readonly rules: readonly ConditionalFormatRule[];
  readonly date1904: boolean;
  /** The local calendar date; no clock read happens in this module. */
  readonly today: CfCalendarDate;
  readonly cells?: CfCellSource;
  readonly dxfFor?: CfDxfResolver;
}

const EMPTY_CELL_SOURCE: CfCellSource = {
  valueAt: () => BLANK,
  displayTextAt: () => '',
};

const EMPTY_OVERLAY: CfOverlay = Object.freeze({
  fill: null,
  font: null,
  border: null,
  numberFormat: null,
});

const NOTHING: CfCellEvaluation = Object.freeze({
  ruleIds: [],
  stoppedBy: null,
  overlay: EMPTY_OVERLAY,
  colorScale: null,
  dataBar: null,
  iconSet: null,
  unsupportedRuleIds: [],
  painted: false,
});

// ---------------------------------------------------------------------------
// Priority order and range index (§2, §6)
// ---------------------------------------------------------------------------

/**
 * Sort rules by Rule priority. Valid numeric priorities sort ascending; rules
 * without a finite priority follow in document order, and equal priorities
 * keep document order (stable). The resolution of malformed priorities is
 * documented in the module tests.
 */
export function orderRulesByPriority(
  rules: readonly ConditionalFormatRule[],
): ConditionalFormatRule[] {
  const decorated = rules.map((rule, index) => ({
    rule,
    index,
    priority: rule.priority !== undefined && Number.isFinite(rule.priority) ? rule.priority : null,
  }));
  decorated.sort((left, right) => {
    if (left.priority === null || right.priority === null) {
      if (left.priority === null && right.priority === null) return left.index - right.index;
      return left.priority === null ? 1 : -1;
    }
    if (left.priority !== right.priority) return left.priority - right.priority;
    return left.index - right.index;
  });
  return decorated.map((entry) => entry.rule);
}

/** One rule attached to one contiguous range, as the range index reports it. */
export interface CfRulePlacement {
  readonly rule: ConditionalFormatRule;
  readonly range: GridRange;
  readonly rangeIndex: number;
  /** Position in priority order, used to keep the first placement per rule. */
  readonly order: number;
}

interface CfRowBand {
  readonly startRow: number;
  readonly placements: readonly CfRulePlacement[];
}

function buildRowBands(orderedRules: readonly ConditionalFormatRule[]): {
  boundaries: readonly number[];
  bands: readonly CfRowBand[];
} {
  const boundarySet = new Set<number>();
  const placements: CfRulePlacement[] = [];
  orderedRules.forEach((rule, order) => {
    rule.ranges.forEach((range, rangeIndex) => {
      boundarySet.add(range.startRow);
      boundarySet.add(range.endRow + 1);
      placements.push({ rule, range, rangeIndex, order });
    });
  });
  const boundaries = [...boundarySet].sort((left, right) => left - right);
  const bands: CfRowBand[] = [];
  for (let index = 0; index < boundaries.length - 1; index += 1) {
    const startRow = boundaries[index]!;
    bands.push({
      startRow,
      placements: placements
        .filter(
          (placement) => placement.range.startRow <= startRow && placement.range.endRow >= startRow,
        )
        .sort((left, right) => left.order - right.order),
    });
  }
  return { boundaries, bands };
}

function bandForRow(
  boundaries: readonly number[],
  bands: readonly CfRowBand[],
  row: number,
): CfRowBand | null {
  if (bands.length === 0 || row < boundaries[0]!) return null;
  let low = 0;
  let high = bands.length - 1;
  let found = -1;
  while (low <= high) {
    const middle = (low + high) >> 1;
    if (bands[middle]!.startRow <= row) {
      found = middle;
      low = middle + 1;
    } else {
      high = middle - 1;
    }
  }
  if (found < 0) return null;
  if (row >= boundaries[found + 1]!) return null;
  return bands[found]!;
}

// ---------------------------------------------------------------------------
// Typed comparisons and wildcards (§3.2, §3.3)
// ---------------------------------------------------------------------------

function valueKey(value: CellValue): string {
  switch (value.type) {
    case 'number':
      return `n:${value.value}`;
    case 'text':
      return `t:${value.value}`;
    case 'boolean':
      return `b:${value.value ? 1 : 0}`;
    case 'error':
      return `e:${value.value}`;
    case 'blank':
      return 'blank';
  }
}

/**
 * Excel's typed order: numbers < text < FALSE < TRUE. A blank coerces to 0
 * against numeric and boolean operands and to the empty string against text.
 * Returns null when the cell is an error, which no comparison matches.
 */
function compareCellToOperand(cell: CellValue, operand: CfOperand): number | null {
  if (cell.type === 'error') return null;
  switch (operand.type) {
    case 'number': {
      if (cell.type === 'number') return compareNumbers(cell.value, operand.value);
      if (cell.type === 'blank') return compareNumbers(0, operand.value);
      return 1; // text and booleans sort after numbers
    }
    case 'text': {
      if (cell.type === 'text') return compareText(cell.value, operand.value);
      if (cell.type === 'blank') return compareText('', operand.value);
      if (cell.type === 'number') return -1;
      return 1; // booleans sort after text
    }
    case 'boolean': {
      if (cell.type === 'boolean') {
        if (cell.value === operand.value) return 0;
        return cell.value ? 1 : -1;
      }
      if (cell.type === 'blank') return -1; // blank is 0, before FALSE
      if (cell.type === 'number') return -1;
      return -1; // text sorts before booleans
    }
  }
}

function compareNumbers(left: number, right: number): number {
  return left === right ? 0 : left < right ? -1 : 1;
}

function compareText(left: string, right: string): number {
  const a = left.toLowerCase();
  const b = right.toLowerCase();
  return a === b ? 0 : a < b ? -1 : 1;
}

function escapeRegExp(character: string): string {
  return /[\\^$.*+?()[\]{}|/]/.test(character) ? `\\${character}` : character;
}

/** Excel wildcards: `*`, `?` and `~` as the escape character. */
function wildcardRegExp(pattern: string, anchors: 'none' | 'start' | 'end'): RegExp {
  let source = '';
  for (let index = 0; index < pattern.length; index += 1) {
    const character = pattern[index]!;
    if (character === '~') {
      const escaped = pattern[index + 1];
      if (escaped === undefined) {
        source += '~';
      } else {
        source += escapeRegExp(escaped);
        index += 1;
      }
    } else if (character === '*') {
      source += '.*';
    } else if (character === '?') {
      source += '.';
    } else {
      source += escapeRegExp(character);
    }
  }
  const prefix = anchors === 'start' ? '^' : '';
  const suffix = anchors === 'end' ? '$' : '';
  return new RegExp(`${prefix}${source}${suffix}`);
}

function matchesTextRule(
  operator: CfTextOperator,
  haystack: string,
  pattern: string,
  blank: boolean,
): boolean {
  if (blank && pattern !== '') return false;
  const lowerHaystack = haystack.toLowerCase();
  const lowerPattern = pattern.toLowerCase();
  switch (operator) {
    case 'contains':
      return wildcardRegExp(lowerPattern, 'none').test(lowerHaystack);
    case 'notContains':
      return !wildcardRegExp(lowerPattern, 'none').test(lowerHaystack);
    case 'beginsWith':
      return wildcardRegExp(lowerPattern, 'start').test(lowerHaystack);
    case 'endsWith':
      return wildcardRegExp(lowerPattern, 'end').test(lowerHaystack);
  }
}

// ---------------------------------------------------------------------------
// Date periods (§3.3)
// ---------------------------------------------------------------------------

interface SerialInterval {
  readonly start: number;
  /** Exclusive. */
  readonly end: number;
}

function addDays(date: CfCalendarDate, days: number): CfCalendarDate {
  const shifted = new Date(Date.UTC(date.year, date.month - 1, date.day + days));
  return {
    year: shifted.getUTCFullYear(),
    month: shifted.getUTCMonth() + 1,
    day: shifted.getUTCDate(),
  };
}

function startOfMonth(date: CfCalendarDate): CfCalendarDate {
  return { year: date.year, month: date.month, day: 1 };
}

function addMonths(date: CfCalendarDate, months: number): CfCalendarDate {
  const shifted = new Date(Date.UTC(date.year, date.month - 1 + months, 1));
  return {
    year: shifted.getUTCFullYear(),
    month: shifted.getUTCMonth() + 1,
    day: 1,
  };
}

function weekdayIndex(date: CfCalendarDate): number {
  return new Date(Date.UTC(date.year, date.month - 1, date.day)).getUTCDay();
}

function serialOf(date: CfCalendarDate, date1904: boolean): number {
  return datePartsToSerial(date.year, date.month, date.day, date1904);
}

/** Excel's relative-date window. Weeks start on Sunday, as in Excel. */
function periodInterval(
  period: CfTimePeriod,
  today: CfCalendarDate,
  date1904: boolean,
): SerialInterval {
  const todaySerial = serialOf(today, date1904);
  switch (period) {
    case 'yesterday': {
      const start = serialOf(addDays(today, -1), date1904);
      return { start, end: todaySerial };
    }
    case 'today':
      return { start: todaySerial, end: serialOf(addDays(today, 1), date1904) };
    case 'tomorrow': {
      const start = serialOf(addDays(today, 1), date1904);
      return { start, end: serialOf(addDays(today, 2), date1904) };
    }
    case 'last7Days':
      return {
        start: serialOf(addDays(today, -6), date1904),
        end: serialOf(addDays(today, 1), date1904),
      };
    case 'lastWeek':
    case 'thisWeek':
    case 'nextWeek': {
      const weekStart = addDays(today, -weekdayIndex(today));
      const offset = period === 'lastWeek' ? -7 : period === 'nextWeek' ? 7 : 0;
      const startDate = addDays(weekStart, offset);
      return {
        start: serialOf(startDate, date1904),
        end: serialOf(addDays(startDate, 7), date1904),
      };
    }
    case 'lastMonth': {
      const thisMonth = startOfMonth(today);
      return {
        start: serialOf(addMonths(thisMonth, -1), date1904),
        end: serialOf(thisMonth, date1904),
      };
    }
    case 'thisMonth': {
      const thisMonth = startOfMonth(today);
      return {
        start: serialOf(thisMonth, date1904),
        end: serialOf(addMonths(thisMonth, 1), date1904),
      };
    }
    case 'nextMonth': {
      const nextMonth = addMonths(startOfMonth(today), 1);
      return {
        start: serialOf(nextMonth, date1904),
        end: serialOf(addMonths(nextMonth, 1), date1904),
      };
    }
  }
}

function serialForDateRule(cell: CellValue): number | null {
  if (cell.type === 'number') return cell.value;
  if (cell.type === 'blank') return 0; // §3.2: blanks coerce to 0 in numeric comparisons
  return null;
}

function compareSerials(
  operation: CfDateOperator,
  value: number,
  first: number,
  second?: number,
): boolean {
  switch (operation) {
    case 'equal':
      return value === first;
    case 'notEqual':
      return value !== first;
    case 'before':
      return value < first;
    case 'after':
      return value > first;
    case 'between':
      return second !== undefined && value >= first && value <= second;
    case 'notBetween':
      return second !== undefined && !(value >= first && value <= second);
  }
}

// ---------------------------------------------------------------------------
// Thresholds, percentile and colour (§3.3)
// ---------------------------------------------------------------------------

function clamp01(value: number): number {
  return value < 0 ? 0 : value > 1 ? 1 : value;
}

/** Excel's inclusive percentile: linear interpolation between ordered values. */
function percentileOf(sorted: readonly number[], percentile: number): number {
  if (sorted.length === 0) return 0;
  const position = (sorted.length - 1) * (Math.min(100, Math.max(0, percentile)) / 100);
  const lower = Math.floor(position);
  const upper = Math.ceil(position);
  if (lower === upper) return sorted[lower]!;
  const weight = position - lower;
  return sorted[lower]! * (1 - weight) + sorted[upper]! * weight;
}

interface NumericDomain {
  readonly sorted: readonly number[];
  readonly min: number;
  readonly max: number;
}

function domainOf(sorted: readonly number[]): NumericDomain {
  return { sorted, min: sorted[0]!, max: sorted[sorted.length - 1]! };
}

type ThresholdRole = 'min' | 'mid' | 'max';

function resolveThreshold(
  threshold: CfThreshold,
  domain: NumericDomain,
  role: ThresholdRole,
): number {
  switch (threshold.type) {
    case 'automatic':
      if (role === 'min') return domain.min;
      if (role === 'max') return domain.max;
      // A three-colour midpoint defaults to the 50th percentile (§3.3).
      return percentileOf(domain.sorted, 50);
    case 'number':
      return threshold.value;
    case 'percent': {
      const fraction = Math.min(100, Math.max(0, threshold.value)) / 100;
      return domain.min + (domain.max - domain.min) * fraction;
    }
    case 'percentile':
      return percentileOf(domain.sorted, threshold.value);
  }
}

/** Data bars use min(0, range min) and max(0, range max) for automatic points. */
function resolveDataBarThreshold(
  threshold: CfThreshold,
  domain: NumericDomain,
  role: ThresholdRole,
): number {
  if (threshold.type === 'automatic') {
    return role === 'min' ? Math.min(0, domain.min) : Math.max(0, domain.max);
  }
  return resolveThreshold(threshold, domain, role);
}

interface Rgb {
  readonly red: number;
  readonly green: number;
  readonly blue: number;
}

function parseRgb(color: string): Rgb | null {
  const trimmed = color.trim().replace(/^#/, '');
  let hex: string;
  if (/^[0-9a-f]{8}$/i.test(trimmed)) {
    hex = trimmed.slice(2); // OOXML ARGB: drop the alpha byte.
  } else if (/^[0-9a-f]{6}$/i.test(trimmed)) {
    hex = trimmed;
  } else {
    return null;
  }
  return {
    red: Number.parseInt(hex.slice(0, 2), 16),
    green: Number.parseInt(hex.slice(2, 4), 16),
    blue: Number.parseInt(hex.slice(4, 6), 16),
  };
}

function formatRgb(rgb: Rgb): string {
  const channel = (value: number): string =>
    Math.round(Math.min(255, Math.max(0, value)))
      .toString(16)
      .padStart(2, '0')
      .toUpperCase();
  return `#${channel(rgb.red)}${channel(rgb.green)}${channel(rgb.blue)}`;
}

function interpolateRgb(left: Rgb, right: Rgb, fraction: number): Rgb {
  return {
    red: left.red + (right.red - left.red) * fraction,
    green: left.green + (right.green - left.green) * fraction,
    blue: left.blue + (right.blue - left.blue) * fraction,
  };
}

// ---------------------------------------------------------------------------
// Range statistics (§6)
// ---------------------------------------------------------------------------

function numericValuesInRange(range: GridRange, source: CfCellSource): number[] {
  const values: number[] = [];
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    for (let col = range.startCol; col <= range.endCol; col += 1) {
      const value = source.valueAt(row, col);
      if (value.type === 'number' && Number.isFinite(value.value)) values.push(value.value);
    }
  }
  return values;
}

/**
 * Compute the statistics a range-statistic rule needs for one contiguous
 * range. Pure and deterministic; no clock, DOM or I/O. A caller may chunk the
 * scan itself and reuse this function's shape.
 */
export function computeRangeStatistics(
  rule: ConditionalFormatRule,
  range: GridRange,
  source: CfCellSource,
): CfRangeStatistics | null {
  switch (rule.kind) {
    case 'top10': {
      const values = numericValuesInRange(range, source).sort((left, right) => left - right);
      const count = values.length;
      const rank = rule.rank ?? 10;
      const requested = rule.percent
        ? rank > 0 && count > 0
          ? Math.max(1, Math.ceil((count * rank) / 100))
          : 0
        : Math.floor(rank);
      const qualifying = Math.min(count, Math.max(0, requested));
      let threshold: number | null = null;
      if (qualifying > 0) {
        threshold = rule.top ? values[count - qualifying]! : values[qualifying - 1]!;
      }
      return { kind: 'top10', count, rank, percent: rule.percent, top: rule.top, threshold };
    }
    case 'duplicate':
    case 'unique': {
      const counts = new Map<string, number>();
      for (let row = range.startRow; row <= range.endRow; row += 1) {
        for (let col = range.startCol; col <= range.endCol; col += 1) {
          const value = source.valueAt(row, col);
          if (value.type === 'blank') continue;
          const key = duplicateKey(value);
          counts.set(key, (counts.get(key) ?? 0) + 1);
        }
      }
      return { kind: rule.kind, counts };
    }
    case 'colorScale': {
      const values = numericValuesInRange(range, source);
      if (values.length === 0) return null;
      if (rule.thresholds.length < 2 || rule.colors.length !== rule.thresholds.length) return null;
      const domain = domainOf([...values].sort((left, right) => left - right));
      const points: CfColorScalePoint[] = [];
      for (let index = 0; index < rule.thresholds.length; index += 1) {
        const rgb = parseRgb(rule.colors[index]!);
        if (rgb === null) return null;
        const role: ThresholdRole =
          index === 0 ? 'min' : index === rule.thresholds.length - 1 ? 'max' : 'mid';
        points.push({
          value: resolveThreshold(rule.thresholds[index]!, domain, role),
          color: formatRgb(rgb),
        });
      }
      return { kind: 'colorScale', points };
    }
    case 'dataBar': {
      const values = numericValuesInRange(range, source);
      if (values.length === 0) return null;
      const domain = domainOf([...values].sort((left, right) => left - right));
      const first = resolveDataBarThreshold(rule.min, domain, 'min');
      const second = resolveDataBarThreshold(rule.max, domain, 'max');
      return {
        kind: 'dataBar',
        min: Math.min(first, second),
        max: Math.max(first, second),
      };
    }
    case 'iconSet': {
      const values = numericValuesInRange(range, source);
      if (values.length === 0) return null;
      const domain = domainOf([...values].sort((left, right) => left - right));
      const thresholds = rule.thresholds.map((threshold) =>
        resolveThreshold(threshold, domain, 'min'),
      );
      return { kind: 'iconSet', thresholds, min: domain.min, max: domain.max };
    }
    case 'cellIs':
    case 'text':
    case 'timePeriod':
    case 'unsupported':
      return null;
  }
}

/**
 * The key duplicate/unique compares. Typed values, so text `"1"` is not the
 * number `1`; text compares case-insensitively, matching Excel. Blanks never
 * reach this function.
 */
function duplicateKey(value: CellValue): string {
  if (value.type === 'text') return `t:${value.value.toLowerCase()}`;
  return valueKey(value);
}

// ---------------------------------------------------------------------------
// Per-cell rule evaluation (§3.3)
// ---------------------------------------------------------------------------

interface CfRuleHit {
  readonly colorScale?: CfColorScaleOverlay;
  readonly dataBar?: CfDataBarOverlay;
  readonly iconSet?: CfIconSetOverlay;
}

function evaluateCellIsRule(
  rule: Extract<ConditionalFormatRule, { kind: 'cellIs' }>,
  cell: CellValue,
): boolean {
  const first = compareCellToOperand(cell, rule.value1);
  if (first === null) return false;
  switch (rule.operator) {
    case 'lessThan':
      return first < 0;
    case 'lessThanOrEqual':
      return first <= 0;
    case 'equal':
      return first === 0;
    case 'notEqual':
      return first !== 0;
    case 'greaterThanOrEqual':
      return first >= 0;
    case 'greaterThan':
      return first > 0;
    case 'between':
    case 'notBetween': {
      if (rule.value2 === undefined) return false;
      const second = compareCellToOperand(cell, rule.value2);
      if (second === null) return false;
      const within = first >= 0 && second <= 0;
      return rule.operator === 'between' ? within : !within;
    }
  }
}

function evaluateTextRule(
  rule: Extract<ConditionalFormatRule, { kind: 'text' }>,
  cell: CellValue,
  displayText: string,
): boolean {
  if (cell.type === 'error') return false; // errors match no text rule
  return matchesTextRule(rule.operator, displayText, rule.text, cell.type === 'blank');
}

function evaluateDateRule(
  rule: Extract<ConditionalFormatRule, { kind: 'timePeriod' }>,
  cell: CellValue,
  today: CfCalendarDate,
  date1904: boolean,
): boolean {
  const serial = serialForDateRule(cell);
  if (serial === null) return false;
  if (rule.mode === 'comparison') {
    return compareSerials(rule.operator, serial, rule.value1, rule.value2);
  }
  const interval = periodInterval(rule.period, today, date1904);
  return serial >= interval.start && serial < interval.end;
}

function evaluateRangeRule(
  rule: ConditionalFormatRule,
  cell: CellValue,
  statistics: CfRangeStatistics | null,
): CfRuleHit | null {
  if (statistics === null) return null;
  switch (rule.kind) {
    case 'top10': {
      if (statistics.kind !== 'top10' || statistics.threshold === null) return null;
      if (cell.type !== 'number') return null;
      const qualifies = rule.top
        ? cell.value >= statistics.threshold
        : cell.value <= statistics.threshold;
      return qualifies ? {} : null;
    }
    case 'duplicate': {
      if (statistics.kind !== 'duplicate' || cell.type === 'blank') return null;
      return (statistics.counts.get(duplicateKey(cell)) ?? 0) > 1 ? {} : null;
    }
    case 'unique': {
      if (statistics.kind !== 'unique' || cell.type === 'blank') return null;
      return (statistics.counts.get(duplicateKey(cell)) ?? 0) === 1 ? {} : null;
    }
    case 'colorScale': {
      if (statistics.kind !== 'colorScale' || cell.type !== 'number') return null;
      const overlay = colorScaleOverlay(rule, cell.value, statistics.points);
      return overlay === null ? null : { colorScale: overlay };
    }
    case 'dataBar': {
      if (statistics.kind !== 'dataBar' || cell.type !== 'number') return null;
      const overlay = dataBarOverlay(rule, cell.value, statistics);
      return overlay === null ? null : { dataBar: overlay };
    }
    case 'iconSet': {
      if (statistics.kind !== 'iconSet' || cell.type !== 'number') return null;
      const overlay = iconSetOverlay(rule, cell.value, statistics);
      return overlay === null ? null : { iconSet: overlay };
    }
    case 'cellIs':
    case 'text':
    case 'timePeriod':
    case 'unsupported':
      return null;
  }
}

function colorScaleOverlay(
  rule: Extract<ConditionalFormatRule, { kind: 'colorScale' }>,
  value: number,
  points: readonly CfColorScalePoint[],
): CfColorScaleOverlay | null {
  if (points.length < 2) return null;
  const first = points[0]!;
  const last = points[points.length - 1]!;
  if (value <= first.value) {
    return { kind: 'colorScale', ruleId: rule.id, value, position: 0, color: first.color, points };
  }
  if (value >= last.value) {
    return { kind: 'colorScale', ruleId: rule.id, value, position: 1, color: last.color, points };
  }
  for (let index = 0; index < points.length - 1; index += 1) {
    const left = points[index]!;
    const right = points[index + 1]!;
    if (value >= left.value && value <= right.value) {
      const span = right.value - left.value;
      const fraction = span === 0 ? 0 : (value - left.value) / span;
      const leftRgb = parseRgb(left.color);
      const rightRgb = parseRgb(right.color);
      if (leftRgb === null || rightRgb === null) return null;
      return {
        kind: 'colorScale',
        ruleId: rule.id,
        value,
        position: clamp01((index + fraction) / (points.length - 1)),
        color: formatRgb(interpolateRgb(leftRgb, rightRgb, fraction)),
        points,
      };
    }
  }
  return {
    kind: 'colorScale',
    ruleId: rule.id,
    value,
    position: 1,
    color: last.color,
    points,
  };
}

function dataBarOverlay(
  rule: Extract<ConditionalFormatRule, { kind: 'dataBar' }>,
  value: number,
  statistics: CfDataBarStatistics,
): CfDataBarOverlay | null {
  const { min, max } = statistics;
  const span = max - min;
  const fraction = (candidate: number): number =>
    span <= 0 ? 0 : clamp01((candidate - min) / span);
  const zero = fraction(0);
  const position = fraction(value);
  const negative = value < 0;
  const start = Math.min(zero, position);
  const end = Math.max(zero, position);
  const color = negative ? (rule.negativeColor ?? rule.color) : rule.color;
  if (parseRgb(color) === null) return null;
  return {
    kind: 'dataBar',
    ruleId: rule.id,
    value,
    min,
    max,
    start,
    end,
    axis: clamp01(zero),
    negative,
    color,
    fill: rule.fill ?? 'gradient',
    direction: rule.direction ?? 'leftToRight',
    border: rule.border ?? false,
    showValue: rule.showValue ?? true,
    axisColor: rule.axisColor ?? null,
  };
}

function iconSetOverlay(
  rule: Extract<ConditionalFormatRule, { kind: 'iconSet' }>,
  value: number,
  statistics: CfIconSetStatistics,
): CfIconSetOverlay | null {
  const definition = CF_ICON_SETS.find((candidate) => candidate.id === rule.iconSet);
  if (definition === undefined) return null;
  const iconCount = definition.iconCount;
  const thresholds = statistics.thresholds;
  let index = 0;
  for (const threshold of thresholds) {
    if (value >= threshold) index += 1;
  }
  index = Math.min(iconCount - 1, index);
  const iconIndex = rule.reverse === true ? iconCount - 1 - index : index;
  const iconOnly = rule.iconOnly ?? false;
  return {
    kind: 'iconSet',
    ruleId: rule.id,
    iconSet: rule.iconSet,
    iconCount,
    iconIndex,
    value,
    thresholds,
    showValue: !iconOnly && (rule.showValue ?? true),
    iconOnly,
  };
}

// ---------------------------------------------------------------------------
// Index and cache (§6)
// ---------------------------------------------------------------------------

interface StatsCacheEntry {
  readonly range: GridRange;
}

/**
 * Per-sheet CF index. Rules are sorted once by priority, a row-band range index
 * answers "which rules touch this cell", per-cell results cache by rule plus
 * cell value, and range statistics cache per rule plus contiguous range.
 *
 * A value edit should call `invalidateCell`; a format-only edit needs nothing,
 * because it repaints without re-evaluating. Recomputation is lazy: the next
 * `evaluateCell` call recomputes what it needs.
 */
export class ConditionalFormatIndex {
  readonly sheetId: SheetId;

  private orderedRules: ConditionalFormatRule[];
  private boundaries: readonly number[] = [];
  private bands: readonly CfRowBand[] = [];
  private date1904: boolean;
  private today: CfCalendarDate;
  private cells: CfCellSource;
  private dxfFor: CfDxfResolver | null;

  private readonly cellCache = new Map<string, boolean>();
  private readonly cellCacheKeys = new Map<string, Set<string>>();
  private readonly statsCache = new Map<string, CfRangeStatistics | null>();
  private readonly statsEntries = new Map<string, StatsCacheEntry>();

  constructor(options: ConditionalFormatIndexOptions) {
    this.sheetId = options.sheetId;
    this.orderedRules = orderRulesByPriority(options.rules);
    this.date1904 = options.date1904;
    this.today = options.today;
    this.cells = options.cells ?? EMPTY_CELL_SOURCE;
    this.dxfFor = options.dxfFor ?? null;
    this.rebuildIndex();
  }

  /** Rules in priority order. */
  get rules(): readonly ConditionalFormatRule[] {
    return this.orderedRules;
  }

  get dateSystem1904(): boolean {
    return this.date1904;
  }

  get currentDay(): CfCalendarDate {
    return this.today;
  }

  setRules(rules: readonly ConditionalFormatRule[]): void {
    this.orderedRules = orderRulesByPriority(rules);
    this.rebuildIndex();
    this.clearAllCaches();
  }

  setDateSystem(date1904: boolean): void {
    if (this.date1904 === date1904) return;
    this.date1904 = date1904;
    this.clearCellCache();
  }

  /** Change "today"; date results are recomputed lazily on the next evaluation. */
  setToday(today: CfCalendarDate): void {
    if (
      this.today.year === today.year &&
      this.today.month === today.month &&
      this.today.day === today.day
    ) {
      return;
    }
    this.today = today;
    this.clearCellCache();
  }

  setCellSource(cells: CfCellSource): void {
    this.cells = cells;
    this.clearAllCaches();
  }

  setDxfResolver(dxfFor: CfDxfResolver | null): void {
    this.dxfFor = dxfFor;
    this.clearAllCaches();
  }

  /** Rules covering a cell, in priority order, one placement per rule. */
  placementsAt(row: number, col: number): readonly CfRulePlacement[] {
    const band = bandForRow(this.boundaries, this.bands, row);
    if (band === null) return [];
    const result: CfRulePlacement[] = [];
    const seen = new Set<string>();
    for (const placement of band.placements) {
      if (col < placement.range.startCol || col > placement.range.endCol) continue;
      if (seen.has(placement.rule.id)) continue; // overlapping ranges of one rule
      seen.add(placement.rule.id);
      result.push(placement);
    }
    return result;
  }

  /** Rule ids covering a cell, highest priority first. */
  ruleIdsAt(row: number, col: number): readonly string[] {
    return this.placementsAt(row, col).map((placement) => placement.rule.id);
  }

  /**
   * Evaluate every rule covering one cell. `subject` may be passed for visible
   * cells; otherwise the cell source supplies the value and display text.
   */
  evaluateCell(row: number, col: number, subject?: CfCellSubject): CfCellEvaluation {
    const placements = this.placementsAt(row, col);
    if (placements.length === 0) return NOTHING;
    const resolved: Required<CfCellSubject> =
      subject === undefined
        ? {
            value: this.cells.valueAt(row, col),
            displayText: this.cells.displayTextAt(row, col),
          }
        : {
            value: subject.value,
            displayText: subject.displayText ?? formatCellValue(subject.value),
          };
    return this.evaluatePlacements(row, col, placements, resolved);
  }

  /** Statistics for a rule's contiguous range, or null when the rule is per-cell. */
  statisticsFor(ruleId: string, rangeIndex: number): CfRangeStatistics | null {
    const rule = this.orderedRules.find((candidate) => candidate.id === ruleId);
    if (rule === undefined) return null;
    const range = rule.ranges[rangeIndex];
    if (range === undefined) return null;
    return this.statisticsForPlacement({ rule, range, rangeIndex, order: 0 });
  }

  /**
   * Invalidate one cell after a value edit: its cached per-cell results and the
   * statistics of every cached range containing it (§6).
   */
  invalidateCell(row: number, col: number): void {
    const cellKey = `${row}:${col}`;
    const keys = this.cellCacheKeys.get(cellKey);
    if (keys !== undefined) {
      for (const key of keys) this.cellCache.delete(key);
      this.cellCacheKeys.delete(cellKey);
    }
    for (const [key, entry] of this.statsEntries) {
      if (rangeContains(entry.range, row, col)) {
        this.statsCache.delete(key);
        this.statsEntries.delete(key);
      }
    }
  }

  clearCaches(): void {
    this.clearAllCaches();
  }

  private rebuildIndex(): void {
    const index = buildRowBands(this.orderedRules);
    this.boundaries = index.boundaries;
    this.bands = index.bands;
  }

  private clearCellCache(): void {
    this.cellCache.clear();
    this.cellCacheKeys.clear();
  }

  private clearAllCaches(): void {
    this.clearCellCache();
    this.statsCache.clear();
    this.statsEntries.clear();
  }

  private evaluatePlacements(
    row: number,
    col: number,
    placements: readonly CfRulePlacement[],
    subject: Required<CfCellSubject>,
  ): CfCellEvaluation {
    const ruleIds: string[] = [];
    const unsupportedRuleIds: string[] = [];
    let fill: CfFillPatch | null = null;
    let font: CfFontPatch | null = null;
    let border: CfBorderPatch | null = null;
    let numberFormat: CfNumberFormatPatch | null = null;
    let colorScale: CfColorScaleOverlay | null = null;
    let dataBar: CfDataBarOverlay | null = null;
    let iconSet: CfIconSetOverlay | null = null;
    let stoppedBy: string | null = null;

    for (const placement of placements) {
      const rule = placement.rule;
      if (rule.kind === 'unsupported') {
        unsupportedRuleIds.push(rule.id);
        continue; // an unknown truth value never stops lower rules
      }
      const hit = this.evaluateRule(placement, subject, row, col);
      if (hit === null) continue;
      ruleIds.push(rule.id);
      if (hit.colorScale !== undefined && colorScale === null) colorScale = hit.colorScale;
      if (hit.dataBar !== undefined && dataBar === null) dataBar = hit.dataBar;
      if (hit.iconSet !== undefined && iconSet === null) iconSet = hit.iconSet;
      if (rule.dxfId !== undefined && this.dxfFor !== null) {
        const dxf = this.dxfFor(rule.dxfId);
        if (dxf !== null) {
          if (fill === null && dxf.fill !== undefined) fill = dxf.fill;
          if (font === null && dxf.font !== undefined) font = dxf.font;
          if (border === null && dxf.border !== undefined) border = dxf.border;
          if (numberFormat === null && dxf.numberFormat !== undefined) {
            numberFormat = dxf.numberFormat;
          }
        }
      }
      if (rule.stopIfTrue) {
        stoppedBy = rule.id;
        break;
      }
    }

    const overlay: CfOverlay = Object.freeze({ fill, font, border, numberFormat });
    const painted =
      fill !== null ||
      font !== null ||
      border !== null ||
      numberFormat !== null ||
      colorScale !== null ||
      dataBar !== null ||
      iconSet !== null;
    return {
      ruleIds,
      stoppedBy,
      overlay,
      colorScale,
      dataBar,
      iconSet,
      unsupportedRuleIds,
      painted,
    };
  }

  private evaluateRule(
    placement: CfRulePlacement,
    subject: CfCellSubject,
    row: number,
    col: number,
  ): CfRuleHit | null {
    const rule = placement.rule;
    switch (rule.kind) {
      case 'cellIs': {
        const matched = this.cachedPerCell(placement, subject, row, col, () =>
          evaluateCellIsRule(rule, subject.value),
        );
        return matched ? {} : null;
      }
      case 'text': {
        const matched = this.cachedPerCell(placement, subject, row, col, () =>
          evaluateTextRule(
            rule,
            subject.value,
            subject.displayText ?? formatCellValue(subject.value),
          ),
        );
        return matched ? {} : null;
      }
      case 'timePeriod': {
        const matched = this.cachedPerCell(placement, subject, row, col, () =>
          evaluateDateRule(rule, subject.value, this.today, this.date1904),
        );
        return matched ? {} : null;
      }
      case 'top10':
      case 'duplicate':
      case 'unique':
      case 'colorScale':
      case 'dataBar':
      case 'iconSet': {
        const statistics = this.statisticsForPlacement(placement);
        return evaluateRangeRule(rule, subject.value, statistics);
      }
      case 'unsupported':
        return null;
    }
  }

  /**
   * Cache per-cell rule results by rule plus cell value. The cell address is
   * not part of the key because cell-value, text and date rules are
   * positionless; the address only records which keys to drop on invalidation.
   */
  private cachedPerCell(
    placement: CfRulePlacement,
    subject: CfCellSubject,
    row: number,
    col: number,
    evaluate: () => boolean,
  ): boolean {
    const value = subject.value;
    const displayText = subject.displayText ?? formatCellValue(value);
    const key = `${placement.rule.id}\u0000${valueKey(value)}\u0000${displayText}`;
    this.recordCellCacheKey(row, col, key);
    const cached = this.cellCache.get(key);
    if (cached !== undefined) return cached;
    const result = evaluate();
    this.cellCache.set(key, result);
    return result;
  }

  private recordCellCacheKey(row: number, col: number, key: string): void {
    const cellKey = `${row}:${col}`;
    const keys = this.cellCacheKeys.get(cellKey);
    if (keys === undefined) {
      this.cellCacheKeys.set(cellKey, new Set([key]));
      return;
    }
    keys.add(key);
  }

  private statisticsForPlacement(placement: CfRulePlacement): CfRangeStatistics | null {
    const key = `${placement.rule.id}\u0000${placement.rangeIndex}`;
    if (this.statsCache.has(key)) return this.statsCache.get(key) ?? null;
    const statistics = computeRangeStatistics(placement.rule, placement.range, this.cells);
    this.statsCache.set(key, statistics);
    this.statsEntries.set(key, { range: placement.range });
    return statistics;
  }
}

export function createConditionalFormatIndex(
  options: ConditionalFormatIndexOptions,
): ConditionalFormatIndex {
  return new ConditionalFormatIndex(options);
}
