import type { SaxesAttributeNS, SaxesTagNS } from 'saxes';
import type {
  AutoFilter,
  ColumnFilter,
  ColumnOffset,
  Condition,
  ConditionsFilter,
  GridRange,
  NumberComparisonOperator,
  NumberCondition,
  RowDimension,
  Sheet,
  SheetId,
  Workbook,
} from '@model/types';
import { cellKey, columnToLetters, parseRangeRef, singleCellRange } from '@model/types';
import { evaluateAutoFilter, restoreHiddenLayers } from '@model/autofilter';
import type { CellValue } from '@model/value';
import { errorValue, formatCellValue, isCellErrorCode, numberValue, textValue } from '@model/value';
import {
  decodeXml,
  encodeXml,
  escapeXmlAttribute,
  isWellFormedXml,
  openTagStart,
  parseXml,
} from './xml';
import { WorkbookSaveError } from './errors';

/**
 * The `autoFilter` element of a worksheet part (sort and AutoFilter §3.6).
 *
 * Reading maps the XML onto `Sheet.autoFilter`; writing emits the element. An
 * untouched filter is re-emitted from the exact source text it was parsed from,
 * so a save that does not change it keeps the part byte-identical. Everything
 * calcplus cannot model — a colour, icon, Top 10 or dynamic criterion, an
 * unknown custom operator, an unknown attribute or child — is kept verbatim in
 * `AutoFilter.raw` and each `UnsupportedFilter.raw` (§3.7).
 *
 * This module owns the XML only. It never dispatches a command and never
 * changes criteria; the model layer owns behaviour.
 */

// ---------------------------------------------------------------------------
// Verbatim source records
// ---------------------------------------------------------------------------

/** One XML attribute, name and value exactly as authored. */
export interface XmlAttributeRecord {
  readonly name: string;
  readonly value: string;
}

/** A `filterColumn`'s own attributes, so `hiddenButton` and friends survive a rebuild. */
export interface AutoFilterRawColumn {
  readonly colId: number;
  readonly attributes: readonly XmlAttributeRecord[];
}

/**
 * The parsed source of one `autoFilter` element. Stored in `AutoFilter.raw`.
 * `xml` is the whole element and is emitted unchanged whenever `stateKey`
 * still matches the model; the other fields let a changed filter be rebuilt
 * while keeping unknown attributes and children in their authored order.
 */
export interface AutoFilterRawSource {
  readonly kind: 'autoFilter-source';
  /** The exact source element, emitted verbatim on an untouched save. */
  readonly xml: string;
  readonly prefix: string;
  /** Attributes of `autoFilter` other than `ref`, in document order. */
  readonly attributes: readonly XmlAttributeRecord[];
  /** `filterColumn` entries in document order, with their non-`colId` attributes. */
  readonly columns: readonly AutoFilterRawColumn[];
  /** Child elements that are not `filterColumn`, verbatim. */
  readonly unknownChildren: readonly string[];
  /** Criteria key at parse time; equality means the model still holds the source. */
  readonly stateKey: string;
}

export type UnsupportedFilterReason =
  'colour' | 'icon' | 'top10' | 'dynamic' | 'date-group' | 'unknown-operator' | 'unknown';

/** The parsed source of a `filterColumn` calcplus cannot evaluate (§3.7). */
export interface UnsupportedFilterSource {
  readonly kind: 'filterColumn-source';
  /** The exact `filterColumn` element, preserved verbatim. */
  readonly xml: string;
  /** The element's children, verbatim; used when the column must be rebuilt. */
  readonly inner: string;
  readonly reason: UnsupportedFilterReason;
  /** The child element or operator that made the criterion unsupported. */
  readonly detail: string;
}

export function isAutoFilterRawSource(raw: unknown): raw is AutoFilterRawSource {
  if (raw === null || typeof raw !== 'object') return false;
  const candidate = raw as Partial<AutoFilterRawSource>;
  return (
    candidate.kind === 'autoFilter-source' &&
    typeof candidate.xml === 'string' &&
    typeof candidate.stateKey === 'string' &&
    Array.isArray(candidate.attributes) &&
    Array.isArray(candidate.columns) &&
    Array.isArray(candidate.unknownChildren)
  );
}

export function isUnsupportedFilterSource(raw: unknown): raw is UnsupportedFilterSource {
  if (raw === null || typeof raw !== 'object') return false;
  const candidate = raw as Partial<UnsupportedFilterSource>;
  return (
    candidate.kind === 'filterColumn-source' &&
    typeof candidate.xml === 'string' &&
    typeof candidate.inner === 'string'
  );
}

// ---------------------------------------------------------------------------
// Parsing
// ---------------------------------------------------------------------------

interface XmlFragmentElement {
  readonly local: string;
  readonly prefix: string;
  readonly attributes: readonly XmlAttributeRecord[];
  readonly elements: XmlFragmentElement[];
  readonly start: number;
  readonly end: number;
  readonly openEnd: number;
}

interface ColumnTyping {
  /** Authored `val` text to the typed value the column's cells carry. */
  readonly byText: ReadonlyMap<string, CellValue>;
  readonly numericCells: number;
  readonly textualCells: number;
  readonly otherCells: number;
}

const NO_TYPING: ColumnTyping = {
  byText: new Map(),
  numericCells: 0,
  textualCells: 0,
  otherCells: 0,
};

/**
 * Parse one `autoFilter` element into the model (§3.1, §3.6). `sheet` supplies
 * the stable sheet id and the cells whose typing decides what a checklist value
 * means: number `10` is not text `"10"` (§3.2).
 */
export function parseAutoFilterElement(elementXml: string, sheet: Sheet): AutoFilter | null {
  const root = parseElementFragment(elementXml, `${sheet.partPath} (autoFilter)`);
  if (root === null || root.local !== 'autoFilter') return null;

  const range = rangeOfAttributes(root.attributes);
  const attributes = root.attributes.filter((attribute) => localName(attribute.name) !== 'ref');
  const columns = new Map<ColumnOffset, ColumnFilter>();
  const rawColumns: AutoFilterRawColumn[] = [];
  const unknownChildren: string[] = [];

  for (const child of root.elements) {
    const source = elementXml.slice(child.start, child.end);
    if (child.local !== 'filterColumn') {
      unknownChildren.push(source);
      continue;
    }
    const colIdText = attributeValueOf(child.attributes, 'colId');
    const colId = colIdText === undefined ? Number.NaN : Number(colIdText);
    if (!Number.isInteger(colId) || colId < 0) {
      unknownChildren.push(source);
      continue;
    }
    rawColumns.push({
      colId,
      attributes: child.attributes.filter((attribute) => localName(attribute.name) !== 'colId'),
    });
    columns.set(colId, parseFilterColumn(child, columnTyping(sheet, range, colId), elementXml));
  }

  const parsed: AutoFilter = {
    sheetId: sheet.id,
    range,
    columns,
    raw: {
      kind: 'autoFilter-source',
      xml: elementXml,
      prefix: root.prefix,
      attributes,
      columns: rawColumns,
      unknownChildren,
      stateKey: '',
    },
  };
  return {
    ...parsed,
    raw: { ...(parsed.raw as AutoFilterRawSource), stateKey: autoFilterStateKey(parsed) },
  };
}

function parseFilterColumn(
  node: XmlFragmentElement,
  typing: ColumnTyping,
  sourceXml: string,
): ColumnFilter {
  if (node.elements.length !== 1) {
    return unsupportedColumn(node, sourceXml, 'unknown', 'filterColumn with no single criterion');
  }
  const criterion = node.elements[0]!;
  switch (criterion.local) {
    case 'filters':
      return parseFilters(criterion, typing, node, sourceXml);
    case 'customFilters':
      return parseCustomFilters(criterion, typing, node, sourceXml);
    case 'colorFilter':
      return unsupportedColumn(node, sourceXml, 'colour', 'colorFilter');
    case 'iconFilter':
      return unsupportedColumn(node, sourceXml, 'icon', 'iconFilter');
    case 'top10':
      return unsupportedColumn(node, sourceXml, 'top10', 'top10');
    case 'dynamicFilter':
      return unsupportedColumn(node, sourceXml, 'dynamic', 'dynamicFilter');
    default:
      return unsupportedColumn(node, sourceXml, 'unknown', criterion.local);
  }
}

function parseFilters(
  node: XmlFragmentElement,
  typing: ColumnTyping,
  columnNode: XmlFragmentElement,
  sourceXml: string,
): ColumnFilter {
  const extra = node.attributes.filter((attribute) => localName(attribute.name) !== 'blank');
  if (extra.length > 0) {
    return unsupportedColumn(columnNode, sourceXml, 'unknown', 'filters attributes');
  }
  const selected: CellValue[] = [];
  for (const child of node.elements) {
    if (child.local !== 'filter') {
      return unsupportedColumn(columnNode, sourceXml, 'date-group', child.local);
    }
    const value = attributeValueOf(child.attributes, 'val');
    if (value === undefined) {
      return unsupportedColumn(columnNode, sourceXml, 'unknown', 'filter without val');
    }
    selected.push(typedValueFromText(value, typing));
  }
  const blank = attributeValueOf(node.attributes, 'blank');
  return {
    kind: 'values',
    selected,
    includeBlanks: blank === '1' || blank === 'true',
  };
}

function parseCustomFilters(
  node: XmlFragmentElement,
  typing: ColumnTyping,
  columnNode: XmlFragmentElement,
  sourceXml: string,
): ColumnFilter {
  const operators: { operator: string; value: string }[] = [];
  for (const child of node.elements) {
    if (child.local !== 'customFilter') {
      return unsupportedColumn(columnNode, sourceXml, 'unknown', child.local);
    }
    operators.push({
      operator: attributeValueOf(child.attributes, 'operator') ?? 'equal',
      value: attributeValueOf(child.attributes, 'val') ?? '',
    });
  }
  if (operators.length === 0 || operators.length > 2) {
    return unsupportedColumn(columnNode, sourceXml, 'unknown', 'customFilters count');
  }
  const and = attributeValueOf(node.attributes, 'and');
  const join = and === '1' || and === 'true' ? 'and' : 'or';
  const conditions = conditionsFromOperators(operators, join, typing);
  if (conditions === null) {
    return unsupportedColumn(
      columnNode,
      sourceXml,
      'unknown-operator',
      operators.map((operator) => operator.operator).join(','),
    );
  }
  return { kind: 'conditions', join, conditions };
}

/**
 * Map Excel's `customFilter` spellings onto the model's condition shapes.
 * `between` and `not between` are the documented two-filter encodings (§3.6):
 * `>=` AND `<=` for between, `<` OR `>` for not between.
 */
function conditionsFromOperators(
  operators: readonly { operator: string; value: string }[],
  join: 'and' | 'or',
  typing: ColumnTyping,
): readonly Condition[] | null {
  if (operators.length === 2) {
    const [first, second] = operators;
    const between = betweenCondition(first!, second!, join);
    if (between !== null) return [between];
  }
  const conditions: Condition[] = [];
  for (const operator of operators) {
    const condition = mapCustomOperator(operator.operator, operator.value, typing);
    if (condition === null) return null;
    conditions.push(condition);
  }
  return conditions;
}

function betweenCondition(
  first: { operator: string; value: string },
  second: { operator: string; value: string },
  join: 'and' | 'or',
): NumberCondition | null {
  if (join === 'and') {
    if (first.operator !== 'greaterThanOrEqual' || second.operator !== 'lessThanOrEqual')
      return null;
    return numericBetween('between', first.value, second.value);
  }
  if (first.operator !== 'lessThan' || second.operator !== 'greaterThan') return null;
  return numericBetween('not-between', first.value, second.value);
}

function numericBetween(
  operator: 'between' | 'not-between',
  fromText: string,
  toText: string,
): NumberCondition | null {
  const from = numericOperand(fromText);
  const to = numericOperand(toText);
  if (from === null || to === null) return null;
  return { kind: 'number', operator, from, to };
}

function numericOperand(text: string): number | null {
  if (text.trim() === '') return null;
  const value = Number(text);
  return Number.isFinite(value) ? value : null;
}

function mapCustomOperator(
  operator: string,
  value: string,
  typing: ColumnTyping,
): Condition | null {
  switch (operator) {
    case 'equal':
      return equalityCondition(value, false, typing);
    case 'notEqual':
      return equalityCondition(value, true, typing);
    case 'greaterThan':
      return numberCondition('>', value);
    case 'greaterThanOrEqual':
      return numberCondition('>=', value);
    case 'lessThan':
      return numberCondition('<', value);
    case 'lessThanOrEqual':
      return numberCondition('<=', value);
    default:
      return null;
  }
}

/**
 * `equal`/`notEqual` cover both text and number filters in Excel, so the
 * wildcard shape and the column's typed cells decide (§3.2). A negation of a
 * "begins with" or "ends with" pattern has no model operator and stays
 * unsupported rather than being distorted.
 */
function equalityCondition(
  value: string,
  negated: boolean,
  typing: ColumnTyping,
): Condition | null {
  const wildcard = wildcardTextCondition(value, negated);
  if (wildcard !== 'no-pattern') return wildcard;

  const typed = typing.byText.get(value);
  if (typed !== undefined) {
    if (typed.type === 'number') {
      return { kind: 'number', operator: negated ? '<>' : '=', value: typed.value };
    }
    if (typed.type === 'text') {
      return { kind: 'text', operator: negated ? 'not-equals' : 'equals', value };
    }
    // Booleans and errors are checklist values; they never match a condition.
    return null;
  }
  const numeric = numericOperand(value);
  if (numeric !== null) {
    return { kind: 'number', operator: negated ? '<>' : '=', value: numeric };
  }
  return { kind: 'text', operator: negated ? 'not-equals' : 'equals', value };
}

/** `null` means the pattern has no model operator; `'no-pattern'` means it has no wildcard. */
function wildcardTextCondition(value: string, negated: boolean): Condition | null | 'no-pattern' {
  if (value === '*') {
    return { kind: 'text', operator: negated ? 'does-not-contain' : 'contains', value: '' };
  }
  if (value.length > 1 && value.startsWith('*') && value.endsWith('*')) {
    return {
      kind: 'text',
      operator: negated ? 'does-not-contain' : 'contains',
      value: value.slice(1, -1),
    };
  }
  if (value.length > 1 && value.endsWith('*')) {
    if (negated) return null;
    return { kind: 'text', operator: 'begins-with', value: value.slice(0, -1) };
  }
  if (value.length > 1 && value.startsWith('*')) {
    if (negated) return null;
    return { kind: 'text', operator: 'ends-with', value: value.slice(1) };
  }
  return 'no-pattern';
}

function numberCondition(operator: NumberComparisonOperator, text: string): Condition | null {
  const value = numericOperand(text);
  if (value === null) return null;
  return { kind: 'number', operator, value };
}

function unsupportedColumn(
  node: XmlFragmentElement,
  sourceXml: string,
  reason: UnsupportedFilterReason,
  detail: string,
): ColumnFilter {
  const closeStart = node.end <= node.openEnd ? -1 : sourceXml.lastIndexOf('</', node.end - 1);
  const inner = closeStart < node.openEnd ? '' : sourceXml.slice(node.openEnd, closeStart);
  return {
    kind: 'unsupported',
    raw: {
      kind: 'filterColumn-source',
      xml: sourceXml.slice(node.start, node.end),
      inner,
      reason,
      detail,
    },
  };
}

// ---------------------------------------------------------------------------
// Cell typing for checklist values (§3.2)
// ---------------------------------------------------------------------------

function columnTyping(sheet: Sheet, range: GridRange, offset: ColumnOffset): ColumnTyping {
  const column = range.startCol + offset;
  if (column < range.startCol || column > range.endCol) return NO_TYPING;
  const byText = new Map<string, CellValue>();
  let numericCells = 0;
  let textualCells = 0;
  let otherCells = 0;
  for (let row = range.startRow + 1; row <= range.endRow; row += 1) {
    const cell = sheet.cells.get(cellKey(row, column));
    if (cell === undefined) continue;
    const value = cell.value;
    if (value.type === 'blank') continue;
    for (const text of valueTexts(value)) {
      if (!byText.has(text)) byText.set(text, value);
    }
    if (value.type === 'number') numericCells += 1;
    else if (value.type === 'text') textualCells += 1;
    else otherCells += 1;
  }
  return { byText, numericCells, textualCells, otherCells };
}

function valueTexts(value: CellValue): string[] {
  switch (value.type) {
    case 'number':
      return [formatCellValue(value)];
    case 'text':
      return [value.value];
    case 'boolean':
      return [formatCellValue(value)];
    case 'error':
      return [value.value];
    case 'blank':
      return [];
  }
}

function typedValueFromText(text: string, typing: ColumnTyping): CellValue {
  const typed = typing.byText.get(text);
  if (typed !== undefined) return typed;
  const numeric = numericOperand(text);
  if (numeric !== null && (typing.textualCells === 0 || typing.numericCells > 0)) {
    return numberValue(numeric);
  }
  if (isCellErrorCode(text)) return errorValue(text);
  return textValue(text);
}

// ---------------------------------------------------------------------------
// Criteria identity
// ---------------------------------------------------------------------------

/**
 * A deterministic key for the filter's document state. `AutoFilter.raw` records
 * the key at parse time; while the model still produces the same key the
 * source element is re-emitted verbatim, which is what makes an untouched save
 * byte-identical (§3.6).
 */
export function autoFilterStateKey(filter: AutoFilter): string {
  const { range } = filter;
  const parts = [`${range.startRow},${range.startCol},${range.endRow},${range.endCol}`];
  for (const offset of [...filter.columns.keys()].sort((left, right) => left - right)) {
    const column = filter.columns.get(offset);
    if (column !== undefined) parts.push(`${offset}=${columnFilterKey(column)}`);
  }
  return parts.join('|');
}

function columnFilterKey(filter: ColumnFilter): string {
  switch (filter.kind) {
    case 'values':
      return `v:${filter.includeBlanks ? '1' : '0'}:${filter.selected
        .map((value) => `${value.type}:${formatCellValue(value)}`)
        .join(',')}`;
    case 'conditions':
      return `c:${filter.join}:${filter.conditions.map(conditionKey).join(';')}`;
    case 'unsupported': {
      const source = isUnsupportedFilterSource(filter.raw)
        ? filter.raw.xml
        : stableText(filter.raw);
      return `u:${source}`;
    }
  }
}

function conditionKey(condition: Condition): string {
  if ('from' in condition) {
    return `${condition.kind}:${condition.operator}:${plainNumber(condition.from)}:${plainNumber(
      condition.to,
    )}`;
  }
  return `${condition.kind}:${condition.operator}:${String(condition.value)}`;
}

function stableText(value: unknown): string {
  try {
    return JSON.stringify(value) ?? String(value);
  } catch {
    return String(value);
  }
}

// ---------------------------------------------------------------------------
// Serialisation (§3.6)
// ---------------------------------------------------------------------------

/**
 * The `autoFilter` element for the model. An untouched filter is emitted from
 * its recorded source text; a changed one is rebuilt with the same attributes,
 * unknown children and per-column attributes in their authored order, and
 * unsupported criteria's inner XML verbatim.
 */
export function serialiseAutoFilterElement(filter: AutoFilter): string {
  const source = isAutoFilterRawSource(filter.raw) ? filter.raw : null;
  if (source !== null && source.stateKey === autoFilterStateKey(filter)) return source.xml;

  const prefix = source?.prefix ?? '';
  const name = qualified(prefix, 'autoFilter');
  const attributes = [` ref="${escapeXmlAttribute(formatRange(filter.range))}"`];
  for (const attribute of source?.attributes ?? []) {
    attributes.push(` ${attribute.name}="${escapeXmlAttribute(attribute.value)}"`);
  }

  const order = columnOrder(filter, source);
  const children = order
    .map((colId) => {
      const column = filter.columns.get(colId);
      return column === undefined ? '' : filterColumnXml(prefix, colId, column, source);
    })
    .join('');
  const unknown = (source?.unknownChildren ?? []).join('');

  if (children === '' && unknown === '') return `<${name}${attributes.join('')}/>`;
  return `<${name}${attributes.join('')}>${children}${unknown}</${name}>`;
}

function columnOrder(
  filter: AutoFilter,
  source: AutoFilterRawSource | null,
): readonly ColumnOffset[] {
  const order: ColumnOffset[] = [];
  for (const column of source?.columns ?? []) {
    if (filter.columns.has(column.colId) && !order.includes(column.colId)) order.push(column.colId);
  }
  for (const offset of [...filter.columns.keys()].sort((left, right) => left - right)) {
    if (!order.includes(offset)) order.push(offset);
  }
  return order;
}

function filterColumnXml(
  prefix: string,
  colId: ColumnOffset,
  filter: ColumnFilter,
  source: AutoFilterRawSource | null,
): string {
  const name = qualified(prefix, 'filterColumn');
  const attributes = [` colId="${colId}"`];
  const preserved = source?.columns.find((column) => column.colId === colId)?.attributes ?? [];
  for (const attribute of preserved) {
    attributes.push(` ${attribute.name}="${escapeXmlAttribute(attribute.value)}"`);
  }
  const inner = columnFilterXml(prefix, filter);
  if (inner === '') return `<${name}${attributes.join('')}/>`;
  return `<${name}${attributes.join('')}>${inner}</${name}>`;
}

function columnFilterXml(prefix: string, filter: ColumnFilter): string {
  switch (filter.kind) {
    case 'values': {
      const name = qualified(prefix, 'filters');
      const filterName = qualified(prefix, 'filter');
      const attributes = filter.includeBlanks ? ' blank="1"' : '';
      if (filter.selected.length === 0) return `<${name}${attributes}/>`;
      return `<${name}${attributes}>${filter.selected
        .map((value) => `<${filterName} val="${escapeXmlAttribute(formatCellValue(value))}"/>`)
        .join('')}</${name}>`;
    }
    case 'conditions':
      return customFiltersXml(prefix, filter);
    case 'unsupported':
      return isUnsupportedFilterSource(filter.raw) ? filter.raw.inner : '';
  }
}

function customFiltersXml(prefix: string, filter: ConditionsFilter): string {
  const name = qualified(prefix, 'customFilters');
  const operatorName = qualified(prefix, 'customFilter');
  const operators: { operator: string; value: string }[] = [];
  for (const condition of filter.conditions) operators.push(...operatorsOf(condition));
  const and = betweenJoin(filter) ?? (filter.join === 'and' ? '1' : '0');
  return `<${name} and="${and}">${operators
    .map(
      (operator) =>
        `<${operatorName} operator="${operator.operator}" val="${escapeXmlAttribute(
          operator.value,
        )}"/>`,
    )
    .join('')}</${name}>`;
}

/** `between` is written with AND, `not between` with OR, whatever the column join says. */
function betweenJoin(filter: ConditionsFilter): '1' | '0' | null {
  for (const condition of filter.conditions) {
    if ('from' in condition) return condition.operator === 'between' ? '1' : '0';
  }
  return null;
}

function operatorsOf(condition: Condition): { operator: string; value: string }[] {
  switch (condition.kind) {
    case 'text':
      return textOperators(condition);
    case 'number':
      if ('from' in condition) {
        const low = plainNumber(Math.min(condition.from, condition.to));
        const high = plainNumber(Math.max(condition.from, condition.to));
        return condition.operator === 'between'
          ? [
              { operator: 'greaterThanOrEqual', value: low },
              { operator: 'lessThanOrEqual', value: high },
            ]
          : [
              { operator: 'lessThan', value: low },
              { operator: 'greaterThan', value: high },
            ];
      }
      return [
        { operator: comparisonOperator(condition.operator), value: plainNumber(condition.value) },
      ];
    case 'date':
      if ('from' in condition) {
        const low = plainNumber(Math.min(condition.from, condition.to));
        const high = plainNumber(Math.max(condition.from, condition.to));
        return condition.operator === 'between'
          ? [
              { operator: 'greaterThanOrEqual', value: low },
              { operator: 'lessThanOrEqual', value: high },
            ]
          : [
              { operator: 'lessThan', value: low },
              { operator: 'greaterThan', value: high },
            ];
      }
      return [dateOperator(condition.operator, plainNumber(condition.value))];
  }
}

function textOperators(condition: Extract<Condition, { kind: 'text' }>): {
  operator: string;
  value: string;
}[] {
  switch (condition.operator) {
    case 'equals':
      return [{ operator: 'equal', value: condition.value }];
    case 'not-equals':
      return [{ operator: 'notEqual', value: condition.value }];
    case 'begins-with':
      return [{ operator: 'equal', value: `${condition.value}*` }];
    case 'ends-with':
      return [{ operator: 'equal', value: `*${condition.value}` }];
    case 'contains':
      return [{ operator: 'equal', value: `*${condition.value}*` }];
    case 'does-not-contain':
      return [{ operator: 'notEqual', value: `*${condition.value}*` }];
  }
}

function comparisonOperator(operator: NumberComparisonOperator): string {
  switch (operator) {
    case '=':
      return 'equal';
    case '<>':
      return 'notEqual';
    case '>':
      return 'greaterThan';
    case '>=':
      return 'greaterThanOrEqual';
    case '<':
      return 'lessThan';
    case '<=':
      return 'lessThanOrEqual';
  }
}

function dateOperator(
  operator: 'equals' | 'not-equals' | 'before' | 'after',
  value: string,
): { operator: string; value: string } {
  switch (operator) {
    case 'equals':
      return { operator: 'equal', value };
    case 'not-equals':
      return { operator: 'notEqual', value };
    case 'before':
      return { operator: 'lessThan', value };
    case 'after':
      return { operator: 'greaterThan', value };
  }
}

// ---------------------------------------------------------------------------
// Save planning
// ---------------------------------------------------------------------------

export interface AutoFilterSavePlan {
  /** `true` when the worksheet part must be written for the filter. */
  readonly changed: boolean;
  /** The element to write, or `null` to remove the sheet's filter. */
  readonly xml: string | null;
}

/**
 * Decide what the save must do with one sheet's filter. A filter parsed from
 * the file and left alone is not a change, so a no-op save keeps the part's
 * bytes (§3.6). A filter the model created or changed is rebuilt.
 */
export function planAutoFilterSave(
  filter: AutoFilter | null,
  sourceBytes: Uint8Array | null,
): AutoFilterSavePlan {
  if (filter !== null) {
    const source = isAutoFilterRawSource(filter.raw) ? filter.raw : null;
    if (source !== null && source.stateKey === autoFilterStateKey(filter)) {
      return { changed: false, xml: source.xml };
    }
    return { changed: true, xml: serialiseAutoFilterElement(filter) };
  }
  const present = sourceBytes !== null && sourceHasAutoFilterElement(sourceBytes);
  return { changed: present, xml: null };
}

const AUTO_FILTER_BYTES = new TextEncoder().encode('autoFilter');

/** A byte scan: does this worksheet part carry an `autoFilter` element? */
function sourceHasAutoFilterElement(bytes: Uint8Array): boolean {
  for (let index = 0; index + AUTO_FILTER_BYTES.length < bytes.length; index += 1) {
    if (bytes[index] !== 0x3c) continue; // '<'
    let cursor = index + 1;
    let name = '';
    while (cursor < bytes.length && isNameByte(bytes[cursor]!)) {
      name += String.fromCharCode(bytes[cursor]!);
      cursor += 1;
    }
    if (name.length === 0) continue;
    const local = name.slice(name.lastIndexOf(':') + 1);
    if (local === 'autoFilter') return true;
    index = cursor - 1;
  }
  return false;
}

function isNameByte(value: number): boolean {
  return (
    (value >= 0x41 && value <= 0x5a) ||
    (value >= 0x61 && value <= 0x7a) ||
    (value >= 0x30 && value <= 0x39) ||
    value === 0x5f ||
    value === 0x2d ||
    value === 0x2e ||
    value === 0x3a
  );
}

// ---------------------------------------------------------------------------
// Open-side visibility rebuild (§3.3)
// ---------------------------------------------------------------------------

/**
 * Rebuild the two visibility layers from the saved effective flags (§3.3).
 * `restoreHiddenLayers` splits the imported hidden rows into the filter layer
 * and the manual layer; a row that fails the saved criteria is hidden even if
 * the file forgot to flag it, so the grid and the file agree.
 */
export function applyAutoFilterHiddenLayers(sheet: Sheet): void {
  const filter = sheet.autoFilter;
  if (filter === null) return;
  const savedHidden: number[] = [];
  for (const [row, dimension] of sheet.rowDimensions) {
    if (dimension.hidden === true) savedHidden.push(row);
  }
  const layers = restoreHiddenLayers(sheet, filter, savedHidden);
  for (const row of layers.filterHidden) {
    const dimension = sheet.rowDimensions.get(row);
    if (dimension?.hidden === true) continue;
    sheet.rowDimensions.set(row, { ...(dimension ?? {}), hidden: true });
  }
}

/** Rows the sheet's active filter hides; empty without one. */
export function filterHiddenRows(sheet: Sheet): ReadonlySet<number> {
  const filter = sheet.autoFilter;
  if (filter === null) return new Set<number>();
  return evaluateAutoFilter(sheet, filter).filterHidden;
}

/**
 * The hidden-flag changes one changed filter needs, per §3.3. The source part
 * records only the effective flag, so the manual layer is recovered from the
 * source criteria: failing rows were the filter layer, every other hidden row
 * was manual. The model then supplies the new filter layer, and a row whose
 * saved flag no longer matches is reported with the flag it must carry.
 */
export function filterHiddenFlagChanges(sheet: Sheet, sourceSheet: Sheet): Map<number, boolean> {
  const sourceHidden = new Set<number>();
  for (const [row, dimension] of sourceSheet.rowDimensions) {
    if (dimension.hidden === true) sourceHidden.add(row);
  }
  const manual = restoreHiddenLayers(
    sourceSheet,
    sourceSheet.autoFilter,
    sourceHidden,
  ).manualHidden;
  const next = filterHiddenRows(sheet);
  const changes = new Map<number, boolean>();
  for (const row of new Set<number>([...sourceHidden, ...manual, ...next])) {
    const hidden = manual.has(row) || next.has(row);
    if (sourceHidden.has(row) !== hidden) changes.set(row, hidden);
  }
  return changes;
}

/**
 * The effective hidden flag a save writes is `manual || filter` (§3.3), so a
 * row this sheet is already rewriting is marked hidden when the filter hides
 * it. Rows the save does not rewrite keep their source bytes.
 */
export function withFilterHiddenRows(
  sheet: Sheet,
  rows: ReadonlyMap<number, RowDimension | null>,
): ReadonlyMap<number, RowDimension | null> {
  const hidden = filterHiddenRows(sheet);
  if (hidden.size === 0) return rows;
  const merged = new Map(rows);
  for (const row of hidden) {
    const current = merged.has(row) ? merged.get(row) : (sheet.rowDimensions.get(row) ?? null);
    if (current?.hidden === true) continue;
    merged.set(row, { ...(current ?? {}), hidden: true });
  }
  return merged;
}

// ---------------------------------------------------------------------------
// The hidden `_xlnm._FilterDatabase` defined name (§3.6)
// ---------------------------------------------------------------------------

export const FILTER_DATABASE_NAME = '_xlnm._FilterDatabase';

/**
 * Direct children of `<worksheet>` that follow `autoFilter` in the schema
 * order (CT_Worksheet). A new filter element goes before the first of them.
 */
export const AFTER_AUTO_FILTER: ReadonlySet<string> = new Set([
  'sortState',
  'dataConsolidate',
  'customSheetViews',
  'mergeCells',
  'phoneticPr',
  'conditionalFormatting',
  'dataValidations',
  'hyperlinks',
  'printOptions',
  'pageMargins',
  'pageSetup',
  'headerFooter',
  'rowBreaks',
  'colBreaks',
  'customProperties',
  'cellWatches',
  'ignoredErrors',
  'smartTags',
  'drawing',
  'legacyDrawing',
  'legacyDrawingHF',
  'picture',
  'oleObjects',
  'controls',
  'webPublishItems',
  'tableParts',
  'extLst',
]);

export function isFilterDatabaseName(name: string): boolean {
  return name.toLowerCase() === FILTER_DATABASE_NAME.toLowerCase();
}

/** `'Sheet 1'!$A$1:$E$5`, the system name's formula for a filter range. */
export function filterDatabaseNameText(sheetName: string, range: GridRange): string {
  const start = `$${columnToLetters(range.startCol)}$${range.startRow + 1}`;
  const end = `$${columnToLetters(range.endCol)}$${range.endRow + 1}`;
  return `'${sheetName.replaceAll("'", "''")}'!${start}:${end}`;
}

export interface FilterDatabasePatchInput {
  readonly bytes: Uint8Array;
  readonly partName: string;
  readonly workbook: Workbook;
  /** Sheets whose name this save maintains: every sheet with a filter, plus removals. */
  readonly sheetIds: ReadonlySet<SheetId>;
}

/**
 * Keep the hidden `_xlnm._FilterDatabase` name in step with the model's filter
 * ranges: create it when a filter is switched on, retarget it when the range
 * moves and remove it when the filter goes (§3.6). Sheet-scoped names are
 * addressed by their current workbook index. Returns `null` when nothing
 * changed, so an untouched save keeps the workbook part's bytes.
 */
export function patchFilterDatabaseNames(input: FilterDatabasePatchInput): Uint8Array | null {
  const xml = decodeXml(input.bytes);
  const parsed = parseDefinedNames(xml, input.partName);
  const desired = new Map<number, string>();
  input.workbook.sheets.forEach((sheet, index) => {
    if (!input.sheetIds.has(sheet.id)) return;
    if (sheet.autoFilter === null) return;
    desired.set(index, filterDatabaseNameText(sheet.name, sheet.autoFilter.range));
  });

  const edits: { start: number; end: number; text: string }[] = [];
  const handled = new Set<number>();
  for (const record of parsed.names) {
    if (!isFilterDatabaseName(record.name)) continue;
    const index = record.localSheetId === null ? Number.NaN : Number(record.localSheetId);
    if (!Number.isInteger(index) || index < 0 || index >= input.workbook.sheets.length) continue;
    const sheet = input.workbook.sheets[index]!;
    if (!input.sheetIds.has(sheet.id)) continue;
    const text = desired.get(index) ?? null;
    if (text === null) {
      edits.push({ start: record.start, end: record.end, text: '' });
      continue;
    }
    if (handled.has(index)) {
      edits.push({ start: record.start, end: record.end, text: '' });
      continue;
    }
    handled.add(index);
    const retargeted = retargetNameText(record.text, text);
    if (retargeted === record.text) continue;
    if (record.selfClosing) {
      const open = xml.slice(record.start, record.openEnd).replace(/\/\s*$/, '>');
      const closing = `</${qualified(record.prefix, 'definedName')}>`;
      edits.push({
        start: record.start,
        end: record.end,
        text: `${open}${escapeNameText(retargeted)}${closing}`,
      });
      continue;
    }
    edits.push({ start: record.openEnd, end: record.closeStart, text: escapeNameText(retargeted) });
  }

  const missing = [...desired.keys()].filter((index) => !handled.has(index)).sort((a, b) => a - b);
  const created =
    missing.length === 0
      ? null
      : missing
          .map((index) => {
            const prefix = parsed.container?.prefix ?? parsed.prefix;
            const name = qualified(prefix, 'definedName');
            return `<${name} name="${FILTER_DATABASE_NAME}" localSheetId="${index}" hidden="1">${escapeNameText(
              desired.get(index)!,
            )}</${name}>`;
          })
          .join('');

  if (created !== null) {
    const container = parsed.container;
    if (container === null) {
      const prefix = parsed.prefix;
      const containerTag = qualified(prefix, 'definedNames');
      edits.push({
        start: parsed.workbookClose,
        end: parsed.workbookClose,
        text: `<${containerTag}>${created}</${containerTag}>`,
      });
    } else if (container.selfClosing) {
      const open = xml.slice(container.start, container.openEnd).replace(/\/\s*$/, '>');
      const closing = `</${qualified(container.prefix, 'definedNames')}>`;
      edits.push({
        start: container.start,
        end: container.end,
        text: `${open}${created}${closing}`,
      });
    } else {
      edits.push({ start: container.contentEnd, end: container.contentEnd, text: created });
    }
  }

  if (edits.length === 0) return null;
  let patched = xml;
  for (const edit of [...edits].sort((left, right) => right.start - left.start)) {
    patched = patched.slice(0, edit.start) + edit.text + patched.slice(edit.end);
  }
  if (!isWellFormedXml(patched)) {
    throw new WorkbookSaveError(
      'unsafe-part',
      `The workbook part ${input.partName} could not store its AutoFilter name safely. The file was not written.`,
    );
  }
  return encodeXml(patched);
}

interface DefinedNameElement {
  readonly name: string;
  readonly localSheetId: string | null;
  readonly prefix: string;
  readonly start: number;
  readonly openEnd: number;
  readonly closeStart: number;
  readonly end: number;
  readonly selfClosing: boolean;
  readonly text: string;
}

interface DefinedNamesContainer {
  readonly prefix: string;
  readonly start: number;
  readonly openEnd: number;
  readonly contentEnd: number;
  readonly end: number;
  readonly selfClosing: boolean;
}

function parseDefinedNames(
  xml: string,
  partName: string,
): {
  readonly prefix: string;
  readonly workbookClose: number;
  readonly container: DefinedNamesContainer | null;
  readonly names: readonly DefinedNameElement[];
} {
  let prefix = '';
  let container: DefinedNamesContainer | null = null;
  const names: DefinedNameElement[] = [];
  let current: {
    name: string;
    localSheetId: string | null;
    prefix: string;
    start: number;
    openEnd: number;
    text: string;
  } | null = null;
  let depth = 0;

  parseXml(
    xml,
    {
      onOpenTag(tag, position) {
        depth += 1;
        const start = openTagStart(xml, position);
        if (depth === 1) {
          prefix = tag.prefix;
          return;
        }
        if (depth === 2 && tag.local === 'definedNames') {
          if (container === null) {
            container = {
              prefix: tag.prefix,
              start,
              openEnd: position,
              contentEnd: position,
              end: position,
              selfClosing: false,
            };
          }
          return;
        }
        if (depth === 3 && tag.local === 'definedName') {
          current = {
            name: attributeValueOf(attributesOf(tag), 'name') ?? '',
            localSheetId: attributeValueOf(attributesOf(tag), 'localSheetId') ?? null,
            prefix: tag.prefix,
            start,
            openEnd: position,
            text: '',
          };
        }
      },
      onCloseTag(tag, start, end) {
        depth -= 1;
        if (depth === 1 && tag.local === 'definedNames' && container !== null) {
          const open = container as DefinedNamesContainer;
          container = {
            ...open,
            contentEnd: start,
            end,
            selfClosing: open.openEnd === end,
          };
          return;
        }
        if (depth === 2 && tag.local === 'definedName' && current !== null) {
          names.push({
            ...current,
            closeStart: start,
            end,
            selfClosing: current.openEnd === end,
          });
          current = null;
        }
      },
      onText(text) {
        if (current !== null) current.text += text;
      },
      onCdata(text) {
        if (current !== null) current.text += text;
      },
    },
    partName,
  );

  const closing = xml.lastIndexOf('</');
  return { prefix, workbookClose: closing < 0 ? xml.length : closing, container, names };
}

/** Compare name formulas without the `$`, quoting or case noise Excel allows. */
function normaliseNameText(text: string): string {
  return text.replace(/\$/g, '').replace(/'/g, '').replace(/\s+/g, '').toLowerCase();
}

/**
 * Keep the name's authored sheet reference — quoting included — and replace
 * only the range, so retargeting a filter changes as few bytes as possible.
 * A different sheet is written in the canonical quoted form.
 */
function retargetNameText(existing: string, desired: string): string {
  const existingSeparator = existing.lastIndexOf('!');
  const desiredSeparator = desired.lastIndexOf('!');
  if (existingSeparator < 0 || desiredSeparator < 0) return desired;
  const existingPrefix = existing.slice(0, existingSeparator);
  const desiredPrefix = desired.slice(0, desiredSeparator);
  if (normaliseNameText(existingPrefix) !== normaliseNameText(desiredPrefix)) return desired;
  return `${existingPrefix}${desired.slice(desiredSeparator)}`;
}

/**
 * Element text needs only `&`, `<` and `>` escaped; leaving apostrophes and
 * quotes literal keeps the name looking like the one Excel writes.
 */
function escapeNameText(text: string): string {
  return text.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;');
}

// ---------------------------------------------------------------------------
// XML helpers
// ---------------------------------------------------------------------------

function parseElementFragment(xml: string, partName: string): XmlFragmentElement | null {
  // A prefixed element cannot be parsed on its own: its namespace binding sits
  // on the worksheet root. Wrap the fragment with declarations for every
  // prefix it uses, then report positions back in the fragment's own indexes.
  const declarations = fragmentNamespaceDeclarations(xml);
  const wrapper = declarations === '' ? '' : `<calcplus-fragment${declarations}>`;
  const source = wrapper === '' ? xml : `${wrapper}${xml}</calcplus-fragment>`;
  const offset = wrapper.length;

  let root: XmlFragmentElement | null = null;
  const stack: {
    local: string;
    prefix: string;
    attributes: XmlAttributeRecord[];
    elements: XmlFragmentElement[];
    start: number;
    openEnd: number;
  }[] = [];

  parseXml(
    source,
    {
      onOpenTag(tag, position) {
        stack.push({
          local: tag.local,
          prefix: tag.prefix,
          attributes: attributesOf(tag),
          elements: [],
          start: openTagStart(source, position) - offset,
          openEnd: position - offset,
        });
      },
      onCloseTag(_tag, _start, end) {
        const open = stack.pop();
        if (open === undefined) return;
        const element: XmlFragmentElement = {
          local: open.local,
          prefix: open.prefix,
          attributes: open.attributes,
          elements: open.elements,
          start: open.start,
          end: end - offset,
          openEnd: open.openEnd,
        };
        const parent = stack.at(-1);
        if (parent === undefined) root = element;
        else parent.elements.push(element);
      },
    },
    partName,
  );
  if (root === null) return null;
  const rootElement = root as XmlFragmentElement;
  return wrapper === '' ? rootElement : (rootElement.elements[0] ?? null);
}

/** Declarations for every prefix the fragment uses in a tag or attribute name. */
function fragmentNamespaceDeclarations(xml: string): string {
  const prefixes = new Set<string>();
  for (const pattern of [/<\s*([A-Za-z_][\w.-]*):/g, /\s([A-Za-z_][\w.-]*):[\w.-]+\s*=/g]) {
    pattern.lastIndex = 0;
    let match: RegExpExecArray | null;
    while ((match = pattern.exec(xml)) !== null) {
      const prefix = match[1]!;
      if (prefix !== 'xmlns' && prefix !== 'xml') prefixes.add(prefix);
    }
  }
  if (prefixes.size === 0) return '';
  return [...prefixes]
    .map((prefix) => ` xmlns:${prefix}="urn:calcplus:fragment:${prefix}"`)
    .join('');
}

function attributesOf(tag: SaxesTagNS): XmlAttributeRecord[] {
  const attributes: XmlAttributeRecord[] = [];
  for (const attribute of Object.values(tag.attributes) as SaxesAttributeNS[]) {
    // Namespace declarations belong to the enclosing worksheet element; a
    // rebuilt filter must never repeat them.
    if (/^xmlns(:|$)/.test(attribute.name)) continue;
    attributes.push({ name: attribute.name, value: attribute.value });
  }
  return attributes;
}

function attributeValueOf(
  attributes: readonly XmlAttributeRecord[],
  local: string,
): string | undefined {
  for (const attribute of attributes) {
    if (localName(attribute.name) === local) return attribute.value;
  }
  return undefined;
}

function localName(name: string): string {
  const separator = name.lastIndexOf(':');
  return separator < 0 ? name : name.slice(separator + 1);
}

function qualified(prefix: string, local: string): string {
  return prefix === '' ? local : `${prefix}:${local}`;
}

function rangeOfAttributes(attributes: readonly XmlAttributeRecord[]): GridRange {
  const ref = attributeValueOf(attributes, 'ref');
  if (ref !== undefined) {
    try {
      return parseRangeRef(ref);
    } catch {
      // A malformed ref keeps the element in the model so its XML is preserved
      // on save; the filter then covers the first cell only.
    }
  }
  return singleCellRange(0, 0);
}

function formatRange(range: GridRange): string {
  const start = `${columnToLetters(range.startCol)}${range.startRow + 1}`;
  const end = `${columnToLetters(range.endCol)}${range.endRow + 1}`;
  return `${start}:${end}`;
}

function plainNumber(value: number): string {
  return Object.is(value, -0) ? '0' : String(value);
}
