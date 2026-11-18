import type { GridRange, Sheet, SheetId, Workbook } from '@model/types';
import { columnToLetters, singleCellRange } from '@model/types';
import type {
  CfCellOperator,
  CfDifferentialFormat,
  CfDxfResolver,
  CfIconSetId,
  CfOperand,
  CfTextOperator,
  CfThreshold,
  CfTimePeriod,
  CfUnsupportedReason,
  ConditionalFormatRule,
} from '@model/conditional-format';
import { CF_ICON_SETS } from '@model/conditional-format';
import { resolveColor, underlineFromValue } from '@model/style';
import type { ConditionalFormatRecord } from '@model/types';
import { WorkbookSaveError } from './errors';
import { decodeXml, escapeXmlAttribute, openTagStart, parseXml, attributeValue } from './xml';

/**
 * Conditional formatting in a worksheet part (conditional formatting §7).
 *
 * Reading maps both the base `<conditionalFormatting>` elements and the
 * `x14:conditionalFormatting` extension block onto `ConditionalFormatRule`s.
 * The two forms of one rule are merged by their shared rule id, the x14 payload
 * driving rendering while both XML forms stay preserved. This module owns the
 * XML only: it never evaluates a rule and never mutates the model.
 *
 * Byte identity: each base element is kept verbatim in
 * `Sheet.conditionalFormats[].rawXml`, and a no-op save re-emits the source
 * part untouched. A changed record writes the §7 shape and keeps every
 * unmodelled attribute and child in its authored order. The x14 block is never
 * rewritten in this milestone (there is no rule authoring), so it survives
 * byte-for-byte; its ranges come from the base rule it duplicates.
 */

// ---------------------------------------------------------------------------
// Verbatim source records
// ---------------------------------------------------------------------------

/** One XML attribute, name and value exactly as authored. */
export interface XmlAttributeRecord {
  readonly name: string;
  readonly value: string;
}

/** One authored `cfRule` element, from either XML form. */
export interface CfRuleXmlSource {
  readonly source: 'base' | 'x14';
  /** The exact `cfRule` element (`<cfRule>` or `<x14:cfRule>`). */
  readonly xml: string;
  /** The authored `type` attribute, even when it is unknown. */
  readonly ruleType: string;
  /** The GUID that links an x14 rule to its base copy, when authored. */
  readonly ruleId: string | null;
  /** The enclosing element's `sqref` text; empty for an x14 rule. */
  readonly sqref: string;
}

/**
 * The parsed source of one rule, as stored in `ConditionalFormatRule.raw`.
 * `base` and `x14` are both present when the two XML forms describe the same
 * rule; the x14 form drives the payload (§7).
 */
export interface ConditionalFormatRuleSource {
  readonly kind: 'cfRule-source';
  readonly base: CfRuleXmlSource | null;
  readonly x14: CfRuleXmlSource | null;
  /** The modelled payload at parse time; a change means the rule was re-authored. */
  readonly payloadKey: string;
  /** The rule's own state at parse time; equality means the source still applies. */
  readonly stateKey: string;
}

/** The parsed source of one base `<conditionalFormatting>` element. */
export interface ConditionalFormatElementSource {
  readonly kind: 'conditionalFormatting-source';
  /** The exact source element, emitted verbatim on an untouched save. */
  readonly xml: string;
  readonly prefix: string;
  readonly attributes: readonly XmlAttributeRecord[];
  readonly sqref: string;
  readonly ranges: readonly GridRange[];
  /** Element and rule state at parse time; equality means the source still applies. */
  readonly stateKey: string;
}

/** One base element with the record it becomes and its own rules. */
export interface ParsedConditionalFormatElement {
  readonly record: ConditionalFormatRecord;
  readonly element: ConditionalFormatElementSource;
  /** The element's own rules, in document order (never the x14-merged set). */
  readonly rules: readonly ConditionalFormatRule[];
}

export function isConditionalFormatRuleSource(raw: unknown): raw is ConditionalFormatRuleSource {
  if (raw === null || typeof raw !== 'object') return false;
  const candidate = raw as Partial<ConditionalFormatRuleSource>;
  return (
    candidate.kind === 'cfRule-source' &&
    typeof candidate.stateKey === 'string' &&
    typeof candidate.payloadKey === 'string' &&
    (candidate.base === null || typeof candidate.base === 'object') &&
    (candidate.x14 === null || typeof candidate.x14 === 'object')
  );
}

export function isConditionalFormatElementSource(
  raw: unknown,
): raw is ConditionalFormatElementSource {
  if (raw === null || typeof raw !== 'object') return false;
  const candidate = raw as Partial<ConditionalFormatElementSource>;
  return (
    candidate.kind === 'conditionalFormatting-source' &&
    typeof candidate.xml === 'string' &&
    typeof candidate.stateKey === 'string' &&
    Array.isArray(candidate.attributes)
  );
}

// ---------------------------------------------------------------------------
// Parsing
// ---------------------------------------------------------------------------

export interface ConditionalFormatParseOptions {
  /** The workbook's `<indexedColors>` palette, for authored colours. */
  readonly indexedPalette?: readonly string[] | null;
}

export interface ConditionalFormatParse {
  readonly records: readonly ConditionalFormatRecord[];
  readonly rules: readonly ConditionalFormatRule[];
  readonly elements: readonly ParsedConditionalFormatElement[];
}

interface XmlFragment {
  readonly local: string;
  readonly prefix: string;
  readonly attributes: readonly XmlAttributeRecord[];
  readonly elements: readonly XmlFragment[];
  readonly text: string;
  readonly start: number;
  readonly end: number;
  readonly openEnd: number;
  readonly selfClosing: boolean;
}

interface ElementSpan {
  start: number;
  end: number;
}

interface DirectChild {
  readonly local: string;
  readonly start: number;
  end: number;
}

interface WorksheetScan {
  readonly base: readonly ElementSpan[];
  readonly x14: readonly ElementSpan[];
  readonly directChildren: readonly DirectChild[];
  readonly rootCloseStart: number;
}

/**
 * Locate the base `conditionalFormatting` elements (direct children of the
 * worksheet) and every prefixed `conditionalFormatting` extension element.
 * One pass, no model work: the save plan and the open path both need it.
 */
function scanWorksheet(xml: string, partName: string): WorksheetScan {
  const base: ElementSpan[] = [];
  const x14: ElementSpan[] = [];
  const directChildren: DirectChild[] = [];
  const frames: {
    span: ElementSpan | null;
    direct: DirectChild | null;
  }[] = [];
  // Prefixes may be declared on an ancestor (the worksheet root), so the scan
  // wraps the fragment with declarations and reports positions back in the
  // fragment's own indexes. An XML declaration stays outside the wrapper.
  const declarationEnd = xml.startsWith('<?xml') ? xml.indexOf('?>') + 2 : 0;
  const body = xml.slice(declarationEnd);
  const declarations = fragmentNamespaceDeclarations(body);
  const wrapper = declarations === '' ? '' : `<calcplus-fragment${declarations}>`;
  const source =
    wrapper === '' ? xml : `${xml.slice(0, declarationEnd)}${wrapper}${body}</calcplus-fragment>`;
  const offset = wrapper.length;
  const rootLevel = wrapper === '' ? 1 : 2;
  const childLevel = rootLevel + 1;
  let rootCloseStart = xml.lastIndexOf('</');
  if (rootCloseStart < 0) rootCloseStart = xml.length;

  parseXml(
    source,
    {
      onOpenTag(tag, position) {
        const start = openTagStart(source, position) - offset;
        const openEnd = position - offset;
        const selfClosing = /\/\s*>$/.test(source.slice(openTagStart(source, position), position));
        const level = frames.length + 1;
        let span: ElementSpan | null = null;
        if (tag.prefix !== '' && tag.local === 'conditionalFormatting') {
          span = { start, end: openEnd };
          x14.push(span);
        }
        let direct: DirectChild | null = null;
        if (level === childLevel) {
          direct = { local: tag.local, start, end: openEnd };
          directChildren.push(direct);
          if (tag.local === 'conditionalFormatting' && tag.prefix === '') {
            const baseSpan: ElementSpan = { start, end: openEnd };
            base.push(baseSpan);
            span = baseSpan;
          }
        }
        if (selfClosing && direct !== null) direct.end = openEnd;
        frames.push({ span, direct });
      },
      onCloseTag(_tag, start, end) {
        const frame = frames.pop();
        if (frame === undefined) return;
        if (frame.direct !== null) frame.direct.end = end - offset;
        if (frame.span !== null) frame.span.end = end - offset;
        if (frames.length === rootLevel - 1) rootCloseStart = start - offset;
      },
    },
    partName,
  );
  return { base, x14, directChildren, rootCloseStart };
}

/**
 * Parse every conditional-format element of a worksheet part into records and
 * rules (§7). Records carry the base elements verbatim; rules merge the base and
 * x14 forms by their shared id, with the x14 payload winning.
 */
export function parseConditionalFormats(
  xml: string,
  sheet: Sheet,
  options: ConditionalFormatParseOptions = {},
): ConditionalFormatParse {
  const scan = scanWorksheet(xml, sheet.partPath);
  const elements = scan.base.flatMap((span) => {
    const parsed = parseBaseElement(xml.slice(span.start, span.end), sheet, options);
    return parsed === null ? [] : [parsed];
  });
  const x14Rules = scan.x14.flatMap((span) =>
    parseX14Rules(xml.slice(span.start, span.end), sheet, options),
  );
  return {
    records: elements.map((entry) => entry.record),
    rules: mergeConditionalFormatRules(sheet, elements, x14Rules),
    elements,
  };
}

/**
 * Populate `Sheet.conditionalFormats` and return the merged rules. The open path
 * already holds the worksheet's direct children, so the x14 block is read from
 * its own slice instead of re-parsing the part.
 */
export function applyConditionalFormats(
  sheet: Sheet,
  xml: string,
  rootChildren: readonly { readonly local: string; readonly start: number; readonly end: number }[],
  options: ConditionalFormatParseOptions = {},
): readonly ConditionalFormatRule[] {
  const elements: ParsedConditionalFormatElement[] = [];
  for (const child of rootChildren) {
    if (child.local !== 'conditionalFormatting') continue;
    const parsed = parseBaseElement(xml.slice(child.start, child.end), sheet, options);
    if (parsed !== null) elements.push(parsed);
  }
  const x14Rules: ConditionalFormatRule[] = [];
  const extLst = rootChildren.find((child) => child.local === 'extLst');
  if (extLst !== undefined) {
    const slice = xml.slice(extLst.start, extLst.end);
    for (const span of scanWorksheet(slice, `${sheet.partPath} (extLst)`).x14) {
      x14Rules.push(...parseX14Rules(slice.slice(span.start, span.end), sheet, options));
    }
  }
  const records = elements.map((entry) => entry.record);
  sheet.conditionalFormats.splice(0, sheet.conditionalFormats.length, ...records);
  return mergeConditionalFormatRules(sheet, elements, x14Rules);
}

/**
 * The sheet's rules rebuilt from the model records plus the x14 block of the
 * last saved part. Used after a snapshot restore, where the records survive but
 * the open-time rule list does not. The records are authoritative for the base
 * rules; the source part supplies the x14 payloads.
 */
export function conditionalFormatRulesFor(
  sheet: Sheet,
  sourceBytes: Uint8Array | null,
  options: ConditionalFormatParseOptions = {},
): readonly ConditionalFormatRule[] {
  const elements = sheet.conditionalFormats.flatMap((record) => {
    const parsed = parseBaseElement(record.rawXml, sheet, options);
    if (parsed === null) return [];
    // The record, not the re-parse, owns the range: a structural edit may have
    // moved it without rewriting the element yet.
    return [{ ...parsed, record }];
  });
  const x14Rules: ConditionalFormatRule[] = [];
  if (sourceBytes !== null) {
    const xml = decodeXml(sourceBytes);
    const scan = scanWorksheet(xml, sheet.partPath);
    for (const span of scan.x14) {
      x14Rules.push(...parseX14Rules(xml.slice(span.start, span.end), sheet, options));
    }
  }
  return mergeConditionalFormatRules(sheet, elements, x14Rules);
}

function parseBaseElement(
  elementXml: string,
  sheet: Sheet,
  options: ConditionalFormatParseOptions,
): ParsedConditionalFormatElement | null {
  const element = parseFragment(elementXml, `${sheet.partPath} (conditionalFormatting)`);
  if (element === null || element.local !== 'conditionalFormatting') return null;
  const sqref = attributeValueOf(element.attributes, 'sqref') ?? '';
  const ranges = parseSqref(sqref);
  const rules: ConditionalFormatRule[] = [];
  element.elements.forEach((child) => {
    if (child.local !== 'cfRule') return;
    rules.push(
      buildRule({
        sheet,
        payload: basePayload(child, options),
        attributes: child.attributes,
        raw: {
          source: 'base',
          xml: elementXml.slice(child.start, child.end),
          ruleType: attributeValueOf(child.attributes, 'type') ?? '',
          ruleId: baseRuleId(child),
          sqref,
        },
        base: null,
        ranges,
      }),
    );
  });
  const attributes = element.attributes;
  return {
    record: { range: boundingRange(ranges), rawXml: elementXml },
    element: {
      kind: 'conditionalFormatting-source',
      xml: elementXml,
      prefix: element.prefix,
      attributes,
      sqref,
      ranges,
      stateKey: conditionalFormatElementStateKey(sqref, ranges, rules),
    },
    rules,
  };
}

/**
 * Rules in the `x14:conditionalFormattings` block. An x14 rule carries a GUID
 * that its base copy repeats, so the ranges come from the base rule; a rule the
 * base form does not describe keeps its own `sqref` when a producer wrote one.
 */
function parseX14Rules(
  elementXml: string,
  sheet: Sheet,
  options: ConditionalFormatParseOptions,
): ConditionalFormatRule[] {
  const element = parseFragment(elementXml, 'x14:conditionalFormatting');
  if (element === null || element.local !== 'conditionalFormatting') return [];
  const rules: ConditionalFormatRule[] = [];
  element.elements.forEach((child) => {
    if (child.local !== 'cfRule') return;
    const sqref = attributeValueOf(element.attributes, 'sqref') ?? '';
    rules.push(
      buildRule({
        sheet,
        payload: x14Payload(child, options),
        attributes: child.attributes,
        raw: {
          source: 'x14',
          xml: elementXml.slice(child.start, child.end),
          ruleType: attributeValueOf(child.attributes, 'type') ?? '',
          ruleId: attributeValueOf(child.attributes, 'id') ?? null,
          sqref,
        },
        base: null,
        ranges: parseSqref(sqref),
      }),
    );
  });
  return rules;
}

interface BuildRuleInput {
  readonly sheet: Sheet;
  readonly payload: CfRulePayload;
  readonly attributes: readonly XmlAttributeRecord[];
  readonly raw: CfRuleXmlSource;
  /** The base form of this rule, when the x14 form drives it. */
  readonly base: CfRuleXmlSource | null;
  readonly ranges: readonly GridRange[];
}

/**
 * Merge the base and x14 forms (§7). The shared GUID decides duplicates; a file
 * without GUIDs falls back to same priority plus the same rule family, which is
 * the only pairing Excel itself can mean for one data bar or icon set.
 */
function mergeConditionalFormatRules(
  sheet: Sheet,
  elements: readonly ParsedConditionalFormatElement[],
  x14Rules: readonly ConditionalFormatRule[],
): ConditionalFormatRule[] {
  const baseRules: ConditionalFormatRule[] = elements.flatMap((entry) => [...entry.rules]);
  const merged: ConditionalFormatRule[] = [];
  const consumed = new Set<ConditionalFormatRule>();
  for (const baseRule of baseRules) {
    const partner = x14Rules.find(
      (candidate) => !consumed.has(candidate) && sameRuleIdentity(baseRule, candidate),
    );
    if (partner === undefined) {
      merged.push({ ...baseRule, id: uniqueRuleId(baseRule, merged, sheet) });
      continue;
    }
    consumed.add(partner);
    merged.push(mergeRulePair(baseRule, partner, merged, sheet));
  }
  for (const x14Rule of x14Rules) {
    if (consumed.has(x14Rule)) continue;
    merged.push({ ...x14Rule, id: uniqueRuleId(x14Rule, merged, sheet) });
  }
  return merged;
}

function sameRuleIdentity(
  baseRule: ConditionalFormatRule,
  x14Rule: ConditionalFormatRule,
): boolean {
  if (ruleIdOf(baseRule) !== null && ruleIdOf(baseRule) === ruleIdOf(x14Rule)) return true;
  if (ruleIdOf(baseRule) !== null || ruleIdOf(x14Rule) !== null) return false;
  const type = ruleTypeOf(baseRule);
  if (type !== ruleTypeOf(x14Rule) || (type !== 'dataBar' && type !== 'iconSet')) return false;
  return (
    baseRule.priority === x14Rule.priority ||
    baseRule.priority === undefined ||
    x14Rule.priority === undefined
  );
}

function mergeRulePair(
  baseRule: ConditionalFormatRule,
  x14Rule: ConditionalFormatRule,
  merged: readonly ConditionalFormatRule[],
  sheet: Sheet,
): ConditionalFormatRule {
  const raw: ConditionalFormatRuleSource = {
    kind: 'cfRule-source',
    base: isConditionalFormatRuleSource(baseRule.raw) ? baseRule.raw.base : null,
    x14: isConditionalFormatRuleSource(x14Rule.raw) ? x14Rule.raw.x14 : null,
    payloadKey: '',
    stateKey: '',
  };
  const combined: ConditionalFormatRule = {
    ...x14Rule,
    id: uniqueRuleId(x14Rule, merged, sheet),
    priority: x14Rule.priority ?? baseRule.priority,
    stopIfTrue: x14Rule.stopIfTrue || baseRule.stopIfTrue,
    ranges: baseRule.ranges,
    ...(x14Rule.dxfId === undefined && baseRule.dxfId !== undefined
      ? { dxfId: baseRule.dxfId }
      : {}),
    raw,
  } as ConditionalFormatRule;
  return {
    ...combined,
    raw: {
      ...raw,
      payloadKey: cfRulePayloadKey(combined),
      stateKey: conditionalFormatRuleStateKey(combined),
    },
  } as ConditionalFormatRule;
}

function uniqueRuleId(
  rule: ConditionalFormatRule,
  existing: readonly ConditionalFormatRule[],
  sheet: Sheet,
): string {
  const authored = ruleIdOf(rule);
  if (authored !== null && !existing.some((candidate) => candidate.id === authored)) {
    return authored;
  }
  let index = existing.length;
  while (existing.some((candidate) => candidate.id === `${sheet.id}:cf:${index}`)) index += 1;
  return `${sheet.id}:cf:${index}`;
}

function ruleIdOf(rule: ConditionalFormatRule): string | null {
  const raw = rule.raw;
  if (!isConditionalFormatRuleSource(raw)) return null;
  return raw.base?.ruleId ?? raw.x14?.ruleId ?? null;
}

function ruleTypeOf(rule: ConditionalFormatRule): string {
  const raw = rule.raw;
  if (!isConditionalFormatRuleSource(raw)) return '';
  return raw.base?.ruleType ?? raw.x14?.ruleType ?? '';
}

function baseRuleId(element: XmlFragment): string | null {
  const authored = attributeValueOf(element.attributes, 'id');
  if (authored !== undefined) return authored;
  // Excel links the base rule to its x14 copy with an `x14:id` child.
  for (const child of element.elements) {
    if (child.local !== 'extLst') continue;
    for (const ext of child.elements) {
      for (const id of ext.elements) {
        if (id.local === 'id' && id.text.trim() !== '') return id.text.trim();
      }
    }
  }
  return null;
}

// ---------------------------------------------------------------------------
// Rule payloads (§1, §3)
// ---------------------------------------------------------------------------

type CfRulePayload =
  | {
      readonly kind: 'cellIs';
      readonly operator: CfCellOperator;
      readonly value1: CfOperand;
      readonly value2?: CfOperand;
    }
  | { readonly kind: 'text'; readonly operator: CfTextOperator; readonly text: string }
  | { readonly kind: 'timePeriod'; readonly mode: 'period'; readonly period: CfTimePeriod }
  | {
      readonly kind: 'top10';
      readonly rank?: number;
      readonly top: boolean;
      readonly percent: boolean;
    }
  | { readonly kind: 'duplicate' }
  | { readonly kind: 'unique' }
  | {
      readonly kind: 'colorScale';
      readonly thresholds: readonly CfThreshold[];
      readonly colors: readonly string[];
    }
  | {
      readonly kind: 'dataBar';
      readonly min: CfThreshold;
      readonly max: CfThreshold;
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
      readonly thresholds: readonly CfThreshold[];
      readonly reverse?: boolean;
      readonly showValue?: boolean;
      readonly iconOnly?: boolean;
    }
  | {
      readonly kind: 'unsupported';
      readonly reason: CfUnsupportedReason;
      readonly ruleType?: string;
    };

function buildRule(input: BuildRuleInput): ConditionalFormatRule {
  const attributes = input.attributes;
  const priority = integerAttribute(attributes, 'priority');
  const dxfId = integerAttribute(attributes, 'dxfId');
  const rule: ConditionalFormatRule = {
    id: '',
    sheetId: input.sheet.id,
    ranges: input.ranges,
    ...(priority === undefined ? {} : { priority }),
    stopIfTrue: xsdTrue(attributeValueOf(attributes, 'stopIfTrue')),
    ...(dxfId === undefined ? {} : { dxfId }),
    ...input.payload,
    raw: {
      kind: 'cfRule-source',
      base: input.base ?? (input.raw.source === 'base' ? input.raw : null),
      x14: input.raw.source === 'x14' ? input.raw : null,
      payloadKey: '',
      stateKey: '',
    },
  } as ConditionalFormatRule;
  return {
    ...rule,
    raw: {
      ...(rule.raw as ConditionalFormatRuleSource),
      payloadKey: cfRulePayloadKey(rule),
      stateKey: conditionalFormatRuleStateKey(rule),
    },
  } as ConditionalFormatRule;
}

function unsupported(reason: CfUnsupportedReason, ruleType: string): UnsupportedPayload {
  return { kind: 'unsupported', reason, ...(ruleType === '' ? {} : { ruleType }) };
}

/** Map a base `cfRule` onto the engine's kinds (§1, §3.3). */
function basePayload(rule: XmlFragment, options: ConditionalFormatParseOptions): CfRulePayload {
  const type = attributeValueOf(rule.attributes, 'type') ?? '';
  switch (type) {
    case 'cellIs':
      return cellIsPayload(rule);
    case 'containsText':
      return textPayload(rule, 'contains');
    case 'notContainsText':
      return textPayload(rule, 'notContains');
    case 'beginsWith':
      return textPayload(rule, 'beginsWith');
    case 'endsWith':
      return textPayload(rule, 'endsWith');
    case 'timePeriod':
      return timePeriodPayload(rule);
    case 'top10': {
      const rank = integerAttribute(rule.attributes, 'rank');
      return {
        kind: 'top10',
        ...(rank === undefined ? {} : { rank }),
        top: !xsdFalse(attributeValueOf(rule.attributes, 'top')),
        percent: xsdTrue(attributeValueOf(rule.attributes, 'percent')),
      };
    }
    case 'duplicateValues':
      return { kind: 'duplicate' };
    case 'uniqueValues':
      return { kind: 'unique' };
    case 'colorScale':
      return colorScalePayload(rule, options);
    case 'dataBar':
      return dataBarPayload(rule, options);
    case 'iconSet':
      return iconSetPayload(rule);
    case 'expression':
      return unsupported('formula', type);
    case 'aboveAverage':
      return unsupported('aboveAverage', type);
    case 'containsBlanks':
    case 'notContainsBlanks':
    case 'containsErrors':
    case 'notContainsErrors':
      return unsupported('blankOrError', type);
    default:
      return unsupported('unknownRuleType', type);
  }
}

/** Map an `x14:cfRule` onto the engine's kinds; only data bars and icon sets exist there. */
function x14Payload(rule: XmlFragment, options: ConditionalFormatParseOptions): CfRulePayload {
  const type = attributeValueOf(rule.attributes, 'type') ?? '';
  switch (type) {
    case 'dataBar':
      return x14DataBarPayload(rule, options);
    case 'iconSet':
      return x14IconSetPayload(rule);
    default:
      return unsupported('unknownRuleType', type);
  }
}

const CELL_OPERATORS: Readonly<Record<string, CfCellOperator>> = {
  lessThan: 'lessThan',
  lessThanOrEqual: 'lessThanOrEqual',
  equal: 'equal',
  notEqual: 'notEqual',
  greaterThanOrEqual: 'greaterThanOrEqual',
  greaterThan: 'greaterThan',
  between: 'between',
  notBetween: 'notBetween',
};

function cellIsPayload(rule: XmlFragment): CfRulePayload {
  const operatorText = attributeValueOf(rule.attributes, 'operator') ?? '';
  const operator = CELL_OPERATORS[operatorText];
  if (operator === undefined) return unsupported('unknownOperator', 'cellIs');
  const formulas = childTexts(rule, 'formula');
  const operands: CfOperand[] = [];
  for (const text of formulas) {
    const operand = operandOf(text);
    // A reference or function is a formula threshold: preserved, never guessed.
    if (operand === null) return unsupported('formulaThreshold', 'cellIs');
    operands.push(operand);
  }
  const first = operands[0];
  if (first === undefined) return unsupported('unknownRuleType', 'cellIs');
  if (operator === 'between' || operator === 'notBetween') {
    const second = operands[1];
    if (second === undefined) return unsupported('unknownRuleType', 'cellIs');
    return { kind: 'cellIs', operator, value1: first, value2: second };
  }
  return { kind: 'cellIs', operator, value1: first };
}

/** A literal operand. `null` when the text is a reference or an expression. */
function operandOf(text: string): CfOperand | null {
  let trimmed = text.trim();
  if (trimmed.startsWith('=')) trimmed = trimmed.slice(1).trim();
  if (trimmed === '') return null;
  if (/^(TRUE|FALSE)$/i.test(trimmed)) {
    return { type: 'boolean', value: trimmed.toUpperCase() === 'TRUE' };
  }
  if (/^"(?:[^"]|"")*"$/.test(trimmed)) {
    return { type: 'text', value: trimmed.slice(1, -1).replaceAll('""', '"') };
  }
  const numeric = Number(trimmed);
  if (Number.isFinite(numeric)) return { type: 'number', value: numeric };
  return null;
}

const TEXT_OPERATORS: Readonly<Record<string, CfTextOperator>> = {
  containsText: 'contains',
  notContains: 'notContains',
  beginsWith: 'beginsWith',
  endsWith: 'endsWith',
};

function textPayload(rule: XmlFragment, operator: CfTextOperator): CfRulePayload {
  const authored = attributeValueOf(rule.attributes, 'operator');
  if (authored !== undefined && TEXT_OPERATORS[authored] === undefined) {
    return unsupported('unknownOperator', 'containsText');
  }
  const text = attributeValueOf(rule.attributes, 'text');
  if (text === undefined) return unsupported('unknownRuleType', 'containsText');
  return { kind: 'text', operator, text };
}

const TIME_PERIODS: readonly CfTimePeriod[] = [
  'yesterday',
  'today',
  'tomorrow',
  'last7Days',
  'lastWeek',
  'thisWeek',
  'nextWeek',
  'lastMonth',
  'thisMonth',
  'nextMonth',
];

function timePeriodPayload(rule: XmlFragment): CfRulePayload {
  const authored = attributeValueOf(rule.attributes, 'timePeriod') ?? '';
  const period = TIME_PERIODS.find((candidate) => candidate === authored);
  if (period === undefined) return unsupported('unknownRuleType', 'timePeriod');
  return { kind: 'timePeriod', mode: 'period', period };
}

function childTexts(rule: XmlFragment, local: string): string[] {
  return rule.elements.filter((child) => child.local === local).map((child) => child.text.trim());
}

function child(rule: XmlFragment, local: string): XmlFragment | null {
  return rule.elements.find((candidate) => candidate.local === local) ?? null;
}

function colorHex(node: XmlFragment, options: ConditionalFormatParseOptions): string | null {
  const attributes = new Map<string, string>();
  for (const attribute of node.attributes) {
    attributes.set(localName(attribute.name), attribute.value);
  }
  return resolveColor(
    {
      rgb: attributes.get('rgb'),
      indexed: attributes.get('indexed'),
      theme: attributes.get('theme'),
      tint: attributes.get('tint'),
      auto: attributes.get('auto'),
    },
    options.indexedPalette ?? null,
  );
}

function colorScalePayload(
  rule: XmlFragment,
  options: ConditionalFormatParseOptions,
): CfRulePayload {
  const container = child(rule, 'colorScale');
  if (container === null) return unsupported('unknownRuleType', 'colorScale');
  const parsed = thresholdsOf(container);
  if (parsed.kind === 'unsupported') return parsed;
  const colors: string[] = [];
  for (const node of container.elements.filter((candidate) => candidate.local === 'color')) {
    const hex = colorHex(node, options);
    if (hex === null) return unsupported('unknownRuleType', 'colorScale');
    colors.push(hex);
  }
  if (parsed.thresholds.length < 2 || colors.length !== parsed.thresholds.length) {
    return unsupported('unknownRuleType', 'colorScale');
  }
  return { kind: 'colorScale', thresholds: parsed.thresholds, colors };
}

interface Thresholds {
  readonly kind: 'thresholds';
  readonly thresholds: readonly CfThreshold[];
}

type UnsupportedPayload = Extract<CfRulePayload, { kind: 'unsupported' }>;

function thresholdsOf(container: XmlFragment): Thresholds | UnsupportedPayload {
  const thresholds: CfThreshold[] = [];
  for (const node of container.elements.filter((candidate) => candidate.local === 'cfvo')) {
    const parsed = thresholdOf(node);
    if (parsed === 'formula') return unsupported('formulaThreshold', container.local);
    if (parsed === null) return unsupported('unknownRuleType', container.local);
    thresholds.push(parsed);
  }
  return { kind: 'thresholds', thresholds };
}

function thresholdOf(node: XmlFragment): CfThreshold | 'formula' | null {
  const type = attributeValueOf(node.attributes, 'type') ?? '';
  const valueText = attributeValueOf(node.attributes, 'val');
  switch (type) {
    case 'min':
    case 'max':
    case 'autoMin':
    case 'autoMax':
      return { type: 'automatic' };
    case 'num':
    case 'number': {
      const value = Number(valueText ?? '');
      return Number.isFinite(value) ? { type: 'number', value } : null;
    }
    case 'percent': {
      const value = Number(valueText ?? '');
      return Number.isFinite(value) ? { type: 'percent', value } : null;
    }
    case 'percentile': {
      const value = Number(valueText ?? '');
      return Number.isFinite(value) ? { type: 'percentile', value } : null;
    }
    case 'formula':
      return 'formula';
    default:
      return null;
  }
}

function dataBarPayload(rule: XmlFragment, options: ConditionalFormatParseOptions): CfRulePayload {
  const container = child(rule, 'dataBar');
  if (container === null) return unsupported('unknownRuleType', 'dataBar');
  const parsed = thresholdsOf(container);
  if (parsed.kind === 'unsupported') return parsed;
  const [min, max] = parsed.thresholds;
  const colorNode = container.elements.find((candidate) => candidate.local === 'color') ?? null;
  const color = colorNode === null ? null : colorHex(colorNode, options);
  if (min === undefined || max === undefined || color === null) {
    return unsupported('unknownRuleType', 'dataBar');
  }
  const showValue = attributeValueOf(container.attributes, 'showValue');
  return {
    kind: 'dataBar',
    min,
    max,
    color,
    fill: 'gradient',
    ...(showValue === undefined ? {} : { showValue: !xsdFalse(showValue) }),
  };
}

function x14DataBarPayload(
  rule: XmlFragment,
  options: ConditionalFormatParseOptions,
): CfRulePayload {
  const container = child(rule, 'dataBar');
  if (container === null) return unsupported('unknownRuleType', 'dataBar');
  const parsed = thresholdsOf(container);
  if (parsed.kind === 'unsupported') return parsed;
  const [min, max] = parsed.thresholds;
  const fillNode = container.elements.find((candidate) => candidate.local === 'fillColor') ?? null;
  const color = fillNode === null ? null : colorHex(fillNode, options);
  if (min === undefined || max === undefined || color === null) {
    return unsupported('unknownRuleType', 'dataBar');
  }
  const negativeNode =
    container.elements.find((candidate) => candidate.local === 'negativeFillColor') ?? null;
  const axisNode = container.elements.find((candidate) => candidate.local === 'axisColor') ?? null;
  const negativeColor = negativeNode === null ? null : colorHex(negativeNode, options);
  const axisColor = axisNode === null ? null : colorHex(axisNode, options);
  const gradient = attributeValueOf(container.attributes, 'gradient');
  const border = attributeValueOf(container.attributes, 'border');
  const showValue = attributeValueOf(container.attributes, 'showValue');
  const direction = attributeValueOf(container.attributes, 'direction');
  return {
    kind: 'dataBar',
    min,
    max,
    color,
    ...(negativeColor === null ? {} : { negativeColor }),
    ...(axisColor === null ? {} : { axisColor }),
    fill: gradient !== undefined && xsdFalse(gradient) ? 'solid' : 'gradient',
    ...(border === undefined ? {} : { border: xsdTrue(border) }),
    ...(showValue === undefined ? {} : { showValue: !xsdFalse(showValue) }),
    direction: direction === 'rightToLeft' ? 'rightToLeft' : 'leftToRight',
  };
}

function iconSetPayload(rule: XmlFragment): CfRulePayload {
  const container = child(rule, 'iconSet');
  if (container === null) return unsupported('unknownRuleType', 'iconSet');
  const iconSet = attributeValueOf(container.attributes, 'iconSet') ?? '';
  if (!CF_ICON_SETS.some((candidate) => candidate.id === iconSet)) {
    return unsupported('unknownRuleType', 'iconSet');
  }
  const parsed = thresholdsOf(container);
  if (parsed.kind === 'unsupported') return parsed;
  const showValue = attributeValueOf(container.attributes, 'showValue');
  const iconOnly = showValue !== undefined && xsdFalse(showValue);
  return {
    kind: 'iconSet',
    iconSet: iconSet as CfIconSetId,
    thresholds: parsed.thresholds,
    ...(xsdTrue(attributeValueOf(container.attributes, 'reverse')) ? { reverse: true } : {}),
    ...(iconOnly ? { iconOnly: true, showValue: false } : {}),
  };
}

function x14IconSetPayload(rule: XmlFragment): CfRulePayload {
  const container = child(rule, 'iconSet');
  if (container === null) return unsupported('unknownRuleType', 'iconSet');
  const iconSet = attributeValueOf(container.attributes, 'iconSet') ?? '';
  if (!CF_ICON_SETS.some((candidate) => candidate.id === iconSet)) {
    return unsupported('unknownRuleType', 'iconSet');
  }
  if (xsdTrue(attributeValueOf(container.attributes, 'custom'))) {
    return unsupported('unknownRuleType', 'iconSet');
  }
  const parsed = thresholdsOf(container);
  if (parsed.kind === 'unsupported') return parsed;
  const showValue = attributeValueOf(container.attributes, 'showValue');
  const iconOnly = showValue !== undefined && xsdFalse(showValue);
  return {
    kind: 'iconSet',
    iconSet: iconSet as CfIconSetId,
    thresholds: parsed.thresholds,
    ...(xsdTrue(attributeValueOf(container.attributes, 'reverse')) ? { reverse: true } : {}),
    ...(iconOnly ? { iconOnly: true, showValue: false } : {}),
  };
}

// ---------------------------------------------------------------------------
// sqref
// ---------------------------------------------------------------------------

/** The ranges named by an `sqref`: space-separated areas, `$` and sheet names allowed. */
export function parseSqref(sqref: string): GridRange[] {
  const ranges: GridRange[] = [];
  for (const area of sqref.split(/\s+/)) {
    if (area === '') continue;
    const range = parseSqrefArea(area);
    if (range !== null) ranges.push(range);
  }
  return ranges;
}

function parseSqrefArea(area: string): GridRange | null {
  const separator = area.lastIndexOf('!');
  const reference = separator < 0 ? area : area.slice(separator + 1);
  const normalised = reference.replaceAll('$', '');
  const match = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/i.exec(normalised);
  if (match === null) return null;
  const startCol = columnIndex(match[1]!);
  const startRow = Number(match[2]) - 1;
  const endCol = match[3] === undefined ? startCol : columnIndex(match[3]);
  const endRow = match[4] === undefined ? startRow : Number(match[4]) - 1;
  if (
    ![startRow, startCol, endRow, endCol].every((value) => Number.isInteger(value) && value >= 0)
  ) {
    return null;
  }
  return {
    startRow: Math.min(startRow, endRow),
    startCol: Math.min(startCol, endCol),
    endRow: Math.max(startRow, endRow),
    endCol: Math.max(startCol, endCol),
  };
}

function columnIndex(letters: string): number {
  let value = 0;
  for (const character of letters.toUpperCase()) {
    value = value * 26 + (character.charCodeAt(0) - 64);
  }
  return value - 1;
}

function boundingRange(ranges: readonly GridRange[]): GridRange {
  const first = ranges[0];
  if (first === undefined) return singleCellRange(0, 0);
  let range = first;
  for (const candidate of ranges.slice(1)) {
    range = {
      startRow: Math.min(range.startRow, candidate.startRow),
      startCol: Math.min(range.startCol, candidate.startCol),
      endRow: Math.max(range.endRow, candidate.endRow),
      endCol: Math.max(range.endCol, candidate.endCol),
    };
  }
  return range;
}

function sqrefOf(ranges: readonly GridRange[]): string {
  return ranges
    .map((range) => {
      const start = `${columnToLetters(range.startCol)}${range.startRow + 1}`;
      const end = `${columnToLetters(range.endCol)}${range.endRow + 1}`;
      return `${start}:${end}`;
    })
    .join(' ');
}

function rangeText(range: GridRange): string {
  return `${columnToLetters(range.startCol)}${range.startRow + 1}:${columnToLetters(
    range.endCol,
  )}${range.endRow + 1}`;
}

// ---------------------------------------------------------------------------
// State keys (§7)
// ---------------------------------------------------------------------------

/**
 * The rule's own document state: the payload plus the attributes the model
 * understands. `raw` records the key at parse time; equality means the source
 * XML still describes the rule, so it can be re-emitted verbatim.
 */
export function conditionalFormatRuleStateKey(rule: ConditionalFormatRule): string {
  const attributes = [
    rule.priority === undefined ? '' : String(rule.priority),
    rule.stopIfTrue ? '1' : '0',
    rule.dxfId === undefined ? '' : String(rule.dxfId),
  ].join(',');
  return `${cfRulePayloadKey(rule)}|${attributes}`;
}

function conditionalFormatElementStateKey(
  sqref: string,
  ranges: readonly GridRange[],
  rules: readonly ConditionalFormatRule[],
): string {
  return `${sqref}|${ranges.map(rangeText).join(' ')}|${rules
    .map(conditionalFormatRuleStateKey)
    .join(';')}`;
}

/** The modelled kind fields, independent of attributes and ranges. */
function cfRulePayloadKey(rule: ConditionalFormatRule): string {
  switch (rule.kind) {
    case 'cellIs':
      return [
        'cellIs',
        rule.operator,
        operandKey(rule.value1),
        rule.value2 === undefined ? '' : operandKey(rule.value2),
      ].join(':');
    case 'text':
      return `text:${rule.operator}:${rule.text}`;
    case 'timePeriod':
      return rule.mode === 'period'
        ? `timePeriod:period:${rule.period}`
        : `timePeriod:${rule.operator}:${rule.value1}:${rule.value2 ?? ''}`;
    case 'top10':
      return `top10:${rule.rank ?? 10}:${rule.top ? '1' : '0'}:${rule.percent ? '1' : '0'}`;
    case 'duplicate':
      return 'duplicate';
    case 'unique':
      return 'unique';
    case 'colorScale':
      return `colorScale:${rule.thresholds.map(thresholdKey).join(',')}:${rule.colors.join(',')}`;
    case 'dataBar':
      return [
        'dataBar',
        thresholdKey(rule.min),
        thresholdKey(rule.max),
        rule.color,
        rule.negativeColor ?? '',
        rule.fill ?? 'gradient',
        rule.showValue === false ? '0' : '1',
        rule.border === true ? '1' : '0',
        rule.direction ?? 'leftToRight',
        rule.axisColor ?? '',
      ].join(':');
    case 'iconSet':
      return [
        'iconSet',
        rule.iconSet,
        rule.thresholds.map(thresholdKey).join(','),
        rule.reverse === true ? '1' : '0',
        rule.iconOnly === true ? 'iconOnly' : rule.showValue === false ? 'hidden' : 'value',
      ].join(':');
    case 'unsupported':
      return `unsupported:${rule.reason}:${rule.ruleType ?? ''}`;
  }
}

function operandKey(operand: CfOperand): string {
  return `${operand.type}:${String(operand.value)}`;
}

function thresholdKey(threshold: CfThreshold): string {
  return threshold.type === 'automatic' ? 'auto' : `${threshold.type}:${threshold.value}`;
}

// ---------------------------------------------------------------------------
// Serialisation (§7)
// ---------------------------------------------------------------------------

/**
 * Write one base element back. An untouched element is emitted from its source
 * text; a changed one rewrites only the `sqref` and each changed rule, keeping
 * every unmodelled attribute and child in its authored order. A rebuilt element
 * drops the source's inter-element whitespace (its children are re-emitted).
 */
export function serialiseConditionalFormatElement(
  element: ConditionalFormatElementSource,
  rules: readonly ConditionalFormatRule[] = [],
): string {
  if (rules.length === 0) return element.xml;
  const key = conditionalFormatElementStateKey(element.sqref, element.ranges, rules);
  if (key === element.stateKey) return element.xml;
  const parsed = parseFragment(element.xml, 'conditionalFormatting');
  if (parsed === null) return element.xml;
  const sqref = sqrefOf(element.ranges);
  const attributes = updateAttributes(parsed.attributes, new Map([['sqref', sqref]]));
  const openTag = `<${qualified(parsed.prefix, 'conditionalFormatting')}${renderAttributes(
    attributes,
  )}>`;
  let ruleIndex = 0;
  const children = parsed.elements
    .map((child) => {
      if (child.local !== 'cfRule') return element.xml.slice(child.start, child.end);
      const rule = rules[ruleIndex];
      ruleIndex += 1;
      return rule === undefined
        ? element.xml.slice(child.start, child.end)
        : serialiseConditionalFormatRule(rule);
    })
    .join('');
  return `${openTag}${children}</${qualified(parsed.prefix, 'conditionalFormatting')}>`;
}

/**
 * Write one `cfRule` back. An unchanged rule emits its source text. A rule whose
 * payload changed cannot be written in this milestone, because authoring is v2,
 * so the save is refused rather than writing a rule the file never described.
 * Attribute-only changes (priority, dxfId, stop-if-true and the kind's own
 * attributes) rewrite just those attributes, so unknown attributes and every
 * child element keep their authored bytes and order.
 */
export function serialiseConditionalFormatRule(rule: ConditionalFormatRule): string {
  const raw = isConditionalFormatRuleSource(rule.raw) ? rule.raw : null;
  const source = raw?.base ?? raw?.x14 ?? null;
  if (source === null || raw === null) {
    throw new WorkbookSaveError(
      'unsupported-edit',
      'A conditional-format rule without its source XML cannot be written. The file was not written.',
    );
  }
  if (raw.stateKey === conditionalFormatRuleStateKey(rule)) return source.xml;
  if (raw.base !== null && raw.x14 !== null) {
    // The two forms of one rule describe the same rule; changing only the base
    // copy would leave the file inconsistent, and authoring is v2.
    throw new WorkbookSaveError(
      'unsupported-edit',
      'The base and extension forms of a conditional-format rule must change together; authoring arrives in v2. The file was not written.',
    );
  }
  if (raw.payloadKey !== cfRulePayloadKey(rule)) {
    throw new WorkbookSaveError(
      'unsupported-edit',
      'Conditional-format rule authoring arrives in v2; the file was not written.',
    );
  }
  const parsed = parseFragment(source.xml, `${source.source} cfRule`);
  if (parsed === null) {
    throw new WorkbookSaveError(
      'unsafe-part',
      'A conditional-format rule could not be read back for saving. The file was not written.',
    );
  }
  const attributes = updateAttributes(parsed.attributes, ruleAttributeUpdates(rule));
  const openTag = `<${qualified(parsed.prefix, 'cfRule')}${renderAttributes(attributes)}${
    parsed.selfClosing ? '/>' : '>'
  }`;
  if (parsed.selfClosing) return openTag;
  return `${openTag}${source.xml.slice(parsed.openEnd)}`;
}

/**
 * The attributes the model can change. `null` removes an attribute, so a flag
 * that became false is dropped instead of being written as `0`.
 */
function ruleAttributeUpdates(rule: ConditionalFormatRule): Map<string, string | null> {
  const updates = new Map<string, string | null>();
  updates.set('priority', rule.priority === undefined ? null : String(rule.priority));
  updates.set('stopIfTrue', rule.stopIfTrue ? '1' : null);
  updates.set('dxfId', rule.dxfId === undefined ? null : String(rule.dxfId));
  if (rule.kind === 'cellIs') updates.set('operator', rule.operator);
  if (rule.kind === 'text') {
    updates.set('operator', textOperatorAttribute(rule.operator));
    updates.set('text', rule.text);
  }
  if (rule.kind === 'timePeriod' && rule.mode === 'period') {
    updates.set('timePeriod', rule.period);
  }
  if (rule.kind === 'top10') {
    updates.set('rank', String(rule.rank ?? 10));
    updates.set('top', rule.top ? '1' : '0');
    updates.set('percent', rule.percent ? '1' : '0');
  }
  return updates;
}

function textOperatorAttribute(operator: CfTextOperator): string {
  switch (operator) {
    case 'contains':
      return 'containsText';
    case 'notContains':
      return 'notContainsText';
    case 'beginsWith':
      return 'beginsWith';
    case 'endsWith':
      return 'endsWith';
  }
}

/** Update modelled attributes in place, keep the authored order, drop `null`s. */
function updateAttributes(
  attributes: readonly XmlAttributeRecord[],
  updates: ReadonlyMap<string, string | null>,
): XmlAttributeRecord[] {
  const applied = new Set<string>();
  const result: XmlAttributeRecord[] = [];
  for (const attribute of attributes) {
    const local = localName(attribute.name);
    if (!updates.has(local)) {
      result.push(attribute);
      continue;
    }
    applied.add(local);
    const value = updates.get(local);
    if (value !== null && value !== undefined) {
      result.push({ name: attribute.name, value });
    }
  }
  for (const [name, value] of updates) {
    if (value === null || value === undefined || applied.has(name)) continue;
    result.push({ name, value });
  }
  return result;
}

function renderAttributes(attributes: readonly XmlAttributeRecord[]): string {
  return attributes
    .map((attribute) => ` ${attribute.name}="${escapeXmlAttribute(attribute.value)}"`)
    .join('');
}

// ---------------------------------------------------------------------------
// Save planning (§7)
// ---------------------------------------------------------------------------

export interface ConditionalFormatTextEdit {
  readonly start: number;
  readonly end: number;
  readonly text: string;
}

export interface ConditionalFormatSavePlan {
  /** `true` when the worksheet part must be written for the conditional formats. */
  readonly changed: boolean;
  /** The desired base elements, in document order: the identity check's target. */
  readonly elements: readonly string[];
  /** Position edits against the source part, for the position patch writer. */
  readonly edits: readonly ConditionalFormatTextEdit[];
}

const NO_PLAN: ConditionalFormatSavePlan = Object.freeze({
  changed: false,
  elements: Object.freeze([]) as readonly string[],
  edits: Object.freeze([]) as readonly ConditionalFormatTextEdit[],
});

/**
 * Decide what the save must do with one sheet's conditional formatting. The
 * records hold each base element verbatim, so comparing them with the source
 * part answers "did anything change?" without a model diff: an untouched sheet
 * produces no edits and keeps the part's bytes (§7).
 */
export function planConditionalFormatsSave(
  sheet: Sheet,
  sourceBytes: Uint8Array | null,
): ConditionalFormatSavePlan {
  const desired = sheet.conditionalFormats.map((record) => record.rawXml);
  if (sourceBytes === null) {
    if (desired.length === 0) return NO_PLAN;
    throw new WorkbookSaveError(
      'unsupported-edit',
      `The worksheet ${sheet.partPath} has no source skeleton to receive its conditional formats. The file was not written.`,
    );
  }
  if (desired.length === 0 && !sourceHasConditionalFormatting(sourceBytes)) return NO_PLAN;
  const xml = decodeXml(sourceBytes);

  let scan: WorksheetScan;
  try {
    scan = scanWorksheet(xml, sheet.partPath);
  } catch {
    throw new WorkbookSaveError(
      'unsafe-part',
      `The worksheet ${sheet.partPath} could not be read for conditional formatting. The file was not written.`,
    );
  }
  const source = scan.base.map((span) => xml.slice(span.start, span.end));
  const edits: ConditionalFormatTextEdit[] = [];
  const shared = Math.min(source.length, desired.length);
  for (let index = 0; index < shared; index += 1) {
    if (source[index] === desired[index]) continue;
    const span = scan.base[index]!;
    edits.push({ start: span.start, end: span.end, text: desired[index]! });
  }
  for (let index = shared; index < desired.length; index += 1) {
    const position = insertionPosition(scan);
    edits.push({ start: position, end: position, text: desired[index]! });
  }
  for (let index = shared; index < source.length; index += 1) {
    const span = scan.base[index]!;
    edits.push({ start: span.start, end: span.end, text: '' });
  }
  return {
    changed: edits.length > 0,
    elements: desired,
    edits,
  };
}

/** A byte scan: does this worksheet part carry a `conditionalFormatting` element? */
function sourceHasConditionalFormatting(bytes: Uint8Array): boolean {
  const needle = 'conditionalFormatting';
  const first = needle.charCodeAt(0);
  const last = needle.charCodeAt(needle.length - 1);
  for (let index = 0; index + needle.length <= bytes.length; index += 1) {
    if (bytes[index] !== first || bytes[index + needle.length - 1] !== last) continue;
    let offset = 1;
    while (offset < needle.length && bytes[index + offset] === needle.charCodeAt(offset)) {
      offset += 1;
    }
    if (offset === needle.length) return true;
  }
  return false;
}

/** Children of `<worksheet>` that follow `conditionalFormatting` in CT_Worksheet. */
const AFTER_CONDITIONAL_FORMATTING: ReadonlySet<string> = new Set([
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

function insertionPosition(scan: WorksheetScan): number {
  const last = scan.base.at(-1);
  if (last !== undefined) return last.end;
  const next = scan.directChildren.find((child) => AFTER_CONDITIONAL_FORMATTING.has(child.local));
  return next?.start ?? scan.rootCloseStart;
}

// ---------------------------------------------------------------------------
// Differential formats (`dxfs`, §2 and §7)
// ---------------------------------------------------------------------------

export interface DxfTableEntry {
  /** The append-only id: the entry's index in the table. */
  readonly id: number;
  readonly rawXml: string;
}

export interface DxfTable {
  readonly partPath: string;
  readonly entries: readonly DxfTableEntry[];
  /** The authored `count` attribute, when one is present. */
  readonly declaredCount: number | null;
}

/**
 * Read the styles part's `dxfs` table. Ids are positions and are never
 * renumbered (§7); this milestone appends nothing, so the table is read-only.
 */
export function parseDxfTable(bytes: Uint8Array, partPath: string): DxfTable {
  const xml = decodeXml(bytes);
  const entries: DxfTableEntry[] = [];
  const frames: { dxf: { start: number; end: number } | null }[] = [];
  let declaredCount: number | null = null;
  parseXml(
    xml,
    {
      onOpenTag(tag, position) {
        const start = openTagStart(xml, position);
        if (frames.length === 1 && tag.local === 'dxfs') {
          const count = attributeValue(tag, 'count');
          const parsed = count === undefined ? Number.NaN : Number(count);
          if (Number.isFinite(parsed)) declaredCount = parsed;
        }
        const dxf = frames.length === 2 && tag.local === 'dxf' ? { start, end: position } : null;
        if (dxf !== null) entries.push({ id: entries.length, rawXml: '' });
        frames.push({ dxf });
      },
      onCloseTag(_tag, _start, end) {
        const frame = frames.pop();
        if (frame === undefined || frame.dxf === null) return;
        const index = entries.length - 1;
        const entry = entries[index];
        if (entry !== undefined) {
          entries[index] = { id: entry.id, rawXml: xml.slice(frame.dxf.start, end) };
        }
      },
    },
    partPath,
  );
  return { partPath, entries, declaredCount };
}

/** Resolve a rule's `dxfId` to the differential format it indexes. */
export function createDxfResolver(table: DxfTable | null): CfDxfResolver {
  if (table === null) return () => null;
  const cache = new Map<number, CfDifferentialFormat | null>();
  return (dxfId: number) => {
    const cached = cache.get(dxfId);
    if (cached !== undefined) return cached;
    const entry = table.entries[dxfId];
    const resolved =
      entry === undefined ? null : parseDifferentialFormat(entry.rawXml, table.partPath);
    cache.set(dxfId, resolved);
    return resolved;
  };
}

/**
 * Read one `<dxf>` into the four properties conditional formatting merges.
 * Properties the record does not set stay `undefined`, so manual formatting
 * shows through where the rule is silent (§3.1).
 */
export function parseDifferentialFormat(rawXml: string, partPath: string): CfDifferentialFormat {
  const dxf = parseFragment(rawXml, `${partPath} (dxf)`);
  if (dxf === null) return { rawXml };
  const format: {
    fill?: CfDifferentialFormat['fill'];
    font?: CfDifferentialFormat['font'];
    border?: CfDifferentialFormat['border'];
    numberFormat?: CfDifferentialFormat['numberFormat'];
  } = {};
  for (const node of dxf.elements) {
    switch (node.local) {
      case 'font':
        format.font = dxfFont(node);
        break;
      case 'fill':
        format.fill = dxfFill(node);
        break;
      case 'border':
        format.border = dxfBorder(node);
        break;
      case 'numFmt': {
        const numFmtId = integerAttribute(node.attributes, 'numFmtId');
        const code = attributeValueOf(node.attributes, 'formatCode');
        format.numberFormat = {
          ...(numFmtId === undefined ? {} : { numFmtId }),
          ...(code === undefined ? {} : { code: decodeAttr(code) }),
        };
        break;
      }
      default:
        break;
    }
  }
  return { ...format, rawXml };
}

function dxfFont(node: XmlFragment): NonNullable<CfDifferentialFormat['font']> {
  const font: {
    name?: string;
    size?: number;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    underlineStyle?: ReturnType<typeof underlineFromValue>['underlineStyle'];
    strike?: boolean;
    color?: string | null;
  } = {};
  for (const child of node.elements) {
    switch (child.local) {
      case 'b':
        font.bold = !xsdFalse(attributeValueOf(child.attributes, 'val'));
        break;
      case 'i':
        font.italic = !xsdFalse(attributeValueOf(child.attributes, 'val'));
        break;
      case 'strike':
        font.strike = !xsdFalse(attributeValueOf(child.attributes, 'val'));
        break;
      case 'u': {
        const underline = underlineFromValue(attributeValueOf(child.attributes, 'val'));
        font.underline = underline.underline;
        font.underlineStyle = underline.underlineStyle;
        break;
      }
      case 'sz': {
        const value = Number(attributeValueOf(child.attributes, 'val') ?? '');
        if (Number.isFinite(value) && value > 0) font.size = value;
        break;
      }
      case 'name': {
        const value = attributeValueOf(child.attributes, 'val');
        if (value !== undefined && value !== '') font.name = decodeAttr(value);
        break;
      }
      case 'color':
        font.color = colorHex(child, {});
        break;
      default:
        break;
    }
  }
  return font;
}

function dxfFill(node: XmlFragment): NonNullable<CfDifferentialFormat['fill']> {
  const fill: { pattern?: string; color?: string | null; backgroundColor?: string | null } = {};
  for (const child of node.elements) {
    if (child.local !== 'patternFill') continue;
    const pattern = attributeValueOf(child.attributes, 'patternType');
    if (pattern !== undefined) fill.pattern = pattern;
    for (const color of child.elements) {
      if (color.local === 'fgColor') fill.color = colorHex(color, {});
      if (color.local === 'bgColor') fill.backgroundColor = colorHex(color, {});
    }
  }
  return fill;
}

function dxfBorder(node: XmlFragment): NonNullable<CfDifferentialFormat['border']> {
  const border: {
    left?: { style: string; color: string | null };
    right?: { style: string; color: string | null };
    top?: { style: string; color: string | null };
    bottom?: { style: string; color: string | null };
    diagonal?: { style: string; color: string | null };
    diagonalUp?: boolean;
    diagonalDown?: boolean;
  } = {};
  if (attributeValueOf(node.attributes, 'diagonalUp') !== undefined) {
    border.diagonalUp = !xsdFalse(attributeValueOf(node.attributes, 'diagonalUp'));
  }
  if (attributeValueOf(node.attributes, 'diagonalDown') !== undefined) {
    border.diagonalDown = !xsdFalse(attributeValueOf(node.attributes, 'diagonalDown'));
  }
  for (const child of node.elements) {
    if (!['left', 'right', 'top', 'bottom', 'diagonal'].includes(child.local)) continue;
    const colorNode = child.elements.find((candidate) => candidate.local === 'color') ?? null;
    const edge = {
      style: attributeValueOf(child.attributes, 'style') ?? 'none',
      color: colorNode === null ? null : colorHex(colorNode, {}),
    };
    switch (child.local) {
      case 'left':
        border.left = edge;
        break;
      case 'right':
        border.right = edge;
        break;
      case 'top':
        border.top = edge;
        break;
      case 'bottom':
        border.bottom = edge;
        break;
      case 'diagonal':
        border.diagonal = edge;
        break;
      default:
        break;
    }
  }
  return border;
}

// ---------------------------------------------------------------------------
// Unsupported-rule reporting (§8)
// ---------------------------------------------------------------------------

export interface UnsupportedConditionalFormat {
  readonly sheetId: SheetId;
  readonly sheetName: string;
  /** The sqref areas the rule covers; each is its own evaluation domain. */
  readonly ranges: readonly GridRange[];
  /** The authored rule type, for example `expression` or `aboveAverage`. */
  readonly ruleType: string;
  readonly reason: CfUnsupportedReason;
  /** Which XML form carries the rule; x14 wins when both describe it. */
  readonly source: 'base' | 'x14';
  readonly priority?: number;
  readonly dxfId?: number;
}

/**
 * The unsupported rules of one sheet's rule list (§8). Preserved and never
 * painted; the compatibility and export reports consume this listing.
 */
export function unsupportedConditionalFormatsInRules(
  sheet: Sheet,
  rules: readonly ConditionalFormatRule[],
): UnsupportedConditionalFormat[] {
  const listing: UnsupportedConditionalFormat[] = [];
  for (const rule of rules) {
    if (rule.kind !== 'unsupported') continue;
    const raw = isConditionalFormatRuleSource(rule.raw) ? rule.raw : null;
    listing.push({
      sheetId: sheet.id,
      sheetName: sheet.name,
      ranges: rule.ranges,
      ruleType: ruleTypeOf(rule),
      reason: rule.reason,
      source: raw !== null && raw.x14 !== null ? 'x14' : 'base',
      ...(rule.priority === undefined ? {} : { priority: rule.priority }),
      ...(rule.dxfId === undefined ? {} : { dxfId: rule.dxfId }),
    });
  }
  return listing;
}

/** Every unsupported rule in the workbook, sheet by sheet, in sheet order (§8). */
export function listUnsupportedConditionalFormats(
  workbook: Workbook,
  sourceBytesFor: (sheet: Sheet) => Uint8Array | null = (sheet) =>
    workbook.sourceParts.get(sheet.partPath) ?? null,
): UnsupportedConditionalFormat[] {
  const listing: UnsupportedConditionalFormat[] = [];
  for (const sheet of workbook.sheets) {
    const rules = conditionalFormatRulesFor(sheet, sourceBytesFor(sheet));
    listing.push(...unsupportedConditionalFormatsInRules(sheet, rules));
  }
  return listing;
}

// ---------------------------------------------------------------------------
// XML helpers
// ---------------------------------------------------------------------------

function parseFragment(xml: string, partName: string): XmlFragment | null {
  // A prefixed element cannot be parsed on its own: its namespace binding sits
  // on an ancestor. Wrap the fragment with declarations for every prefix it
  // uses, then report positions back in the fragment's own indexes. An XML
  // declaration, when one is present, stays outside the wrapper.
  const declarationEnd = xml.startsWith('<?xml') ? xml.indexOf('?>') + 2 : 0;
  const body = xml.slice(declarationEnd);
  const declarations = fragmentNamespaceDeclarations(body);
  const wrapper = declarations === '' ? '' : `<calcplus-fragment${declarations}>`;
  const source =
    wrapper === '' ? xml : `${xml.slice(0, declarationEnd)}${wrapper}${body}</calcplus-fragment>`;
  const offset = wrapper.length;
  interface Pending {
    readonly local: string;
    readonly prefix: string;
    readonly attributes: readonly XmlAttributeRecord[];
    elements: XmlFragment[];
    text: string;
    readonly start: number;
    readonly openEnd: number;
    readonly selfClosing: boolean;
  }
  const stack: Pending[] = [];
  let root: XmlFragment | null = null;
  try {
    parseXml(
      source,
      {
        onOpenTag(tag, position) {
          const tagText = source.slice(openTagStart(source, position), position);
          stack.push({
            local: tag.local,
            prefix: tag.prefix,
            attributes: fragmentAttributes(tag),
            elements: [],
            text: '',
            start: openTagStart(source, position) - offset,
            openEnd: position - offset,
            selfClosing: /\/\s*>$/.test(tagText),
          });
        },
        onCloseTag(_tag, _start, end) {
          const open = stack.pop();
          if (open === undefined) return;
          const node: XmlFragment = {
            local: open.local,
            prefix: open.prefix,
            attributes: open.attributes,
            elements: open.elements,
            text: open.text,
            start: open.start,
            end: end - offset,
            openEnd: open.openEnd,
            selfClosing: open.selfClosing,
          };
          const parent = stack.at(-1);
          if (parent === undefined) root = node;
          else parent.elements.push(node);
        },
        onText(text) {
          const top = stack.at(-1);
          if (top !== undefined) top.text += text;
        },
        onCdata(text) {
          const top = stack.at(-1);
          if (top !== undefined) top.text += text;
        },
      },
      partName,
    );
  } catch {
    return null;
  }
  if (root === null) return null;
  return wrapper === '' ? root : ((root as XmlFragment).elements[0] ?? null);
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

interface SaxesAttributesTag {
  readonly attributes: Record<string, { readonly name: string; readonly value: string }>;
}

function fragmentAttributes(tag: SaxesAttributesTag): XmlAttributeRecord[] {
  const attributes: XmlAttributeRecord[] = [];
  for (const attribute of Object.values(tag.attributes)) {
    // Namespace declarations belong to an ancestor; a rebuilt element must not
    // repeat them.
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

function integerAttribute(
  attributes: readonly XmlAttributeRecord[],
  local: string,
): number | undefined {
  const text = attributeValueOf(attributes, local);
  if (text === undefined) return undefined;
  const value = Number(text);
  return Number.isInteger(value) ? value : undefined;
}

function localName(name: string): string {
  const separator = name.lastIndexOf(':');
  return separator < 0 ? name : name.slice(separator + 1);
}

function qualified(prefix: string, local: string): string {
  return prefix === '' ? local : `${prefix}:${local}`;
}

/** OOXML xsd:boolean accepts `0`, `1`, `true` and `false`. */
function xsdTrue(value: string | undefined): boolean {
  return value === '1' || value === 'true';
}

/** `false` unless the attribute is explicitly false: Excel omits the default. */
function xsdFalse(value: string | undefined): boolean {
  return value === '0' || value === 'false';
}

function decodeAttr(value: string): string {
  return value
    .replaceAll('&quot;', '"')
    .replaceAll('&apos;', "'")
    .replaceAll('&lt;', '<')
    .replaceAll('&gt;', '>')
    .replaceAll('&amp;', '&');
}
