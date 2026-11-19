import type { GridRange, Sheet, ValidationRecord } from '@model/types';
import { singleCellRange } from '@model/types';
import { parseSqref } from './conditional-format-xml';

/**
 * The `dataValidations` element of a worksheet part (file-compat §3, §9 F13).
 *
 * Reading maps every `<dataValidation>` element onto `Sheet.validation` as one
 * record per element: the element's `sqref` areas are the record's reach (a
 * union is represented by its bounding rectangle, exactly as the
 * conditional-format records do), and the element XML is kept verbatim in
 * `rawXml`. That is the `autofilter-xml.ts` pattern — modelled range plus raw
 * source — with one deliberate difference: v1 never authors validation
 * (feature matrix "Data validation: IN (preserve)"), so there is no write
 * planner and no state key. A save leaves the source element alone, and a
 * structural edit rewrites the `sqref` through the model's range transform
 * (`src/model/range-transform.ts`) and the source-part renumber (`renumberWorksheetPart`).
 *
 * The module never enforces a rule and never dispatches a command.
 */

export interface ValidationElementSource {
  readonly kind: 'dataValidation-source';
  /** The exact `dataValidation` element, preserved verbatim. */
  readonly xml: string;
  /** The `sqref` text as authored. */
  readonly sqref: string;
  /** Every area the `sqref` names; empty when it cannot be parsed. */
  readonly ranges: readonly GridRange[];
}

export interface ParsedValidationElement {
  readonly record: ValidationRecord;
  readonly source: ValidationElementSource;
}

/** Parse one `dataValidation` element into a record, or `null` when it has no ref. */
export function parseValidationElement(elementXml: string): ParsedValidationElement | null {
  const openTag = openTagOf(elementXml);
  if (openTag === null) return null;
  const attributes = parseAttributes(openTag.text);
  const sqref = attributes.get('sqref');
  if (sqref === undefined || sqref.trim() === '') return null;
  const ranges = parseSqref(sqref);
  const range = boundingRange(ranges);
  return {
    record: { range, rawXml: elementXml },
    source: { kind: 'dataValidation-source', xml: elementXml, sqref, ranges },
  };
}

/**
 * Parse the sheet's `<dataValidations>` children. `rootChildren` are the direct
 * children of `<worksheet>` with their source positions; the element text is
 * sliced from the part so unknown attributes and children stay verbatim.
 */
export function parseValidationElements(
  xml: string,
  rootChildren: readonly { readonly local: string; readonly start: number; readonly end: number }[],
): ParsedValidationElement[] {
  const elements: ParsedValidationElement[] = [];
  for (const child of rootChildren) {
    if (child.local !== 'dataValidations') continue;
    const slice = xml.slice(child.start, child.end);
    for (const element of scanElements(slice, 'dataValidation')) {
      const parsed = parseValidationElement(element.text);
      if (parsed !== null) elements.push(parsed);
    }
  }
  return elements;
}

/** Populate `Sheet.validation` from the worksheet part. */
export function applyValidations(
  sheet: Sheet,
  xml: string,
  rootChildren: readonly { readonly local: string; readonly start: number; readonly end: number }[],
): readonly ValidationRecord[] {
  const records = parseValidationElements(xml, rootChildren).map((entry) => entry.record);
  sheet.validation.splice(0, sheet.validation.length, ...records);
  return records;
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

// ---------------------------------------------------------------------------
// Text helpers
// ---------------------------------------------------------------------------

interface OpenTag {
  readonly text: string;
  readonly start: number;
  readonly end: number;
}

interface ScannedElement {
  readonly text: string;
}

function openTagOf(xml: string): OpenTag | null {
  const start = xml.indexOf('<');
  if (start < 0) return null;
  const end = findTagEnd(xml, start);
  if (end < 0) return null;
  return { text: xml.slice(start, end), start, end };
}

/** Find every element with one local name in a fragment (they never nest in themselves). */
function scanElements(xml: string, local: string): ScannedElement[] {
  const elements: ScannedElement[] = [];
  const pattern = new RegExp(`<(?:[\\w.-]+:)?${local}\\b`, 'g');
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(xml)) !== null) {
    const start = match.index;
    const openEnd = findTagEnd(xml, start);
    if (openEnd < 0) break;
    const openTag = xml.slice(start, openEnd);
    if (/\/\s*>$/.test(openTag)) {
      elements.push({ text: openTag });
      pattern.lastIndex = openEnd;
      continue;
    }
    const qualified = /^<([^\s/>]+)/.exec(openTag)?.[1] ?? local;
    const close = `</${qualified}>`;
    const closeIndex = xml.indexOf(close, openEnd);
    if (closeIndex < 0) break;
    const end = closeIndex + close.length;
    elements.push({ text: xml.slice(start, end) });
    pattern.lastIndex = end;
  }
  return elements;
}

function findTagEnd(xml: string, start: number): number {
  let quoted: string | null = null;
  for (let index = start + 1; index < xml.length; index += 1) {
    const character = xml[index] ?? '';
    if (quoted !== null) {
      if (character === quoted) quoted = null;
      continue;
    }
    if (character === '"' || character === "'") {
      quoted = character;
      continue;
    }
    if (character === '>') return index + 1;
  }
  return -1;
}

function parseAttributes(openTag: string): Map<string, string> {
  const attributes = new Map<string, string>();
  const pattern = /([A-Za-z_][\w.:-]*)\s*=\s*("([^"]*)"|'([^']*)')/g;
  for (const match of openTag.matchAll(pattern)) {
    const name = match[1] ?? '';
    const local = name.includes(':') ? name.slice(name.lastIndexOf(':') + 1) : name;
    attributes.set(local, decodeEntities(match[3] ?? match[4] ?? ''));
  }
  return attributes;
}

function decodeEntities(text: string): string {
  return text
    .replaceAll('&quot;', '"')
    .replaceAll('&apos;', "'")
    .replaceAll('&lt;', '<')
    .replaceAll('&gt;', '>')
    .replaceAll('&amp;', '&');
}
