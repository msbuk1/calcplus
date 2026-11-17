import type { SheetId, Workbook } from '@model/types';
import type { Relationship } from '../types';
import { WorkbookSaveError } from '../errors';
import {
  RELATIONSHIPS_NS,
  decodeXml,
  encodeXml,
  escapeXmlAttribute,
  isWellFormedXml,
  openTagStart,
  parseXml,
} from '../xml';

/**
 * Structural sheet edits patch the workbook part rather than regenerating it.
 * Existing `<sheet>` elements keep every unmodelled attribute, a rename
 * replaces only `name`, a delete removes only that element, and an insert adds
 * one element at its tab position. Reorder and hide authoring are v2; the
 * source order is preserved.
 */

const WORKSHEET_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';
const WORKSHEET_RELATIONSHIP_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';

export interface WorkbookSheetElement {
  readonly name: string;
  readonly numericSheetId: number;
  readonly relationshipId: string;
  readonly relationshipIdAttributeName: string;
  readonly prefix: string;
  readonly start: number;
  readonly openEnd: number;
  readonly end: number;
  readonly selfClosing: boolean;
}

interface SheetsContainer {
  readonly start: number;
  readonly openEnd: number;
  readonly contentEnd: number;
  readonly end: number;
  readonly selfClosing: boolean;
  readonly prefix: string;
}

export interface ParsedWorkbookSheets {
  readonly xml: string;
  readonly container: SheetsContainer | null;
  readonly elements: WorkbookSheetElement[];
  readonly rootPrefix: string;
}

export function parseWorkbookSheets(bytes: Uint8Array, partName: string): ParsedWorkbookSheets {
  const xml = decodeXml(bytes);
  const elements: WorkbookSheetElement[] = [];
  const stack: string[] = [];
  let container: SheetsContainer | null = null;
  let rootPrefix = '';
  let current: {
    name: string;
    numericSheetId: number;
    relationshipId: string;
    relationshipIdAttributeName: string;
    prefix: string;
    start: number;
    openEnd: number;
  } | null = null;

  parseXml(
    xml,
    {
      onOpenTag(tag, position) {
        const parent = stack.at(-1) ?? null;
        const start = openTagStart(xml, position);
        stack.push(tag.local);
        if (tag.local === 'workbook' && parent === null) {
          rootPrefix = tag.prefix;
          return;
        }
        if (tag.local === 'sheets' && parent === 'workbook') {
          container = {
            start,
            openEnd: position,
            contentEnd: position,
            end: position,
            selfClosing: false,
            prefix: tag.prefix,
          };
          return;
        }
        if (tag.local === 'sheet' && parent === 'sheets') {
          const relationship = relationshipId(tag);
          current = {
            name: attribute(tag, 'name') ?? '',
            numericSheetId: Number(attribute(tag, 'sheetId') ?? '0'),
            relationshipId: relationship.value,
            relationshipIdAttributeName: relationship.name,
            prefix: tag.prefix,
            start,
            openEnd: position,
          };
        }
      },
      onCloseTag(tag, start, end) {
        stack.pop();
        if (tag.local === 'sheet' && current !== null) {
          elements.push({
            ...current,
            selfClosing: current.openEnd === end,
            end,
          });
          current = null;
          return;
        }
        if (tag.local === 'sheets' && container !== null) {
          const openEnd = (container as SheetsContainer).openEnd;
          container = {
            ...(container as SheetsContainer),
            contentEnd: start,
            end,
            selfClosing: openEnd === end,
          };
        }
      },
    },
    partName,
  );

  if (container === null) {
    throw new WorkbookSaveError(
      'unsafe-part',
      `The workbook part ${partName} has no sheets element. The file was not written.`,
    );
  }
  return { xml, container, elements: elements.sort((a, b) => a.start - b.start), rootPrefix };
}

export interface WorkbookSheetAddition {
  readonly sheetId: SheetId;
  readonly name: string;
  readonly partPath: string;
  readonly relationshipId: string;
  readonly numericSheetId: number;
  readonly target: string;
  /** Zero-based tab position in the model. */
  readonly modelIndex: number;
}

export interface WorkbookSheetRemoval {
  readonly sheetId: SheetId;
  readonly partPath: string | null;
  readonly relationshipId: string;
}

export interface WorkbookSheetRename {
  readonly sheetId: SheetId;
  readonly name: string;
  readonly element: WorkbookSheetElement;
}

export interface WorkbookSheetPlan {
  readonly additions: readonly WorkbookSheetAddition[];
  readonly removals: readonly WorkbookSheetRemoval[];
  readonly renames: readonly WorkbookSheetRename[];
  readonly changed: boolean;
  /** The model sheet order, so insert positions can skip other additions. */
  readonly modelSheetIds: readonly SheetId[];
  /** Existing workbook elements matched to their stable sheet id. */
  readonly elementBySheetId: ReadonlyMap<SheetId, WorkbookSheetElement>;
}

export interface WorkbookSheetPlanInput {
  readonly workbook: Workbook;
  readonly parsed: ParsedWorkbookSheets;
  /** Every relationship in the workbook relationships part, sheet or not. */
  readonly workbookRelationships: readonly Relationship[];
  readonly workbookPartPath: string;
  readonly workbookRelsPath: string;
  readonly sheetPartPaths: ReadonlyMap<SheetId, string>;
  readonly sheetRelationshipIds: ReadonlyMap<SheetId, string>;
}

export function planWorkbookSheetChanges(input: WorkbookSheetPlanInput): WorkbookSheetPlan {
  const relationshipIds = new Map(input.sheetRelationshipIds);
  if (relationshipIds.size === 0 && input.parsed.elements.length > 0) {
    // A snapshot written before sheet-structure edits existed has no
    // relationship map. It can only describe an unchanged sheet list, so
    // mapping by order is safe; a structural edit in that state is refused.
    if (input.parsed.elements.length !== input.workbook.sheets.length) {
      throw new WorkbookSaveError(
        'unsafe-part',
        'The workbook sheet relationships are unavailable, so a structural sheet edit cannot be written safely. The file was not written.',
      );
    }
    input.workbook.sheets.forEach((sheet, index) => {
      relationshipIds.set(sheet.id, input.parsed.elements[index]!.relationshipId);
    });
  }

  const elementByRelationshipId = new Map(
    input.parsed.elements.map((element) => [element.relationshipId, element]),
  );
  const elementIndexByRelationshipId = new Map(
    input.parsed.elements.map((element, index) => [element.relationshipId, index]),
  );
  const usedRelationshipIds = new Set<string>();
  for (const relationship of input.workbookRelationships) {
    usedRelationshipIds.add(relationship.id);
  }
  for (const relationshipId of relationshipIds.values()) {
    usedRelationshipIds.add(relationshipId);
  }
  let nextNumericSheetId =
    input.parsed.elements.reduce((highest, element) => {
      const value =
        Number.isInteger(element.numericSheetId) && element.numericSheetId > 0
          ? element.numericSheetId
          : 0;
      return Math.max(highest, value);
    }, 0) + 1;

  const additions: WorkbookSheetAddition[] = [];
  const renames: WorkbookSheetRename[] = [];
  const elementBySheetId = new Map<SheetId, WorkbookSheetElement>();
  const modelSheetIds = input.workbook.sheets.map((sheet) => sheet.id);
  const modelSheetIdSet = new Set(modelSheetIds);
  let lastElementIndex = -1;

  for (const [modelIndex, sheet] of input.workbook.sheets.entries()) {
    const relationshipId = relationshipIds.get(sheet.id);
    const element =
      relationshipId === undefined ? undefined : elementByRelationshipId.get(relationshipId);
    if (element === undefined) {
      const additionRelationshipId = nextRelationshipId(usedRelationshipIds);
      const numericSheetId = nextNumericSheetId;
      nextNumericSheetId += 1;
      additions.push({
        sheetId: sheet.id,
        name: sheet.name,
        partPath: sheet.partPath,
        relationshipId: additionRelationshipId,
        numericSheetId,
        target: relativePartTarget(input.workbookPartPath, sheet.partPath),
        modelIndex,
      });
      continue;
    }
    const elementIndex = elementIndexByRelationshipId.get(element.relationshipId) ?? -1;
    if (elementIndex <= lastElementIndex) {
      throw new WorkbookSaveError(
        'unsafe-part',
        'The workbook sheet order no longer matches the file; reordering sheets is not supported yet. The file was not written.',
      );
    }
    lastElementIndex = elementIndex;
    elementBySheetId.set(sheet.id, element);
    if (element.name !== sheet.name) {
      renames.push({ sheetId: sheet.id, name: sheet.name, element });
    }
  }

  const removals: WorkbookSheetRemoval[] = [];
  for (const [sheetId, relationshipId] of relationshipIds) {
    if (modelSheetIdSet.has(sheetId)) continue;
    removals.push({
      sheetId,
      partPath: input.sheetPartPaths.get(sheetId) ?? null,
      relationshipId,
    });
  }

  return {
    additions,
    removals,
    renames,
    changed: additions.length > 0 || removals.length > 0 || renames.length > 0,
    modelSheetIds,
    elementBySheetId,
  };
}

export function patchWorkbookSheetList(
  parsed: ParsedWorkbookSheets,
  plan: WorkbookSheetPlan,
  partName: string,
): Uint8Array {
  if (!plan.changed) return encodeXml(parsed.xml);
  const { xml } = parsed;
  const container = parsed.container;
  if (container === null) return encodeXml(xml);
  const edits: { start: number; end: number; text: string; priority: number }[] = [];

  for (const removal of plan.removals) {
    const element = parsed.elements.find(
      (candidate) => candidate.relationshipId === removal.relationshipId,
    );
    if (element !== undefined) {
      edits.push({ start: element.start, end: element.end, text: '', priority: 0 });
    }
  }

  for (const rename of plan.renames) {
    const openTag = xml.slice(rename.element.start, rename.element.openEnd);
    const updated = replaceAttributeValue(openTag, 'name', rename.name);
    if (updated !== openTag) {
      edits.push({
        start: rename.element.start,
        end: rename.element.openEnd,
        text: updated,
        priority: 0,
      });
    }
  }

  if (plan.additions.length > 0) {
    const grouped = new Map<number, WorkbookSheetAddition[]>();
    for (const addition of plan.additions) {
      const position = insertionPositionFor(addition, plan, container);
      const list = grouped.get(position) ?? [];
      list.push(addition);
      grouped.set(position, list);
    }
    for (const [position, additions] of grouped) {
      const text = additions
        .map((addition) => sheetElementXml(addition, parsed, container))
        .join('');
      edits.push({ start: position, end: position, text, priority: 1 });
    }
  }

  const output = applyEdits(xml, edits);
  if (!isWellFormedXml(output)) {
    throw new WorkbookSaveError(
      'unsafe-part',
      `The workbook part ${partName} could not store the sheet list safely. The file was not written.`,
    );
  }
  return encodeXml(output);
}

function insertionPositionFor(
  addition: WorkbookSheetAddition,
  plan: WorkbookSheetPlan,
  container: SheetsContainer,
): number {
  for (let index = addition.modelIndex + 1; index < plan.modelSheetIds.length; index += 1) {
    const nextSheetId = plan.modelSheetIds[index];
    if (nextSheetId === undefined) continue;
    const element = plan.elementBySheetId.get(nextSheetId);
    if (element !== undefined) return element.start;
  }
  return container.contentEnd;
}

export function patchWorkbookRelationshipsForSheets(
  bytes: Uint8Array,
  partName: string,
  plan: WorkbookSheetPlan,
  existingRelationships: readonly Relationship[],
): Uint8Array | null {
  if (plan.removals.length === 0 && plan.additions.length === 0) return null;
  let xml = decodeXml(bytes);
  for (const removal of plan.removals) {
    const pattern = new RegExp(
      `\\s*<(?:[\\w.-]+:)?Relationship\\b[^>]*\\bId\\s*=\\s*["']${escapeRegExp(removal.relationshipId)}["'][^>]*/>\\s*`,
      'g',
    );
    xml = xml.replace(pattern, '\n');
  }
  if (plan.additions.length > 0) {
    const closing = /<\/([\w.-]+:)?Relationships\s*>/.exec(xml);
    if (closing === null) {
      throw new WorkbookSaveError(
        'unsafe-part',
        `The relationships part ${partName} has no closing element. The file was not written.`,
      );
    }
    const type =
      existingRelationships.find((relationship) => relationship.type.endsWith('/worksheet'))
        ?.type ?? WORKSHEET_RELATIONSHIP_TYPE;
    const element = `${closing[1] ?? ''}Relationship`;
    const added = plan.additions
      .map(
        (addition) =>
          `<${element} Id="${escapeXmlAttribute(addition.relationshipId)}" Type="${escapeXmlAttribute(type)}" Target="${escapeXmlAttribute(addition.target)}"/>`,
      )
      .join('\n  ');
    xml = `${xml.slice(0, closing.index)}  ${added}\n${xml.slice(closing.index)}`;
  }
  if (!isWellFormedXml(xml)) {
    throw new WorkbookSaveError(
      'unsafe-part',
      `The relationships part ${partName} could not store the sheet relationships safely. The file was not written.`,
    );
  }
  return encodeXml(xml);
}

export function patchContentTypesForSheets(
  bytes: Uint8Array,
  partName: string,
  plan: WorkbookSheetPlan,
): Uint8Array | null {
  if (plan.removals.length === 0 && plan.additions.length === 0) return null;
  let xml = decodeXml(bytes);
  for (const removal of plan.removals) {
    if (removal.partPath === null) continue;
    const pattern = new RegExp(
      `\\s*<(?:[\\w.-]+:)?Override\\b[^>]*\\bPartName\\s*=\\s*["']/${escapeRegExp(removal.partPath)}["'][^>]*/>\\s*`,
      'g',
    );
    xml = xml.replace(pattern, '\n');
  }
  if (plan.additions.length > 0) {
    const closing = /<\/([\w.-]+:)?Types\s*>/.exec(xml);
    if (closing === null) {
      throw new WorkbookSaveError(
        'unsafe-part',
        `The content-types part ${partName} has no closing element. The file was not written.`,
      );
    }
    const element = `${closing[1] ?? ''}Override`;
    const added = plan.additions
      .map(
        (addition) =>
          `<${element} PartName="/${escapeXmlAttribute(addition.partPath)}" ContentType="${WORKSHEET_CONTENT_TYPE}"/>`,
      )
      .join('\n  ');
    xml = `${xml.slice(0, closing.index)}  ${added}\n${xml.slice(closing.index)}`;
  }
  if (!isWellFormedXml(xml)) {
    throw new WorkbookSaveError(
      'unsafe-part',
      `The content-types part ${partName} could not store the sheet entries safely. The file was not written.`,
    );
  }
  return encodeXml(xml);
}

/** Remove the content-type entries of every part a sheet delete took with it. */
export function removeContentTypeOverrides(
  bytes: Uint8Array,
  partName: string,
  partPaths: readonly string[],
): Uint8Array | null {
  if (partPaths.length === 0) return null;
  let xml = decodeXml(bytes);
  const before = xml;
  for (const path of partPaths) {
    const pattern = new RegExp(
      `\\s*<(?:[\\w.-]+:)?Override\\b[^>]*\\bPartName\\s*=\\s*["']/${escapeRegExp(path)}["'][^>]*/>\\s*`,
      'g',
    );
    xml = xml.replace(pattern, '\n');
  }
  if (xml === before) return null;
  if (!isWellFormedXml(xml)) {
    throw new WorkbookSaveError(
      'unsafe-part',
      `The content-types part ${partName} could not drop the removed sheet parts safely. The file was not written.`,
    );
  }
  return encodeXml(xml);
}

function sheetElementXml(
  addition: WorkbookSheetAddition,
  parsed: ParsedWorkbookSheets,
  container: SheetsContainer,
): string {
  const sample = parsed.elements[0];
  const prefix = sample?.prefix ?? container.prefix ?? parsed.rootPrefix;
  const relationshipAttribute = sample?.relationshipIdAttributeName ?? 'r:id';
  const sheetName = prefix === '' ? 'sheet' : `${prefix}:sheet`;
  return `<${sheetName} name="${escapeXmlAttribute(addition.name)}" sheetId="${addition.numericSheetId}" ${relationshipAttribute}="${escapeXmlAttribute(addition.relationshipId)}"/>`;
}

function nextRelationshipId(used: Set<string>): string {
  let highest = 0;
  for (const id of used) {
    const match = /^rId(\d+)$/i.exec(id);
    if (match !== null) highest = Math.max(highest, Number(match[1]));
  }
  let number = highest + 1;
  let candidate = `rId${number}`;
  while (used.has(candidate)) {
    number += 1;
    candidate = `rId${number}`;
  }
  used.add(candidate);
  return candidate;
}

function relativePartTarget(fromPartPath: string, toPartPath: string): string {
  const fromParts = fromPartPath.split('/');
  fromParts.pop();
  const toParts = toPartPath.split('/');
  let common = 0;
  while (
    common < fromParts.length &&
    common < toParts.length &&
    fromParts[common] === toParts[common]
  ) {
    common += 1;
  }
  const up = fromParts.length - common;
  return [...Array.from({ length: up }, () => '..'), ...toParts.slice(common)].join('/');
}

function replaceAttributeValue(openTag: string, localName: string, value: string): string {
  const pattern = new RegExp(`(\\s${localName}\\s*=\\s*)("[^"]*"|'[^']*')`);
  if (!pattern.test(openTag)) return openTag;
  return openTag.replace(
    pattern,
    (_match: string, prefix: string) => `${prefix}"${escapeXmlAttribute(value)}"`,
  );
}

function attribute(
  tag: { attributes: Record<string, { local: string; value: string }> },
  local: string,
): string | undefined {
  for (const attributeValue of Object.values(tag.attributes)) {
    if (attributeValue.local === local) return attributeValue.value;
  }
  return undefined;
}

function relationshipId(tag: {
  attributes: Record<string, { local: string; uri?: string; value: string; name: string }>;
}): { value: string; name: string } {
  for (const attributeValue of Object.values(tag.attributes)) {
    if (attributeValue.local === 'id' && attributeValue.uri === RELATIONSHIPS_NS) {
      return { value: attributeValue.value, name: attributeValue.name };
    }
  }
  return { value: '', name: 'r:id' };
}

function applyEdits(
  xml: string,
  edits: readonly { start: number; end: number; text: string; priority: number }[],
): string {
  const sorted = [...edits].sort((left, right) => {
    if (right.start !== left.start) return right.start - left.start;
    return left.priority - right.priority;
  });
  let output = xml;
  for (const edit of sorted) {
    output = output.slice(0, edit.start) + edit.text + output.slice(edit.end);
  }
  return output;
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
