import type { Relationship } from '../types';
import { WorkbookSaveError } from '../errors';
import { decodeXml, encodeXml, isWellFormedXml, openTagStart, parseXml } from '../xml';

const AFTER_CALC_PR = new Set([
  'oleSize',
  'customWorkbookViews',
  'pivotCaches',
  'smartTagPr',
  'smartTagTypes',
  'webPublishing',
  'fileRecoveryPr',
  'webPublishObjects',
  'extLst',
]);

/** Set `calcPr/@fullCalcOnLoad="1"` without rewriting the workbook part. */
export function setFullCalcOnLoad(bytes: Uint8Array, partName: string): Uint8Array {
  const xml = decodeXml(bytes);
  const calcPr = /<(?:[\w.-]+:)?calcPr\b/.exec(xml);
  let patched: string;
  if (calcPr !== null) {
    const openEnd = xml.indexOf('>', calcPr.index);
    if (openEnd < 0) {
      throw new WorkbookSaveError(
        'unsafe-part',
        `The workbook part ${partName} has a malformed calcPr element. The file was not written.`,
      );
    }
    const tag = xml.slice(calcPr.index, openEnd + 1);
    const updated = /(\sfullCalcOnLoad\s*=\s*)("[^"]*"|'[^']*')/.test(tag)
      ? tag.replace(/(\sfullCalcOnLoad\s*=\s*)("[^"]*"|'[^']*')/, '$1"1"')
      : tag.replace(/\/>$/, ' fullCalcOnLoad="1"/>').replace(/(?<!\/)>$/, ' fullCalcOnLoad="1">');
    patched = `${xml.slice(0, calcPr.index)}${updated}${xml.slice(openEnd + 1)}`;
  } else {
    const closing = /<\/([\w.-]+:)?workbook\s*>/.exec(xml);
    if (closing === null) {
      throw new WorkbookSaveError(
        'unsafe-part',
        `The workbook part ${partName} has no closing workbook element. The file was not written.`,
      );
    }
    const prefix = closing[1] ?? '';
    const position = firstDirectChildAfterCalcPr(xml, partName) ?? closing.index;
    patched = `${xml.slice(0, position)}<${prefix}calcPr fullCalcOnLoad="1"/>${xml.slice(
      position,
    )}`;
  }
  if (!isWellFormedXml(patched)) {
    throw new WorkbookSaveError(
      'unsafe-part',
      `The workbook part ${partName} could not be patched safely. The file was not written.`,
    );
  }
  return encodeXml(patched);
}

/**
 * Set `workbookPr/@date1904` without rewriting anything else. The value is
 * one byte pattern; the rest of the workbook part keeps its original bytes.
 */
export function setWorkbookDate1904(
  bytes: Uint8Array,
  partName: string,
  date1904: boolean,
): Uint8Array {
  const xml = decodeXml(bytes);
  const value = date1904 ? '1' : '0';
  const workbookPr = /<((?:[\w.-]+:)?workbookPr)\b([^>]*)>/.exec(xml);
  let patched: string;
  if (workbookPr !== null) {
    const tagName = workbookPr[1] ?? 'workbookPr';
    const rawAttributes = workbookPr[2] ?? '';
    const selfClosing = rawAttributes.trimEnd().endsWith('/');
    const attributes = selfClosing ? rawAttributes.replace(/\/\s*$/, '') : rawAttributes;
    const dateAttribute = /(\sdate1904\s*=\s*)("[^"]*"|'[^']*')/;
    const updated = dateAttribute.test(attributes)
      ? attributes.replace(dateAttribute, `$1"${value}"`)
      : `${attributes} date1904="${value}"`;
    const replacement = `<${tagName}${updated}${selfClosing ? '/>' : '>'}`;
    patched = `${xml.slice(0, workbookPr.index)}${replacement}${xml.slice(
      workbookPr.index + workbookPr[0].length,
    )}`;
  } else {
    const position = workbookPrInsertionPoint(xml);
    const prefix = workbookPrefix(xml);
    patched = `${xml.slice(0, position)}<${prefix}workbookPr date1904="${value}"/>${xml.slice(
      position,
    )}`;
  }
  if (!isWellFormedXml(patched)) {
    throw new WorkbookSaveError(
      'unsafe-part',
      `The workbook part ${partName} could not store its date system safely. The file was not written.`,
    );
  }
  return encodeXml(patched);
}

function workbookPrefix(xml: string): string {
  const match = /<([A-Za-z_][\w.-]*:)?workbook\b/.exec(xml);
  return match?.[1] ?? '';
}

function workbookPrInsertionPoint(xml: string): number {
  const fileVersion = /<(?:[\w.-]+:)?fileVersion\b[^>]*\/>/.exec(xml);
  if (fileVersion !== null) return fileVersion.index + fileVersion[0].length;
  const next = /<(?:[\w.-]+:)?(?:bookViews|sheets|definedNames|calcPr|oleSize)\b/.exec(xml);
  if (next !== null) return next.index;
  const closing = /<\/(?:[\w.-]+:)?workbook\s*>/.exec(xml);
  return closing?.index ?? xml.length;
}

function firstDirectChildAfterCalcPr(xml: string, partName: string): number | null {
  let depth = 0;
  let found: number | null = null;
  parseXml(
    xml,
    {
      onOpenTag(tag, position) {
        depth += 1;
        if (depth === 2 && found === null && AFTER_CALC_PR.has(tag.local)) {
          found = openTagStart(xml, position);
        }
      },
      onCloseTag() {
        depth -= 1;
      },
    },
    partName,
  );
  return found;
}

export interface RemoveCalcChainInput {
  readonly entries: ReadonlyMap<string, Uint8Array>;
  readonly relationships: ReadonlyMap<string, Relationship[]>;
  readonly contentTypesPath: string;
  readonly workbookRelsPath: string;
}

export interface RemoveCalcChainResult {
  readonly entries: Map<string, Uint8Array>;
  readonly relationships: Map<string, Relationship[]>;
  readonly removedParts: string[];
  readonly patchedParts: string[];
}

/**
 * File-compat §4/§6.1/§10.5: delete calcChain and its relationship/content
 * type entry whenever a save changes formula storage.
 */
export function removeCalcChain(input: RemoveCalcChainInput): RemoveCalcChainResult {
  const entries = new Map(input.entries);
  const relationships = new Map(input.relationships);
  const removed = new Set<string>();
  const removedIdsByPart = new Map<string, Set<string>>();

  for (const [partPath, list] of relationships) {
    const removedIds = new Set<string>();
    for (const relationship of list) {
      if (relationship.type.endsWith('/calcChain')) {
        removed.add(relationship.resolvedTarget);
        removedIds.add(relationship.id);
      }
    }
    if (removedIds.size > 0) {
      removedIdsByPart.set(partPath, removedIds);
      relationships.set(
        partPath,
        list.filter((relationship) => !removedIds.has(relationship.id)),
      );
    }
  }
  for (const path of removed) {
    entries.delete(path);
    entries.delete(relsPathForPart(path));
  }
  if (entries.has('xl/calcChain.xml')) {
    removed.add('xl/calcChain.xml');
    entries.delete('xl/calcChain.xml');
    entries.delete('xl/_rels/calcChain.xml.rels');
  }

  const patchedParts: string[] = [];

  // Remove the relationship elements themselves, not just the parsed list.
  for (const [partPath, before] of input.relationships) {
    const removedIds = removedIdsByPart.get(partPath);
    if (removedIds === undefined) continue;
    const bytes = entries.get(partPath);
    if (bytes === undefined) continue;
    let xml = decodeXml(bytes);
    for (const relationship of before) {
      if (!removedIds.has(relationship.id)) continue;
      const pattern = new RegExp(
        `\\s*<Relationship\\b[^>]*\\bId\\s*=\\s*["']${escapeRegExp(relationship.id)}["'][^>]*/>\\s*`,
      );
      xml = xml.replace(pattern, '\n');
    }
    if (!isWellFormedXml(xml)) {
      throw new WorkbookSaveError(
        'unsafe-part',
        `The relationships part ${partPath} could not be patched safely. The file was not written.`,
      );
    }
    entries.set(partPath, encodeXml(xml));
    patchedParts.push(partPath);
  }

  // A restored snapshot may not carry the parsed relationship map. Always
  // sweep the workbook relationships for a calcChain relationship by type.
  if (removed.size > 0 && !patchedParts.includes(input.workbookRelsPath)) {
    const bytes = entries.get(input.workbookRelsPath);
    if (bytes !== undefined) {
      const xml = decodeXml(bytes);
      const patched = xml.replace(
        /\s*<Relationship\b[^>]*\bType\s*=\s*["'][^"']*\/calcChain["'][^>]*\/>\s*/g,
        '\n',
      );
      if (patched !== xml) {
        if (!isWellFormedXml(patched)) {
          throw new WorkbookSaveError(
            'unsafe-part',
            `The relationships part ${input.workbookRelsPath} could not be patched safely. The file was not written.`,
          );
        }
        entries.set(input.workbookRelsPath, encodeXml(patched));
        patchedParts.push(input.workbookRelsPath);
      }
    }
  }

  if (removed.size > 0) {
    const contentTypes = entries.get(input.contentTypesPath);
    if (contentTypes !== undefined) {
      let xml = decodeXml(contentTypes);
      for (const path of removed) {
        const pattern = new RegExp(
          `\\s*<Override\\b[^>]*\\bPartName\\s*=\\s*["']/${escapeRegExp(path)}["'][^>]*/>\\s*`,
        );
        xml = xml.replace(pattern, '\n');
      }
      if (!isWellFormedXml(xml)) {
        throw new WorkbookSaveError(
          'unsafe-part',
          `The content-types part ${input.contentTypesPath} could not be patched safely. The file was not written.`,
        );
      }
      entries.set(input.contentTypesPath, encodeXml(xml));
      patchedParts.push(input.contentTypesPath);
    }
  }

  return { entries, relationships, removedParts: [...removed].sort(), patchedParts };
}

function relsPathForPart(partPath: string): string {
  const slash = partPath.lastIndexOf('/');
  const directory = slash < 0 ? '' : partPath.slice(0, slash + 1);
  const fileName = slash < 0 ? partPath : partPath.slice(slash + 1);
  return `${directory}_rels/${fileName}.rels`;
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
