import type { SheetId, Workbook } from '@model/types';
import type { Relationship } from '../types';
import { isFilterDatabaseName } from '../autofilter-xml';
import { WorkbookSaveError } from '../errors';
import {
  decodeXml,
  encodeXml,
  escapeXmlAttribute,
  isWellFormedXml,
  openTagStart,
  parseXml,
} from '../xml';

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

interface DefinedNameElement {
  readonly name: string;
  readonly localSheetId: string | null;
  readonly start: number;
  readonly openEnd: number;
  readonly closeStart: number;
  readonly end: number;
  readonly selfClosing: boolean;
  readonly prefix: string;
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

/**
 * A name the model owns and may author, such as the sheet's
 * `_xlnm.Print_Area` system name (print-pdf §4). A `null` formula removes a
 * name the source has; a formula creates one the source lacks or rewrites the
 * one it has. Every other defined name stays on the preserve/rename path.
 */
export interface AuthoredDefinedName {
  readonly name: string;
  readonly sheetId: SheetId;
  readonly formula: string | null;
}

export interface WorkbookDefinedNamePatchInput {
  readonly bytes: Uint8Array;
  readonly partName: string;
  readonly workbook: Workbook;
  /**
   * Source `<sheet>` index -> stable sheet id at open time. A sheet-scoped
   * name stores its owner as `localSheetId`, an index that shifts whenever a
   * sheet is inserted or deleted.
   */
  readonly sourceSheetIds: readonly (SheetId | null)[];
  /** Sheets removed in this save; their local names leave with them. */
  readonly removedSheetIds: ReadonlySet<SheetId>;
  /** Sheet-scoped names the model authors, with `null` meaning "remove". */
  readonly authoredNames?: readonly AuthoredDefinedName[];
}

/**
 * Rewrite the text of `<definedName>` elements whose model formula changed
 * (a sheet rename or delete rewrites references in place), recompute
 * `localSheetId` after an insert or delete, and remove names owned by a
 * deleted sheet. An authored system name (the print area) is created,
 * retargeted or removed too. Everything else in the workbook part stays
 * byte-for-byte as authored. Returns `null` when nothing changed, which keeps
 * a no-op save untouched.
 */
export function patchWorkbookDefinedNames(input: WorkbookDefinedNamePatchInput): Uint8Array | null {
  const xml = decodeXml(input.bytes);
  const parsed = parseDefinedNameElements(xml, input.partName);
  const { records } = parsed;
  const authored = authoredNameMap(input.authoredNames ?? []);
  const handledAuthored = new Set<string>();

  const currentIndexBySheetId = new Map(
    input.workbook.sheets.map((sheet, index) => [sheet.id, index] as const),
  );
  const edits: { start: number; end: number; text: string }[] = [];
  for (const record of records) {
    const sourceIndex = record.localSheetId === null ? null : Number(record.localSheetId);
    const ownerSheetId =
      sourceIndex === null || !Number.isInteger(sourceIndex)
        ? null
        : (input.sourceSheetIds[sourceIndex] ?? null);

    if (ownerSheetId !== null && input.removedSheetIds.has(ownerSheetId)) {
      edits.push({ start: record.start, end: record.end, text: '' });
      continue;
    }

    // A name the model authors is settled by the second pass, which also
    // knows how to create one the source lacks.
    if (authored !== null && authored.has(nameKey(ownerSheetId, record.name))) continue;

    if (isFilterDatabaseName(record.name)) {
      // File-compat §3: the AutoFilter feature owns this name's text, so a
      // range change is written by the filter writer, not here. Only the
      // `localSheetId` index, which moves with the sheet list, is maintained.
      const currentIndex =
        ownerSheetId === null ? null : (currentIndexBySheetId.get(ownerSheetId) ?? null);
      if (currentIndex !== null && Number(record.localSheetId) !== currentIndex) {
        edits.push({
          start: record.start,
          end: record.openEnd,
          text: replaceDefinedNameAttribute(
            xml.slice(record.start, record.openEnd),
            'localSheetId',
            String(currentIndex),
          ),
        });
      }
      continue;
    }

    const currentSheet =
      ownerSheetId === null
        ? null
        : (input.workbook.sheets.find((sheet) => sheet.id === ownerSheetId) ?? null);
    const currentIndex =
      ownerSheetId === null ? null : (currentIndexBySheetId.get(ownerSheetId) ?? null);
    const modelName =
      ownerSheetId === null
        ? input.workbook.names.find((name) => name.name.toUpperCase() === record.name.toUpperCase())
        : currentSheet?.names.find((name) => name.name.toUpperCase() === record.name.toUpperCase());
    if (modelName === undefined) continue;

    const formulaChanged = modelName.formula !== record.text;
    const indexChanged =
      ownerSheetId !== null &&
      currentIndex !== null &&
      Number(record.localSheetId) !== currentIndex;
    if (!formulaChanged && !indexChanged) continue;

    const openTag = xml.slice(record.start, record.openEnd);
    const updatedOpenTag = indexChanged
      ? replaceDefinedNameAttribute(openTag, 'localSheetId', String(currentIndex))
      : openTag;
    if (record.selfClosing) {
      const closing = record.prefix === '' ? '</definedName>' : `</${record.prefix}:definedName>`;
      const open = updatedOpenTag.replace(/\/\s*>$/, '>');
      edits.push({
        start: record.start,
        end: record.end,
        text: `${open}${escapeNameText(modelName.formula)}${closing}`,
      });
      continue;
    }
    if (indexChanged) {
      edits.push({
        start: record.start,
        end: record.openEnd,
        text: updatedOpenTag,
      });
    }
    if (formulaChanged) {
      edits.push({
        start: record.openEnd,
        end: record.closeStart,
        text: escapeNameText(modelName.formula),
      });
    }
  }

  if (authored !== null) {
    for (const record of records) {
      const sourceIndex = record.localSheetId === null ? null : Number(record.localSheetId);
      const ownerSheetId =
        sourceIndex === null || !Number.isInteger(sourceIndex)
          ? null
          : (input.sourceSheetIds[sourceIndex] ?? null);
      if (ownerSheetId === null || input.removedSheetIds.has(ownerSheetId)) continue;
      const key = nameKey(ownerSheetId, record.name);
      const entry = authored.get(key);
      if (entry === undefined) continue;
      handledAuthored.add(key);
      const currentIndex = currentIndexBySheetId.get(ownerSheetId);
      if (currentIndex === undefined) continue;
      if (entry.formula === null) {
        edits.push({ start: record.start, end: record.end, text: '' });
        continue;
      }
      const indexChanged = Number(record.localSheetId) !== currentIndex;
      const formulaChanged = record.text !== entry.formula;
      if (!indexChanged && !formulaChanged) continue;
      const openTag = xml.slice(record.start, record.openEnd);
      const updatedOpenTag = indexChanged
        ? replaceDefinedNameAttribute(openTag, 'localSheetId', String(currentIndex))
        : openTag;
      if (record.selfClosing) {
        const closing = record.prefix === '' ? '</definedName>' : `</${record.prefix}:definedName>`;
        const open = updatedOpenTag.replace(/\/\s*>$/, '>');
        edits.push({
          start: record.start,
          end: record.end,
          text: `${open}${escapeNameText(entry.formula)}${closing}`,
        });
        continue;
      }
      if (indexChanged) {
        edits.push({ start: record.start, end: record.openEnd, text: updatedOpenTag });
      }
      if (formulaChanged) {
        edits.push({
          start: record.openEnd,
          end: record.closeStart,
          text: escapeNameText(entry.formula),
        });
      }
    }

    const created = [...authored.entries()]
      .filter(([key, entry]) => entry.formula !== null && !handledAuthored.has(key))
      .map(([, entry]) => entry)
      .sort(
        (left, right) =>
          (currentIndexBySheetId.get(left.sheetId) ?? 0) -
          (currentIndexBySheetId.get(right.sheetId) ?? 0),
      )
      .map((entry) => {
        const index = currentIndexBySheetId.get(entry.sheetId);
        if (index === undefined) return '';
        const prefix = parsed.container?.prefix ?? parsed.prefix;
        const name = qualified(prefix, 'definedName');
        return `<${name} name="${escapeXmlAttribute(entry.name)}" localSheetId="${index}" hidden="1">${escapeNameText(entry.formula!)}</${name}>`;
      })
      .join('');
    if (created !== '') {
      const container = parsed.container;
      if (container === null) {
        const prefix = parsed.prefix;
        const containerTag = qualified(prefix, 'definedNames');
        edits.push({
          start: parsed.definedNamesInsert,
          end: parsed.definedNamesInsert,
          text: `<${containerTag}>${created}</${containerTag}>`,
        });
      } else if (container.selfClosing) {
        const open = xml.slice(container.start, container.openEnd).replace(/\/\s*>$/, '>');
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
  }

  if (edits.length === 0) return null;

  let patched = xml;
  for (const edit of [...edits].sort((left, right) => right.start - left.start)) {
    patched = patched.slice(0, edit.start) + edit.text + patched.slice(edit.end);
  }
  if (!isWellFormedXml(patched)) {
    throw new WorkbookSaveError(
      'unsafe-part',
      `The workbook part ${input.partName} could not store its defined names safely. The file was not written.`,
    );
  }
  return encodeXml(patched);
}

/**
 * Element text needs only `&`, `<` and `>` escaped. Leaving apostrophes and
 * double quotes literal keeps a defined name looking exactly like the one
 * Excel writes, so retargeting a range changes the fewest bytes
 * (file-compat §9 F23; the AutoFilter name writer does the same).
 */
function escapeNameText(text: string): string {
  return text.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;');
}

function replaceDefinedNameAttribute(openTag: string, localName: string, value: string): string {
  const pattern = new RegExp(`(\\s${localName}\\s*=\\s*)("[^"]*"|'[^']*')`);
  if (pattern.test(openTag)) {
    return openTag.replace(pattern, (_match: string, prefix: string) => `${prefix}"${value}"`);
  }
  const close = openTag.lastIndexOf('>');
  if (close < 0) return openTag;
  const insertAt = openTag[close - 1] === '/' ? close - 1 : close;
  return `${openTag.slice(0, insertAt)} ${localName}="${value}"${openTag.slice(insertAt)}`;
}

/** Children of `<workbook>` that follow `<definedNames>` in schema order. */
const AFTER_DEFINED_NAMES: ReadonlySet<string> = new Set([
  'calcPr',
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

function parseDefinedNameElements(
  xml: string,
  partName: string,
): {
  readonly prefix: string;
  readonly records: DefinedNameElement[];
  readonly container: DefinedNamesContainer | null;
  readonly definedNamesInsert: number;
} {
  const records: DefinedNameElement[] = [];
  let prefix = '';
  let depth = 0;
  let definedNamesInsert = -1;
  let container: DefinedNamesContainer | null = null;
  let current: {
    name: string;
    localSheetId: string | null;
    start: number;
    openEnd: number;
    prefix: string;
    text: string;
  } | null = null;

  parseXml(
    xml,
    {
      onOpenTag(tag, position) {
        depth += 1;
        if (depth === 1) {
          prefix = tag.prefix;
          return;
        }
        if (depth === 2 && definedNamesInsert < 0 && AFTER_DEFINED_NAMES.has(tag.local)) {
          definedNamesInsert = openTagStart(xml, position);
        }
        if (depth === 2 && tag.local === 'definedNames' && container === null) {
          container = {
            prefix: tag.prefix,
            start: openTagStart(xml, position),
            openEnd: position,
            contentEnd: position,
            end: position,
            selfClosing: false,
          };
          return;
        }
        if (depth === 3 && tag.local === 'definedName' && container !== null) {
          current = {
            name: attributeValueOf(tag, 'name') ?? '',
            localSheetId: attributeValueOf(tag, 'localSheetId') ?? null,
            start: openTagStart(xml, position),
            openEnd: position,
            prefix: tag.prefix,
            text: '',
          };
        }
      },
      onCloseTag(tag, start, end) {
        if (depth === 2 && tag.local === 'definedNames' && container !== null) {
          const open = container as DefinedNamesContainer;
          container = {
            ...open,
            contentEnd: start,
            end,
            selfClosing: open.openEnd === end,
          };
        } else if (depth === 3 && tag.local === 'definedName' && current !== null) {
          records.push({
            ...current,
            closeStart: start,
            end,
            selfClosing: current.openEnd === end,
          });
          current = null;
        }
        depth -= 1;
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
  return {
    prefix,
    records,
    container,
    definedNamesInsert:
      definedNamesInsert < 0 ? (closing < 0 ? xml.length : closing) : definedNamesInsert,
  };
}

function authoredNameMap(
  entries: readonly AuthoredDefinedName[],
): Map<string, AuthoredDefinedName> | null {
  if (entries.length === 0) return null;
  return new Map(entries.map((entry) => [nameKey(entry.sheetId, entry.name), entry] as const));
}

function nameKey(sheetId: SheetId | null, name: string): string {
  return `${sheetId ?? ''}\u0000${name.toUpperCase()}`;
}

function qualified(prefix: string, local: string): string {
  return prefix === '' ? local : `${prefix}:${local}`;
}

function attributeValueOf(
  tag: { attributes: Record<string, { local: string; value: string }> },
  local: string,
): string | undefined {
  for (const attribute of Object.values(tag.attributes)) {
    if (attribute.local === local) return attribute.value;
  }
  return undefined;
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
