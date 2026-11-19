import { unzipSync } from 'fflate';
import { SaxesParser } from 'saxes';

/**
 * Repair-free package verification (file-compat §9 F22, §10 criteria 3-4;
 * testing-conformance §5).
 *
 * This is the local half of the "opens without repair" check: it proves the
 * package is well formed before Excel is ever opened. It answers four
 * questions without sharing a parser with the production open path:
 *
 * 1. Is the ZIP readable, with unique, safe part names?
 * 2. Is `[Content_Types].xml` complete: does every part have a content type,
 *    and does every override name a part that exists?
 * 3. Does every internal relationship resolve to a part the package carries,
 *    and does the workbook's sheet list resolve?
 * 4. Is every XML part well formed, with no DTD?
 *
 * The Excel 365 repair prompt itself stays a manual run on the maintainer's
 * machine (file-compat §9); this module is what the fixture tests and
 * `scripts/verify-repair-free.mjs` run locally.
 */

export interface VerifyProblem {
  readonly code: string;
  readonly message: string;
  /** The package part the problem belongs to, when it has one. */
  readonly part?: string;
}

export interface VerifyResult {
  readonly ok: boolean;
  readonly problems: readonly VerifyProblem[];
  readonly parts: readonly string[];
}

interface RelationshipRecord {
  readonly id: string | null;
  readonly type: string | null;
  readonly target: string | null;
  readonly targetMode: string | null;
}

interface ContentTypes {
  readonly defaults: Map<string, string>;
  readonly overrides: Map<string, string>;
}

const XML_DECODER = new TextDecoder('utf-8', { fatal: false });
const RELATIONSHIPS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const SPREADSHEET_MAIN_TYPES = new Set([
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
  'application/vnd.ms-excel.sheet.macroEnabled.main+xml',
  'application/vnd.ms-excel.sheet.binary.macroEnabled.main',
]);

/** Verify one package image. Never throws; every failure becomes a problem. */
export function verifyPackage(bytes: Uint8Array, label = 'package'): VerifyResult {
  const problems: VerifyProblem[] = [];
  const report = (code: string, message: string, part?: string): void => {
    problems.push({ code, message, ...(part === undefined ? {} : { part }) });
  };

  let parts: Map<string, Uint8Array>;
  try {
    parts = new Map(Object.entries(unzipSync(bytes)).filter(([name]) => !name.endsWith('/')));
  } catch (error) {
    report('zip-unreadable', `${label}: the ZIP container cannot be read (${messageOf(error)}).`);
    return { ok: false, problems, parts: [] };
  }

  checkPartNames(bytes, parts, report);
  const xml = new Map<string, string>();
  for (const [name, content] of parts) {
    if (!isXmlPart(name)) continue;
    const text = XML_DECODER.decode(content);
    if (hasDoctype(text)) {
      report('xml-doctype', `${label}: ${name} declares a DTD, which is never accepted.`, name);
      continue;
    }
    const error = xmlWellFormedError(text);
    if (error !== null) {
      report('xml-malformed', `${label}: ${name} is not well-formed XML (${error}).`, name);
      continue;
    }
    xml.set(name, text);
  }

  const contentTypes = checkContentTypes(parts, xml, report);
  checkRelationships(parts, xml, report);
  checkWorkbook(parts, xml, contentTypes, report);

  return { ok: problems.length === 0, problems, parts: [...parts.keys()] };
}

// ---------------------------------------------------------------------------
// ZIP structure
// ---------------------------------------------------------------------------

function checkPartNames(
  bytes: Uint8Array,
  parts: ReadonlyMap<string, Uint8Array>,
  report: (code: string, message: string, part?: string) => void,
): void {
  for (const name of parts.keys()) {
    if (name === '') {
      report('zip-name', 'The package carries a part with an empty name.');
      continue;
    }
    if (name.startsWith('/') || name.includes('\\')) {
      report('zip-name', `The part name ${JSON.stringify(name)} is not a relative OPC path.`, name);
      continue;
    }
    const segments = name.split('/');
    if (segments.some((segment) => segment === '' || segment === '.' || segment === '..')) {
      report('zip-name', `The part name ${JSON.stringify(name)} has an unsafe path segment.`, name);
    }
  }
  const counts = new Map<string, number>();
  for (const name of centralDirectoryNames(bytes)) {
    counts.set(name, (counts.get(name) ?? 0) + 1);
  }
  for (const [name, count] of counts) {
    if (count > 1) {
      report('zip-duplicate', `The part ${name} appears ${count} times in the archive.`, name);
    }
  }
}

/**
 * Names from the central directory. `unzipSync` keeps the last duplicate, so
 * the archive's own records are scanned when there is something to check. A
 * Zip64 package (not produced by any calcplus fixture) skips this check rather
 * than reading a 32-bit field as a real offset.
 */
function centralDirectoryNames(bytes: Uint8Array): string[] {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let eocd = -1;
  for (let offset = bytes.length - 22; offset >= 0; offset -= 1) {
    if (view.getUint32(offset, true) === 0x06054b50) {
      eocd = offset;
      break;
    }
  }
  if (eocd < 0) return [];
  const entryCount = view.getUint16(eocd + 10, true);
  const directorySize = view.getUint32(eocd + 12, true);
  const directoryOffset = view.getUint32(eocd + 16, true);
  if (
    entryCount === 0xffff ||
    directorySize === 0xffffffff ||
    directoryOffset === 0xffffffff ||
    directoryOffset + directorySize > bytes.length
  ) {
    return [];
  }
  const names: string[] = [];
  let offset = directoryOffset;
  for (let index = 0; index < entryCount; index += 1) {
    if (offset + 46 > bytes.length || view.getUint32(offset, true) !== 0x02014b50) break;
    const nameLength = view.getUint16(offset + 28, true);
    const extraLength = view.getUint16(offset + 30, true);
    const commentLength = view.getUint16(offset + 32, true);
    if (offset + 46 + nameLength > bytes.length) break;
    names.push(XML_DECODER.decode(bytes.subarray(offset + 46, offset + 46 + nameLength)));
    offset += 46 + nameLength + extraLength + commentLength;
  }
  return names;
}

function isXmlPart(name: string): boolean {
  const lower = name.toLowerCase();
  return lower.endsWith('.xml') || lower.endsWith('.rels') || lower.endsWith('.vml');
}

// ---------------------------------------------------------------------------
// Content types
// ---------------------------------------------------------------------------

function checkContentTypes(
  parts: ReadonlyMap<string, Uint8Array>,
  xml: ReadonlyMap<string, string>,
  report: (code: string, message: string, part?: string) => void,
): ContentTypes | null {
  const contentTypes: ContentTypes = { defaults: new Map(), overrides: new Map() };
  const text = xml.get('[Content_Types].xml');
  if (text === undefined) {
    report('content-types-missing', 'The package has no readable [Content_Types].xml.');
    return null;
  }
  for (const element of elementsNamed(text, 'Default')) {
    const extension = (element.attributes.get('Extension') ?? '').toLowerCase();
    const type = element.attributes.get('ContentType') ?? '';
    if (extension === '' || type === '') {
      report('content-type-default', 'A Default content type is missing its extension or type.');
      continue;
    }
    if (contentTypes.defaults.has(extension)) {
      report(
        'content-type-default',
        `The extension ${extension} has more than one default content type.`,
      );
    }
    contentTypes.defaults.set(extension, type);
  }
  for (const element of elementsNamed(text, 'Override')) {
    const partName = element.attributes.get('PartName') ?? '';
    const type = element.attributes.get('ContentType') ?? '';
    if (!partName.startsWith('/') || partName === '/') {
      report(
        'content-type-override',
        `An override has an invalid PartName ${JSON.stringify(partName)}.`,
      );
      continue;
    }
    if (type === '') {
      report('content-type-override', `The override for ${partName} has no ContentType.`, partName);
    }
    if (contentTypes.overrides.has(partName)) {
      report('content-type-override', `The part ${partName} has more than one override.`, partName);
    }
    contentTypes.overrides.set(partName, type);
  }

  for (const name of parts.keys()) {
    const type = contentTypeOf(contentTypes, name);
    if (type === null) {
      report(
        'content-type-missing',
        `The part ${name} has no content type: no override and no default for its extension.`,
        name,
      );
    }
  }
  for (const partName of contentTypes.overrides.keys()) {
    if (!parts.has(partName.slice(1))) {
      report(
        'content-type-dangling',
        `An override names a part the package does not carry: ${partName}.`,
        partName.slice(1),
      );
    }
  }
  return contentTypes;
}

function contentTypeOf(contentTypes: ContentTypes, partName: string): string | null {
  const override = contentTypes.overrides.get(`/${partName}`);
  if (override !== undefined) return override;
  const slash = partName.lastIndexOf('.');
  if (slash < 0) return null;
  return contentTypes.defaults.get(partName.slice(slash + 1).toLowerCase()) ?? null;
}

// ---------------------------------------------------------------------------
// Relationships
// ---------------------------------------------------------------------------

function checkRelationships(
  parts: ReadonlyMap<string, Uint8Array>,
  xml: ReadonlyMap<string, string>,
  report: (code: string, message: string, part?: string) => void,
): void {
  for (const name of parts.keys()) {
    if (!name.toLowerCase().endsWith('.rels')) continue;
    const text = xml.get(name);
    if (text === undefined) continue; // Already reported as malformed.
    const basePart = relsBasePart(name);
    if (basePart !== '' && !parts.has(basePart)) {
      report(
        'rels-dangling-base',
        `The relationship part ${name} describes a missing part (${basePart}).`,
        name,
      );
    }
    const relationships = parseRelationships(text);
    const ids = new Set<string>();
    if (relationships.length > 0 && !text.includes(RELATIONSHIPS_NS)) {
      report(
        'rels-namespace',
        `${name} does not declare the package relationships namespace.`,
        name,
      );
    }
    for (const relationship of relationships) {
      if (relationship.id === null || relationship.type === null || relationship.target === null) {
        report('rels-attribute', `A relationship in ${name} is missing Id, Type or Target.`, name);
        continue;
      }
      if (ids.has(relationship.id)) {
        report(
          'rels-duplicate-id',
          `${name} repeats the relationship id ${relationship.id}.`,
          name,
        );
        continue;
      }
      ids.add(relationship.id);
      if (relationship.targetMode === 'External') continue;
      const resolved = resolveTarget(basePart, relationship.target);
      if (resolved === null) {
        report(
          'rels-unsafe-target',
          `${name}: the target ${relationship.target} leaves the package.`,
          name,
        );
        continue;
      }
      if (!parts.has(resolved)) {
        report(
          'rels-missing-target',
          `${name}: the relationship ${relationship.id} points at a missing part (${resolved}).`,
          name,
        );
      }
    }
  }
}

export function relsBasePart(relsPath: string): string {
  const marker = '_rels/';
  const markerIndex = relsPath.lastIndexOf(marker);
  if (markerIndex < 0) return '';
  const directory = relsPath.slice(0, markerIndex);
  const file = relsPath.slice(markerIndex + marker.length).replace(/\.rels$/i, '');
  return `${directory}${file}`;
}

export function resolveTarget(basePart: string, target: string): string | null {
  let raw = target;
  try {
    raw = decodeURIComponent(target);
  } catch {
    raw = target;
  }
  const baseDirectory = basePart.includes('/')
    ? basePart.slice(0, basePart.lastIndexOf('/') + 1)
    : '';
  const combined = raw.startsWith('/') ? raw.slice(1) : `${baseDirectory}${raw}`;
  const segments: string[] = [];
  for (const segment of combined.split('/')) {
    if (segment === '' || segment === '.') continue;
    if (segment === '..') {
      if (segments.length === 0) return null;
      segments.pop();
      continue;
    }
    segments.push(segment);
  }
  return segments.join('/');
}

function parseRelationships(text: string): RelationshipRecord[] {
  const relationships: RelationshipRecord[] = [];
  const parser = new SaxesParser({ xmlns: false });
  parser.on('opentag', (tag) => {
    if (tag.name !== 'Relationship') return;
    const attributes = tag.attributes as Record<string, string>;
    relationships.push({
      id: attribute(attributes, 'Id'),
      type: attribute(attributes, 'Type'),
      target: attribute(attributes, 'Target'),
      targetMode: attribute(attributes, 'TargetMode'),
    });
  });
  try {
    parser.write(text).close();
  } catch {
    return relationships;
  }
  return relationships;
}

function attribute(attributes: Record<string, string>, name: string): string | null {
  for (const [key, value] of Object.entries(attributes)) {
    if (key === name || key.endsWith(`:${name}`)) {
      return typeof value === 'string' ? value : null;
    }
  }
  return null;
}

// ---------------------------------------------------------------------------
// Workbook and sheet list
// ---------------------------------------------------------------------------

function checkWorkbook(
  parts: ReadonlyMap<string, Uint8Array>,
  xml: ReadonlyMap<string, string>,
  contentTypes: ContentTypes | null,
  report: (code: string, message: string, part?: string) => void,
): void {
  const rootText = xml.get('_rels/.rels');
  if (rootText === undefined) {
    report('root-rels-missing', 'The package has no readable _rels/.rels part.');
    return;
  }
  const officeDocument = parseRelationships(rootText).find(
    (relationship) => relationship.type?.endsWith('/officeDocument') === true,
  );
  if (officeDocument === undefined || officeDocument.target === null) {
    report('workbook-relationship-missing', 'The package root has no officeDocument relationship.');
    return;
  }
  const workbookPath = resolveTarget('', officeDocument.target);
  if (workbookPath === null) {
    report('workbook-relationship-unsafe', 'The officeDocument target leaves the package.');
    return;
  }
  if (!parts.has(workbookPath)) {
    report('workbook-missing', `The workbook part ${workbookPath} is missing.`, workbookPath);
    return;
  }
  if (contentTypes !== null) {
    const type = contentTypeOf(contentTypes, workbookPath);
    if (type === null || !SPREADSHEET_MAIN_TYPES.has(type)) {
      report(
        'workbook-content-type',
        `The workbook part ${workbookPath} does not carry a spreadsheet main content type.`,
        workbookPath,
      );
    }
  }

  const workbookText = xml.get(workbookPath);
  if (workbookText === undefined) return;
  const workbookRelsPath = relsPathFor(workbookPath);
  const workbookRelsText = xml.get(workbookRelsPath);
  if (workbookRelsText === undefined) {
    report(
      'workbook-rels-missing',
      `The workbook part has no relationships part (${workbookRelsPath}).`,
    );
    return;
  }
  const relationships = parseRelationships(workbookRelsText);
  const byId = new Map(
    relationships.flatMap((relationship) =>
      relationship.id === null ? [] : [[relationship.id, relationship] as const],
    ),
  );
  const names = new Set<string>();
  const sheetIds = new Set<string>();
  let sheetCount = 0;
  let worksheetCount = 0;
  for (const element of elementsNamed(workbookText, 'sheet')) {
    sheetCount += 1;
    const name = element.attributes.get('name') ?? '';
    const sheetId = element.attributes.get('sheetId') ?? '';
    const relationshipId = element.attributes.get('r:id') ?? element.attributes.get('id') ?? '';
    if (name === '') report('sheet-name', 'A worksheet has no name.', workbookPath);
    if (names.has(name)) {
      report(
        'sheet-name-duplicate',
        `The sheet name ${JSON.stringify(name)} appears twice.`,
        workbookPath,
      );
    }
    names.add(name);
    if (sheetId !== '') {
      if (sheetIds.has(sheetId)) {
        report('sheet-id-duplicate', `The sheet id ${sheetId} appears twice.`, workbookPath);
      }
      sheetIds.add(sheetId);
    }
    const relationship = byId.get(relationshipId);
    if (relationship === undefined) {
      report(
        'sheet-relationship-missing',
        `The sheet ${JSON.stringify(name)} names a missing relationship (${relationshipId}).`,
        workbookPath,
      );
      continue;
    }
    if (relationship.type?.endsWith('/worksheet') === true) worksheetCount += 1;
    if (relationship.target === null) continue;
    const target = resolveTarget(workbookPath, relationship.target);
    if (target === null || !parts.has(target)) {
      report(
        'sheet-part-missing',
        `The sheet ${JSON.stringify(name)} resolves to a missing part (${target ?? relationship.target}).`,
        workbookPath,
      );
    }
  }
  if (sheetCount === 0) {
    report('sheet-list-empty', 'The workbook has no sheets.', workbookPath);
  }
  if (worksheetCount === 0 && sheetCount > 0) {
    report('sheet-list-empty', 'No sheet resolves to a worksheet part.', workbookPath);
  }
}

export function relsPathFor(partPath: string): string {
  const slash = partPath.lastIndexOf('/');
  const directory = slash < 0 ? '' : partPath.slice(0, slash + 1);
  const fileName = slash < 0 ? partPath : partPath.slice(slash + 1);
  return `${directory}_rels/${fileName}.rels`;
}

// ---------------------------------------------------------------------------
// A minimal independent XML reader
// ---------------------------------------------------------------------------

interface XmlElement {
  readonly attributes: Map<string, string>;
}

/** Every element with the given local name, in document order. */
function elementsNamed(xml: string, localName: string): XmlElement[] {
  const elements: XmlElement[] = [];
  const parser = new SaxesParser({ xmlns: false });
  parser.on('opentag', (tag) => {
    if (localNameOf(tag.name) !== localName) return;
    const attributes = new Map<string, string>();
    for (const [key, value] of Object.entries(tag.attributes as Record<string, string>)) {
      attributes.set(key, String(value));
    }
    elements.push({ attributes });
  });
  try {
    parser.write(xml).close();
  } catch {
    return elements;
  }
  return elements;
}

function localNameOf(qualified: string): string {
  const colon = qualified.indexOf(':');
  return colon < 0 ? qualified : qualified.slice(colon + 1);
}

function xmlWellFormedError(xml: string): string | null {
  const parser = new SaxesParser({ xmlns: false });
  try {
    parser.write(xml).close();
    return null;
  } catch (error) {
    return messageOf(error);
  }
}

function hasDoctype(xml: string): boolean {
  return /<!DOCTYPE/i.test(xml);
}

function messageOf(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}
