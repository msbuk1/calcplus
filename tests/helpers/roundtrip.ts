import { readFileSync } from 'node:fs';
import { Buffer } from 'node:buffer';
import { unzipSync } from 'fflate';
import { SaxesParser } from 'saxes';
import { openWorkbook, type OpenedWorkbook } from '@ooxml/open';
import { saveWorkbook, type SaveWorkbookResult } from '@ooxml/save';

/**
 * M6 identity harness (testing-conformance §5).
 *
 * Three checks, each measured independently of the workbook model:
 *
 * - **Identity** — `assertByteIdentical`: a no-op open → save returns the
 *   original file bytes.
 * - **Dirty-part preservation** — `assertOnlyPartsChanged` proves which package
 *   parts changed, and `assertDirtyPartPreserved` proves a rewritten worksheet
 *   kept every unmodelled child, its order, and every untouched row and cell
 *   byte-for-byte.
 * - **Semantic equality** — `semanticManifest` reads a package with its own
 *   saxes pass (styles, number formats, shared strings, values, formulas and
 *   cached results, merges, dimensions, names, views, AutoFilter, conditional
 *   formats, validation, hyperlinks, print setup, protection, relationships),
 *   and `diffSemantic` says exactly what differs between two packages.
 *
 * Nothing here imports the workbook model or the OOXML part parsers, so a
 * regression in either cannot hide behind a shared assumption.
 */

const COMPAT_DIR = new URL('../../fixtures/compat/', import.meta.url);

export function fixtureBytes(name: string): Uint8Array {
  return new Uint8Array(readFileSync(new URL(name, COMPAT_DIR)));
}

export function openFixture(name: string): OpenedWorkbook {
  return openWorkbook(fixtureBytes(name));
}

/** A no-op save: no dirty cells, dimensions, merges or sheet plan. */
export function saveNoOp(opened: OpenedWorkbook): SaveWorkbookResult {
  return saveWorkbook({
    workbook: opened.workbook,
    package: opened.package,
    dirtyCells: new Map(),
  });
}

export function bytesEqual(left: Uint8Array, right: Uint8Array): boolean {
  return Buffer.compare(Buffer.from(left), Buffer.from(right)) === 0;
}

export function assertByteIdentical(
  expected: Uint8Array,
  actual: Uint8Array,
  label = 'output',
): void {
  if (bytesEqual(expected, actual)) return;
  const first = firstDifference(expected, actual);
  throw new Error(
    `${label} is not byte-identical: the files differ at byte ${first} (expected ${expected.byteLength} bytes, got ${actual.byteLength}).`,
  );
}

function firstDifference(left: Uint8Array, right: Uint8Array): number {
  const length = Math.min(left.byteLength, right.byteLength);
  for (let index = 0; index < length; index += 1) {
    if (left[index] !== right[index]) return index;
  }
  return length;
}

/**
 * Part inventory by uncompressed bytes. §5 compares a manifest of name, CRC,
 * uncompressed size and SHA-256; byte equality of every part is strictly
 * stronger and needs no separate digest.
 */
export function partBytesMap(bytes: Uint8Array): Map<string, Uint8Array> {
  // fflate directly, not the package reader: the harness must not share an
  // assumption with the production zip hardening it is checking.
  return new Map(Object.entries(unzipSync(bytes)));
}

/** The decoded XML of one part, for the byte-level dirty-part checks. */
export function partText(bytes: Uint8Array, path: string): string {
  const part = partBytesMap(bytes).get(path);
  if (part === undefined) throw new Error(`The package has no part ${path}.`);
  return decodePart(part);
}

/** Names of the parts whose uncompressed bytes differ between two packages. */
export function changedPartNames(before: Uint8Array, after: Uint8Array): string[] {
  const beforeParts = partBytesMap(before);
  const afterParts = partBytesMap(after);
  const names = new Set([...beforeParts.keys(), ...afterParts.keys()]);
  const changed: string[] = [];
  for (const name of [...names].sort()) {
    const left = beforeParts.get(name);
    const right = afterParts.get(name);
    if (left === undefined || right === undefined || !bytesEqual(left, right)) changed.push(name);
  }
  return changed;
}

/**
 * Assert the saved package carries the same part list and that exactly the
 * declared parts changed. Every other part keeps identical uncompressed bytes,
 * which is the §5 part-inventory rule.
 */
export function assertOnlyPartsChanged(
  before: Uint8Array,
  after: Uint8Array,
  expectedChanged: readonly string[],
): void {
  const actual = changedPartNames(before, after);
  const expected = [...expectedChanged].sort();
  if (actual.join('\n') !== expected.join('\n')) {
    throw new Error(
      `Changed parts do not match the declaration.\nExpected:\n  ${expected.join('\n  ') || '(none)'}\nActual:\n  ${actual.join('\n  ') || '(none)'}`,
    );
  }
}

/**
 * §5 "relationships still resolve": every internal relationship target names a
 * part the package actually carries. External targets are never followed.
 */
export function assertRelationshipsResolve(bytes: Uint8Array): void {
  const parts = partBytesMap(bytes);
  for (const [path, content] of parts) {
    if (!path.endsWith('.rels')) continue;
    const root = parseXmlTree(decodePart(content), path);
    for (const relationship of childrenNamed(root, 'Relationship')) {
      const targetMode = attribute(relationship, 'TargetMode');
      if (targetMode === 'External') continue;
      const target = attribute(relationship, 'Target');
      if (target === null) throw new Error(`A relationship in ${path} has no Target.`);
      const resolved = resolvePartPath(basePartForRelsPath(path), target);
      if (!parts.has(resolved)) {
        throw new Error(`The relationship in ${path} points at a missing part: ${resolved}`);
      }
    }
  }
}

function basePartForRelsPath(relsPath: string): string {
  const marker = '_rels/';
  const index = relsPath.lastIndexOf(marker);
  if (index < 0) return relsPath;
  const directory = relsPath.slice(0, index);
  const file = relsPath.slice(index + marker.length).replace(/\.rels$/, '');
  return `${directory}${file}`;
}

// ---------------------------------------------------------------------------
// A small independent XML reader (saxes only; no calcplus parser).
// ---------------------------------------------------------------------------

export interface XmlAttribute {
  readonly name: string;
  readonly local: string;
  readonly value: string;
}

export interface XmlNode {
  readonly local: string;
  readonly attributes: readonly XmlAttribute[];
  readonly children: XmlNode[];
  text: string;
  readonly start: number;
  openEnd: number;
  end: number;
  selfClosing: boolean;
}

export function parseXmlTree(xml: string, label: string): XmlNode {
  const parser = new SaxesParser({ xmlns: true, position: true });
  const roots: XmlNode[] = [];
  const stack: XmlNode[] = [];
  parser.on('opentag', (tag) => {
    const openEnd = parser.position;
    const node: XmlNode = {
      local: tag.local,
      attributes: Object.values(tag.attributes).map((attribute) => ({
        name: attribute.name,
        local: attribute.local,
        value: attribute.value,
      })),
      children: [],
      text: '',
      start: xml.lastIndexOf('<', Math.max(0, openEnd - 1)),
      openEnd,
      end: openEnd,
      selfClosing: false,
    };
    const parent = stack.at(-1);
    if (parent === undefined) roots.push(node);
    else parent.children.push(node);
    stack.push(node);
  });
  parser.on('closetag', () => {
    const node = stack.pop();
    if (node === undefined) return;
    const end = parser.position;
    node.end = end;
    node.selfClosing = node.openEnd === end;
  });
  parser.on('text', (text) => {
    const node = stack.at(-1);
    if (node !== undefined) node.text += text;
  });
  parser.on('cdata', (text) => {
    const node = stack.at(-1);
    if (node !== undefined) node.text += text;
  });
  parser.on('error', (error) => {
    throw new Error(`Could not parse ${label}: ${error.message}`);
  });
  parser.write(xml).close();
  const root = roots[0];
  if (root === undefined) throw new Error(`${label} has no root element`);
  return root;
}

export function decodePart(bytes: Uint8Array): string {
  return new TextDecoder('utf-8').decode(bytes);
}

export function attribute(node: XmlNode | null, local: string): string | null {
  if (node === null) return null;
  for (const candidate of node.attributes) {
    if (candidate.local === local) return candidate.value;
  }
  return null;
}

export function child(node: XmlNode | null, local: string): XmlNode | null {
  if (node === null) return null;
  return node.children.find((candidate) => candidate.local === local) ?? null;
}

export function childrenNamed(node: XmlNode | null, local: string): XmlNode[] {
  if (node === null) return [];
  return node.children.filter((candidate) => candidate.local === local);
}

export function descendantsNamed(node: XmlNode, local: string): XmlNode[] {
  const found: XmlNode[] = [];
  const queue = [...node.children];
  while (queue.length > 0) {
    const next = queue.shift()!;
    if (next.local === local) found.push(next);
    queue.push(...next.children);
  }
  return found;
}

/** Whitespace-insensitive canonical form, including unmodelled children. */
export function canonicalXml(node: XmlNode): string {
  const attributes = node.attributes
    .map((candidate) => `${candidate.local}=${JSON.stringify(candidate.value)}`)
    .sort()
    .join(' ');
  const children = node.children.map(canonicalXml).join('');
  const text = node.text.trim();
  return `<${node.local} ${attributes}>${text}${children}</${node.local}>`;
}

export function directChildNames(node: XmlNode): string[] {
  return node.children.map((candidate) => candidate.local);
}

export function slicePart(xml: string, node: XmlNode): string {
  return xml.slice(node.start, node.end);
}

// ---------------------------------------------------------------------------
// Dirty-part preservation.
// ---------------------------------------------------------------------------

export interface DirtyPartExpectation {
  /** Cells whose serialised `<c>` element the edit is allowed to replace. */
  readonly changedCells?: readonly string[];
  /** Rows whose serialised `<row>` element the edit is allowed to replace. */
  readonly changedRows?: readonly number[];
  /** Worksheet root children the edit is allowed to replace or insert. */
  readonly rewrittenChildren?: readonly string[];
}

/**
 * Assert a rewritten worksheet part keeps:
 * - every direct child of `<worksheet>` in its original order, with identical
 *   bytes for every child the edit did not rewrite;
 * - every untouched `<row>` element byte-for-byte;
 * - every untouched `<c>` element byte-for-byte.
 */
export function assertDirtyPartPreserved(
  before: string,
  after: string,
  expectation: DirtyPartExpectation = {},
): void {
  const beforeRoot = parseXmlTree(before, 'worksheet before');
  const afterRoot = parseXmlTree(after, 'worksheet after');
  if (directChildNames(beforeRoot).join(',') !== directChildNames(afterRoot).join(',')) {
    throw new Error(
      `Worksheet root children changed.\nBefore: ${directChildNames(beforeRoot).join(', ')}\nAfter:  ${directChildNames(afterRoot).join(', ')}`,
    );
  }

  const rewritten = new Set(expectation.rewrittenChildren ?? []);
  for (const [index, beforeChild] of beforeRoot.children.entries()) {
    if (rewritten.has(beforeChild.local) || beforeChild.local === 'sheetData') continue;
    const afterChild = afterRoot.children[index]!;
    if (slicePart(before, beforeChild) !== slicePart(after, afterChild)) {
      throw new Error(`The unmodelled worksheet child <${beforeChild.local}> changed.`);
    }
  }

  const beforeRows = rowSlices(before, beforeRoot);
  const afterRows = rowSlices(after, afterRoot);
  const afterRowKeys = [...afterRows.keys()];
  let rowCursor = 0;
  for (const number of beforeRows.keys()) {
    const found = afterRowKeys.indexOf(number, rowCursor);
    if (found < 0) throw new Error(`Row ${number} is missing or reordered.`);
    rowCursor = found + 1;
  }
  const changedRows = new Set(expectation.changedRows ?? []);
  for (const [number, beforeSlice] of beforeRows) {
    if (changedRows.has(Number(number))) continue;
    const afterSlice = afterRows.get(number);
    if (afterSlice !== beforeSlice) throw new Error(`Row ${number} is not byte-identical.`);
  }

  const beforeCells = cellSlices(before, beforeRoot);
  const afterCells = cellSlices(after, afterRoot);
  const changedCells = new Set(expectation.changedCells ?? []);
  for (const [ref, beforeSlice] of beforeCells) {
    if (changedCells.has(ref)) continue;
    const afterSlice = afterCells.get(ref);
    if (afterSlice !== beforeSlice) throw new Error(`Cell ${ref} is not byte-identical.`);
  }
}

function rowSlices(xml: string, root: XmlNode): Map<string, string> {
  const rows = new Map<string, string>();
  const sheetData = child(root, 'sheetData');
  for (const row of childrenNamed(sheetData, 'row')) {
    rows.set(attribute(row, 'r') ?? '', xml.slice(row.start, row.end));
  }
  return rows;
}

function cellSlices(xml: string, root: XmlNode): Map<string, string> {
  const cells = new Map<string, string>();
  const sheetData = child(root, 'sheetData');
  for (const row of childrenNamed(sheetData, 'row')) {
    for (const cell of childrenNamed(row, 'c')) {
      cells.set(attribute(cell, 'r') ?? '', xml.slice(cell.start, cell.end));
    }
  }
  return cells;
}

// ---------------------------------------------------------------------------
// Semantic manifest: an independent reading of the parts' meaning.
// ---------------------------------------------------------------------------

export type ManifestValue =
  null | string | number | boolean | ManifestValue[] | { [key: string]: ManifestValue };

export interface WorkbookSemantics {
  readonly date1904: boolean;
  readonly calc: { readonly fullCalcOnLoad: boolean; readonly calcId: string | null };
  readonly protection: ManifestValue;
  readonly names: ManifestValue;
  readonly sheets: ManifestValue;
  readonly sheetList: ManifestValue;
  readonly styles: ManifestValue;
  readonly sharedStrings: ManifestValue;
  readonly relationships: ManifestValue;
  readonly parts: ManifestValue;
  readonly sheetParts: Readonly<Record<string, WorksheetSemantics>>;
}

export interface WorksheetSemantics {
  readonly name: string;
  readonly state: string;
  readonly sheetId: string | null;
  readonly rootChildren: string[];
  readonly dimension: string | null;
  readonly view: ManifestValue;
  readonly pane: ManifestValue;
  readonly sheetFormat: ManifestValue;
  readonly sheetProperties: ManifestValue;
  readonly columns: ManifestValue;
  readonly rows: ManifestValue;
  readonly cells: ManifestValue;
  readonly merges: ManifestValue;
  readonly autoFilter: ManifestValue;
  readonly conditionalFormats: ManifestValue;
  readonly validations: ManifestValue;
  readonly hyperlinks: ManifestValue;
  readonly printOptions: ManifestValue;
  readonly pageMargins: ManifestValue;
  readonly pageSetup: ManifestValue;
  readonly protection: ManifestValue;
  readonly extLst: ManifestValue;
}

export function semanticManifest(bytes: Uint8Array): WorkbookSemantics {
  const parts = partBytesMap(bytes);
  const part = (path: string): XmlNode | null => {
    const content = parts.get(path);
    return content === undefined ? null : parseXmlTree(decodePart(content), path);
  };

  const workbookPath = findWorkbookPath(parts);
  const workbook = part(workbookPath);
  if (workbook === null) throw new Error('The package has no workbook part.');
  const workbookPr = child(workbook, 'workbookPr');
  const calcPr = child(workbook, 'calcPr');
  const protection = child(workbook, 'workbookProtection');

  const relationships = relationshipManifest(parts);
  const workbookRelsPath = relsPathFor(workbookPath);
  const workbookRels = relationshipList(parts.get(workbookRelsPath), workbookRelsPath);

  const sharedStringsPart = resolveRelationship(workbookRels, workbookPath, '/sharedStrings');
  const sharedStrings =
    sharedStringsPart === null ? [] : sharedStringManifest(part(sharedStringsPart));
  const stylesPart = resolveRelationship(workbookRels, workbookPath, '/styles');
  const styles = stylesPart === null ? null : styleManifest(part(stylesPart));

  const sheets: ManifestValue[] = [];
  const sheetParts: Record<string, WorksheetSemantics> = {};
  for (const element of childrenNamed(child(workbook, 'sheets'), 'sheet')) {
    const sheetRelationshipId = relationshipIdOf(element);
    const target = resolveRelationshipId(workbookRels, workbookPath, sheetRelationshipId);
    const sheetRoot = target === null ? null : part(target);
    const name = attribute(element, 'name') ?? '';
    const state = attribute(element, 'state') ?? 'visible';
    const sheetId = attribute(element, 'sheetId');
    sheets.push({
      name,
      state,
      sheetId,
      relationshipId: sheetRelationshipId,
      part: target,
    });
    if (target !== null && sheetRoot !== null) {
      sheetParts[name] = {
        ...worksheetManifest(sheetRoot, sharedStrings),
        name,
        state,
        sheetId,
      };
    }
  }

  const names = childrenNamed(child(workbook, 'definedNames'), 'definedName').map((node) => ({
    name: attribute(node, 'name'),
    localSheetId: attribute(node, 'localSheetId'),
    text: node.text.trim(),
  }));

  return {
    date1904: isTrue(attribute(workbookPr, 'date1904')),
    calc: {
      fullCalcOnLoad: isTrue(attribute(calcPr, 'fullCalcOnLoad')),
      calcId: attribute(calcPr, 'calcId'),
    },
    protection:
      protection === null
        ? null
        : Object.fromEntries(protection.attributes.map((entry) => [entry.local, entry.value])),
    names,
    sheets,
    sheetList: Object.keys(sheetParts),
    styles,
    sharedStrings,
    relationships,
    parts: [...parts.keys()].sort(),
    sheetParts,
  };
}

function findWorkbookPath(parts: ReadonlyMap<string, Uint8Array>): string {
  const rootRels = relationshipList(parts.get('_rels/.rels'), '_rels/.rels');
  const target = resolveRelationship(rootRels, '', '/officeDocument');
  if (target === null) throw new Error('The package root has no officeDocument relationship.');
  return target;
}

interface RelationshipEntry {
  readonly id: string;
  readonly type: string;
  readonly target: string;
  readonly targetMode: string;
}

function relationshipList(
  bytes: Uint8Array | undefined,
  path: string,
): readonly RelationshipEntry[] {
  if (bytes === undefined) return [];
  const root = parseXmlTree(decodePart(bytes), path);
  return childrenNamed(root, 'Relationship').map((node) => ({
    id: attribute(node, 'Id') ?? '',
    type: attribute(node, 'Type') ?? '',
    target: attribute(node, 'Target') ?? '',
    targetMode: attribute(node, 'TargetMode') ?? 'Internal',
  }));
}

function relationshipManifest(
  parts: ReadonlyMap<string, Uint8Array>,
): Record<string, ManifestValue> {
  const manifest: Record<string, ManifestValue> = {};
  for (const [path, bytes] of parts) {
    if (!path.endsWith('.rels')) continue;
    manifest[path] = relationshipList(bytes, path).map((entry) => ({ ...entry }));
  }
  return manifest;
}

function resolveRelationship(
  relationships: readonly RelationshipEntry[],
  basePartPath: string,
  suffix: string,
): string | null {
  const found = relationships.find((entry) => entry.type.endsWith(suffix));
  return found === undefined || found.targetMode === 'External'
    ? null
    : resolvePartPath(basePartPath, found.target);
}

function resolveRelationshipId(
  relationships: readonly RelationshipEntry[],
  basePartPath: string,
  id: string | null,
): string | null {
  if (id === null) return null;
  const found = relationships.find((entry) => entry.id === id);
  return found === undefined || found.targetMode === 'External'
    ? null
    : resolvePartPath(basePartPath, found.target);
}

export function relsPathFor(partPath: string): string {
  const slash = partPath.lastIndexOf('/');
  const directory = slash < 0 ? '' : partPath.slice(0, slash + 1);
  const fileName = slash < 0 ? partPath : partPath.slice(slash + 1);
  return `${directory}_rels/${fileName}.rels`;
}

function resolvePartPath(basePartPath: string, target: string): string {
  if (target.startsWith('/')) return normalise(target.slice(1));
  const base = basePartPath.split('/');
  base.pop();
  return normalise([...base, ...target.split('/')].join('/'));
}

function normalise(path: string): string {
  const segments: string[] = [];
  for (const segment of path.split('/')) {
    if (segment === '' || segment === '.') continue;
    if (segment === '..') {
      segments.pop();
      continue;
    }
    segments.push(segment);
  }
  return segments.join('/');
}

function relationshipIdOf(node: XmlNode): string | null {
  for (const candidate of node.attributes) {
    if (candidate.local === 'id' && candidate.name.includes(':')) return candidate.value;
  }
  return attribute(node, 'id');
}

function sharedStringManifest(root: XmlNode | null): ManifestValue[] {
  if (root === null) return [];
  return childrenNamed(root, 'si').map((entry) => ({
    text: sharedStringText(entry),
    rich: child(entry, 'r') !== null,
  }));
}

function sharedStringText(entry: XmlNode): string {
  let text = '';
  for (const candidate of entry.children) {
    if (candidate.local === 'rPh') continue;
    if (candidate.local === 't') text += candidate.text;
    if (candidate.local === 'r') {
      for (const runText of childrenNamed(candidate, 't')) text += runText.text;
    }
  }
  return text;
}

function styleManifest(root: XmlNode | null): ManifestValue {
  if (root === null) return null;
  const section = (local: string): ManifestValue => {
    for (const candidate of root.children) {
      if (candidate.local === local) return candidate.children.map(canonicalXml);
    }
    return [];
  };
  const attributesOf = (local: string): ManifestValue => {
    for (const candidate of root.children) {
      if (candidate.local === local) {
        return Object.fromEntries(candidate.attributes.map((entry) => [entry.local, entry.value]));
      }
    }
    return null;
  };
  const numberFormats: ManifestValue[] = [];
  for (const candidate of root.children) {
    if (candidate.local !== 'numFmts') continue;
    for (const numFmt of childrenNamed(candidate, 'numFmt')) {
      numberFormats.push({
        id: Number(attribute(numFmt, 'numFmtId') ?? 0),
        code: attribute(numFmt, 'formatCode'),
      });
    }
  }
  return {
    numberFormats,
    fonts: section('fonts'),
    fills: section('fills'),
    borders: section('borders'),
    cellXfs: section('cellXfs'),
    cellStyleXfs: section('cellStyleXfs'),
    dxfs: section('dxfs'),
    cellStyles: section('cellStyles'),
    tableStyles: attributesOf('tableStyles'),
    rootChildren: directChildNames(root),
  };
}

function worksheetManifest(
  root: XmlNode,
  sharedStrings: readonly ManifestValue[],
): WorksheetSemantics {
  const sheetData = child(root, 'sheetData');
  const view = child(child(root, 'sheetViews'), 'sheetView');
  const protection = child(root, 'sheetProtection');
  const extLst = child(root, 'extLst');

  const cells: ManifestValue[] = [];
  for (const row of childrenNamed(sheetData, 'row')) {
    for (const cell of childrenNamed(row, 'c')) {
      cells.push(cellManifest(cell, sharedStrings));
    }
  }

  return {
    name: '',
    state: 'visible',
    sheetId: null,
    rootChildren: directChildNames(root),
    dimension: attribute(child(root, 'dimension'), 'ref'),
    view:
      view === null
        ? null
        : Object.fromEntries(view.attributes.map((entry) => [entry.local, entry.value])),
    pane:
      view === null
        ? null
        : (() => {
            const pane = child(view, 'pane');
            return pane === null
              ? null
              : Object.fromEntries(pane.attributes.map((entry) => [entry.local, entry.value]));
          })(),
    sheetFormat: attributesObject(child(root, 'sheetFormatPr')),
    sheetProperties: child(root, 'sheetPr') === null ? null : canonicalXml(child(root, 'sheetPr')!),
    columns: columnsManifest(root),
    rows: childrenNamed(sheetData, 'row').map((row) => ({
      r: Number(attribute(row, 'r') ?? 0),
      ht: attribute(row, 'ht'),
      hidden: attribute(row, 'hidden'),
      customHeight: attribute(row, 'customHeight'),
      styleId: attribute(row, 's'),
    })),
    cells,
    merges: childrenNamed(child(root, 'mergeCells'), 'mergeCell').map(
      (node) => attribute(node, 'ref') ?? '',
    ),
    autoFilter:
      child(root, 'autoFilter') === null ? null : canonicalXml(child(root, 'autoFilter')!),
    conditionalFormats: childrenNamed(root, 'conditionalFormatting').map(canonicalXml),
    validations:
      child(root, 'dataValidations') === null
        ? null
        : canonicalXml(child(root, 'dataValidations')!),
    hyperlinks:
      child(root, 'hyperlinks') === null ? null : canonicalXml(child(root, 'hyperlinks')!),
    printOptions: attributesObject(child(root, 'printOptions')),
    pageMargins: attributesObject(child(root, 'pageMargins')),
    pageSetup: attributesObject(child(root, 'pageSetup')),
    protection:
      protection === null
        ? null
        : Object.fromEntries(protection.attributes.map((entry) => [entry.local, entry.value])),
    extLst: extLst === null ? null : canonicalXml(extLst),
  };
}

function columnsManifest(root: XmlNode): ManifestValue {
  const columns: ManifestValue[] = [];
  for (const candidate of root.children) {
    if (candidate.local !== 'cols') continue;
    for (const col of childrenNamed(candidate, 'col')) {
      columns.push(Object.fromEntries(col.attributes.map((entry) => [entry.local, entry.value])));
    }
  }
  return columns;
}

function attributesObject(node: XmlNode | null): ManifestValue {
  if (node === null) return null;
  return Object.fromEntries(node.attributes.map((entry) => [entry.local, entry.value]));
}

function cellManifest(cell: XmlNode, sharedStrings: readonly ManifestValue[]): ManifestValue {
  const type = attribute(cell, 't');
  const formula = child(cell, 'f');
  const value = child(cell, 'v');
  const inline = child(cell, 'is');
  const formulaAttributes: Record<string, ManifestValue> = {};
  if (formula !== null) {
    for (const candidate of formula.attributes) {
      formulaAttributes[candidate.local] = candidate.value;
    }
  }
  return {
    ref: attribute(cell, 'r'),
    type,
    styleId: Number(attribute(cell, 's') ?? 0),
    formula: formula === null ? null : formula.text.trim(),
    formulaAttributes: Object.keys(formulaAttributes).length === 0 ? null : formulaAttributes,
    cached: value === null ? null : typedValue(type, value.text, sharedStrings),
    inlineText: inline === null ? null : sharedStringText(inline),
    inlineRich: inline === null ? null : child(inline, 'r') !== null,
  };
}

function typedValue(
  type: string | null,
  raw: string,
  sharedStrings: readonly ManifestValue[],
): ManifestValue {
  switch (type) {
    case 's': {
      const entry = sharedStrings[Number(raw)];
      return entry === undefined ? null : (entry as { text: string }).text;
    }
    case 'str':
    case 'inlineStr':
    case 'd':
      return raw;
    case 'b':
      return raw === '1';
    case 'e':
      return raw;
    case 'n':
    default: {
      if (raw.trim() === '') return null;
      const parsed = Number(raw);
      return Number.isFinite(parsed) ? parsed : raw;
    }
  }
}

function isTrue(value: string | null): boolean {
  return value === '1' || value === 'true' || value === 'on';
}

/**
 * Every path whose value differs between two manifests, in a stable order.
 * An empty list is semantic equality.
 */
export function diffSemantic(left: unknown, right: unknown): string[] {
  const differences: string[] = [];
  walk(left, right, '', differences);
  return differences.sort();
}

function walk(left: unknown, right: unknown, path: string, out: string[]): void {
  if (Object.is(left, right)) return;
  if (typeof left !== typeof right) {
    out.push(path);
    return;
  }
  if (left === null || right === null || typeof left !== 'object' || typeof right !== 'object') {
    out.push(path);
    return;
  }
  if (Array.isArray(left) || Array.isArray(right)) {
    if (!Array.isArray(left) || !Array.isArray(right)) {
      out.push(path);
      return;
    }
    if (left.length !== right.length) {
      out.push(`${path}.length`);
      return;
    }
    for (let index = 0; index < left.length; index += 1) {
      walk(left[index], right[index], `${path}[${index}]`, out);
    }
    return;
  }
  const leftRecord = left as Record<string, unknown>;
  const rightRecord = right as Record<string, unknown>;
  const keys = new Set([...Object.keys(leftRecord), ...Object.keys(rightRecord)]);
  for (const key of [...keys].sort()) {
    walk(leftRecord[key], rightRecord[key], path === '' ? key : `${path}.${key}`, out);
  }
}

/**
 * Assert the semantic difference between two packages is exactly the declared
 * set of paths. The declaration is reviewed like code: an unexpected path is a
 * fidelity failure, and a declared path that no longer differs is a stale
 * expectation.
 */
export function assertSemanticDiff(
  before: Uint8Array,
  after: Uint8Array,
  expectedPaths: readonly string[],
): void {
  const actual = diffSemantic(semanticManifest(before), semanticManifest(after));
  const expected = [...expectedPaths].sort();
  if (actual.join('\n') !== expected.join('\n')) {
    throw new Error(
      `Semantic difference does not match the declaration.\nExpected:\n  ${expected.join('\n  ') || '(none)'}\nActual:\n  ${actual.join('\n  ') || '(none)'}`,
    );
  }
}
