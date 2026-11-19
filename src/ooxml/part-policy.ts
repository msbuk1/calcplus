import type { GridRange, SheetId, Workbook } from '@model/types';
import { parseCellKey, parseRangeRef, rangeContains } from '@model/types';
import { WorkbookSaveError } from './errors';
import type { Relationship, WorkbookPackage } from './types';
import { decodeXml, encodeXml, isWellFormedXml } from './xml';

/**
 * M6 lane 4 — which package parts the save pipeline rewrites, and which it
 * copies verbatim from the source package (file-compat §1-§4, §7, §9 F9-F11).
 *
 * The table below is the contract this module encodes. "Rewritten" means the
 * save pipeline owns those bytes: it may patch them in place, and it rewrites
 * a part only when the model says it changed. "Verbatim" means the part is
 * never serialised from the model; the source bytes travel into the output
 * package unchanged.
 *
 * | Part | Disposition | Why |
 * |---|---|---|
 * | `[Content_Types].xml` | rewritten | sheet and macro changes update overrides |
 * | `xl/workbook.xml` | rewritten | sheet list, names, date system, recalc flag |
 * | `xl/_rels/workbook.xml.rels` | rewritten | sheet and macro relationships |
 * | `xl/worksheets/*` | rewritten | the patch writer's position/tree modes |
 * | `xl/styles.xml` | rewritten | append-only style records |
 * | `xl/sharedStrings.xml` | rewritten | append-only string entries |
 * | `xl/calcChain.xml` | rewritten | removed when formula storage changes |
 * | `_rels/.rels`, other `.rels` | verbatim | never changes, except a delete |
 * | `xl/theme/*` | verbatim | resolved for display, never written |
 * | `docProps/core.xml`, `docProps/app.xml` | verbatim | not modelled in v1 |
 * | `docProps/custom.xml` | verbatim | custom properties are preserved |
 * | charts, chartsheets, drawings, VML, media, OLE embeddings | verbatim | Tier B; anchors and series transform in the model, and those transformed bytes are written back |
 * | pivot tables and pivot caches | verbatim | never refreshed; source ranges block |
 * | external links | verbatim | never resolved or refreshed |
 * | `vbaProject.bin`, `vbaData.xml` | verbatim | never executed or edited; removed only by an explicit `.xlsx` Save As |
 * | signature parts | verbatim | preserved; the save invalidates them |
 * | tables, slicers, timelines, custom XML | verbatim | not modelled in v1 |
 *
 * Composition with the transformed-opaque overlay in `save.ts`: the overlay
 * writes a model part whose bytes differ from the package original (a chart's
 * series or a drawing's anchors after a structural edit). A verbatim part whose
 * model bytes are identical is never written, which is why identity saves stay
 * byte-identical. The policy supplies the pipeline-owned part set so the
 * overlay never writes a save-managed part from a stale model mirror.
 *
 * The second half of the module is the report listing (§7): every preserved
 * part with its five-category status, plus the staleness rules the contract
 * states — a pivot cache is stale after an edit inside its source range, and
 * an external link is never refreshed. The compatibility and export reports
 * (lane 6) consume this listing; nothing here renders UI.
 */

export type SavedFormat = 'xlsx' | 'xlsm';

export type PartCategory =
  | 'contentTypes'
  | 'relationships'
  | 'workbook'
  | 'worksheet'
  | 'styles'
  | 'sharedStrings'
  | 'calcChain'
  | 'theme'
  | 'documentProperties'
  | 'customProperties'
  | 'chart'
  | 'chartSheet'
  | 'drawing'
  | 'vmlDrawing'
  | 'media'
  | 'oleObject'
  | 'pivotTable'
  | 'pivotCache'
  | 'externalLink'
  | 'vbaProject'
  | 'signature'
  | 'table'
  | 'slicer'
  | 'customXml'
  | 'other';

export type PartDisposition = 'rewritten' | 'verbatim' | 'removed';

/** Part paths the save pipeline writes, keyed by the package's own locations. */
export interface PackagePartLocations {
  readonly contentTypesPath: string;
  readonly workbookPartPath: string;
  readonly workbookRelsPath: string;
  readonly stylesPath: string | null;
  readonly sharedStringsPath: string | null;
  readonly sheetPartPaths: ReadonlyMap<SheetId, string>;
}

export interface ClassifiedPart {
  readonly path: string;
  readonly category: PartCategory;
  readonly description: string;
  readonly disposition: PartDisposition;
}

export interface PartPolicy {
  /** The format the save writes. */
  readonly format: SavedFormat;
  /** The format the package was opened in. */
  readonly sourceFormat: SavedFormat;
  readonly parts: readonly ClassifiedPart[];
  /**
   * Parts the save pipeline owns. This is the overlay's skip list, so it also
   * carries the save-managed contract paths the package may not hold right now
   * (a calcChain a previous save removed): a stale model mirror of one of
   * those parts must never be written back.
   */
  readonly pipelineOwnedPartPaths: ReadonlySet<string>;
  /** Parts copied from the source package with identical bytes. */
  readonly verbatimPartPaths: ReadonlySet<string>;
  /** Parts an explicit format change drops. */
  readonly removedPartPaths: ReadonlySet<string>;
}

export interface PartPolicyInput {
  readonly package: WorkbookPackage;
  /** Defaults to the package entries; a save passes its working map. */
  readonly parts?: ReadonlyMap<string, Uint8Array>;
  /** Defaults to the source package's own format. */
  readonly format?: SavedFormat;
}

const DESCRIPTION: Readonly<Record<PartCategory, string>> = {
  contentTypes: 'Content types',
  relationships: 'Package relationships',
  workbook: 'Workbook part',
  worksheet: 'Worksheet part',
  styles: 'Style table',
  sharedStrings: 'Shared-string table',
  calcChain: 'Calculation chain',
  theme: 'Theme',
  documentProperties: 'Document properties',
  customProperties: 'Custom document properties',
  chart: 'Chart',
  chartSheet: 'Chart sheet',
  drawing: 'Drawing',
  vmlDrawing: 'Legacy drawing (VML)',
  media: 'Image or media',
  oleObject: 'Embedded OLE object',
  pivotTable: 'Pivot table',
  pivotCache: 'Pivot cache',
  externalLink: 'External link',
  vbaProject: 'VBA project',
  signature: 'Digital signature',
  table: 'Table',
  slicer: 'Slicer or timeline',
  customXml: 'Custom XML',
  other: 'Unrecognised package part',
};

/** §7 wording for what calcplus will and will not do with each category. */
const NOTE: Readonly<Record<PartCategory, string>> = {
  contentTypes: 'Maintained by the save pipeline.',
  relationships: 'Maintained by the save pipeline.',
  workbook: 'Maintained by the save pipeline.',
  worksheet: 'Maintained by the save pipeline.',
  styles: 'Append-only; existing indices never change.',
  sharedStrings: 'Append-only; existing indices never change.',
  calcChain: 'Not trusted; removed when formula storage changes.',
  theme: 'Preserved; theme colours are resolved for display.',
  documentProperties: 'Preserved byte-for-byte.',
  customProperties: 'Preserved byte-for-byte; not editable.',
  chart: 'Preserved but not rendered in this milestone; series references transform.',
  chartSheet: 'Preserved but not rendered in this milestone.',
  drawing: 'Preserved byte-for-byte; anchors transform with structural edits.',
  vmlDrawing: 'Preserved byte-for-byte; note anchors transform with structural edits.',
  media: 'Preserved byte-for-byte; not rendered in this milestone.',
  oleObject: 'Preserved byte-for-byte; never activated.',
  pivotTable: 'Preserved; never refreshed; output cells show the stored values.',
  pivotCache: 'Preserved; never refreshed.',
  externalLink: 'Preserved; never opened, resolved or refreshed.',
  vbaProject: 'Preserved byte-for-byte; macros are never executed or edited.',
  signature: 'Preserved; saving invalidates it.',
  table: 'Preserved; structured references are not evaluated in this milestone.',
  slicer: 'Preserved; not displayed in this milestone.',
  customXml: 'Preserved; not editable in this milestone.',
  other: 'Preserved byte-for-byte.',
};

const PIPELINE_OWNED_CATEGORIES: ReadonlySet<PartCategory> = new Set<PartCategory>([
  'contentTypes',
  'workbook',
  'worksheet',
  'styles',
  'sharedStrings',
  'calcChain',
]);

const MACRO_PART = /(^|\/)vbaProject\.bin$/i;
const VBA_DATA_PART = /(^|\/)vbaData\.xml$/i;
const VBA_SIGNATURE_PART = /(^|\/)vbaProjectSignature\.bin$/i;

/** Whether a part belongs to the VBA project (never written into an `.xlsx`). */
export function isMacroPartPath(path: string): boolean {
  return MACRO_PART.test(path) || VBA_DATA_PART.test(path);
}

function segments(path: string): readonly string[] {
  return path.split('/');
}

function hasSegment(path: string, segment: string): boolean {
  return segments(path).includes(segment);
}

export function partCategoryOf(path: string, locations: PackagePartLocations): PartCategory {
  if (path === locations.contentTypesPath) return 'contentTypes';
  if (path === locations.workbookPartPath) return 'workbook';
  if (path === locations.workbookRelsPath) return 'relationships';
  if (locations.stylesPath !== null && path === locations.stylesPath) return 'styles';
  if (locations.sharedStringsPath !== null && path === locations.sharedStringsPath) {
    return 'sharedStrings';
  }
  for (const sheetPath of locations.sheetPartPaths.values()) {
    if (path === sheetPath) return 'worksheet';
  }
  // Every relationship part is a relationship part, whatever directory holds it.
  if (path.endsWith('.rels')) return 'relationships';
  // A worksheet part the package map does not know yet (a sheet added in this
  // session) is still a worksheet, never an unknown part.
  if (hasSegment(path, 'worksheets')) return 'worksheet';
  if (MACRO_PART.test(path)) return 'vbaProject';
  if (VBA_DATA_PART.test(path)) return 'vbaProject';
  if (VBA_SIGNATURE_PART.test(path)) return 'signature';
  if (hasSegment(path, '_xmlsignatures') || path.endsWith('.sigs')) return 'signature';
  if (hasSegment(path, 'chartsheets')) return 'chartSheet';
  if (hasSegment(path, 'charts')) return 'chart';
  if (hasSegment(path, 'drawings')) {
    return path.toLowerCase().endsWith('.vml') ? 'vmlDrawing' : 'drawing';
  }
  if (hasSegment(path, 'media')) return 'media';
  if (hasSegment(path, 'embeddings')) return 'oleObject';
  if (hasSegment(path, 'pivotTables')) return 'pivotTable';
  if (hasSegment(path, 'pivotCache')) return 'pivotCache';
  if (hasSegment(path, 'externalLinks')) return 'externalLink';
  if (hasSegment(path, 'tables')) return 'table';
  if (hasSegment(path, 'slicers') || hasSegment(path, 'timelines')) return 'slicer';
  if (hasSegment(path, 'customXml')) return 'customXml';
  if (hasSegment(path, 'theme')) return 'theme';
  if (path === 'docProps/custom.xml') return 'customProperties';
  if (path === 'docProps/core.xml' || path === 'docProps/app.xml') return 'documentProperties';
  if (path.endsWith('calcChain.xml')) return 'calcChain';
  return 'other';
}

/** The package's own macro-enabled state, read from its macro parts. */
export function sourceFormatOf(parts: ReadonlyMap<string, Uint8Array>): SavedFormat {
  for (const path of parts.keys()) {
    if (MACRO_PART.test(path) || VBA_DATA_PART.test(path)) return 'xlsm';
  }
  return 'xlsx';
}

/**
 * Classify every part. `removed` is applied only when an explicit Save As
 * changes the format from macro-enabled to plain, which is the one path that
 * may drop macro parts (file-compat §6.3, §7 "Removed").
 */
export function planPartPolicy(input: PartPolicyInput): PartPolicy {
  const parts = input.parts ?? input.package.entries;
  const locations: PackagePartLocations = input.package;
  const sourceFormat = sourceFormatOf(input.package.entries);
  const format = input.format ?? sourceFormat;
  const dropMacros = sourceFormat === 'xlsm' && format === 'xlsx';

  const classified: ClassifiedPart[] = [];
  const pipelineOwnedPartPaths = new Set<string>();
  const verbatimPartPaths = new Set<string>();
  const removedPartPaths = new Set<string>();
  // The save-managed contract paths, whether or not this package holds them.
  // `calcChain` is deleted by the save itself when formula storage changes, so
  // a model mirror that still has it must never be written back (workbook-model
  // §5, file-compat §4).
  pipelineOwnedPartPaths.add(locations.contentTypesPath);
  pipelineOwnedPartPaths.add(locations.workbookPartPath);
  pipelineOwnedPartPaths.add(locations.workbookRelsPath);
  pipelineOwnedPartPaths.add('xl/calcChain.xml');
  for (const sheetPath of locations.sheetPartPaths.values()) pipelineOwnedPartPaths.add(sheetPath);
  if (locations.stylesPath !== null) pipelineOwnedPartPaths.add(locations.stylesPath);
  if (locations.sharedStringsPath !== null) {
    pipelineOwnedPartPaths.add(locations.sharedStringsPath);
  }
  for (const path of parts.keys()) {
    const category = partCategoryOf(path, locations);
    const disposition: PartDisposition =
      dropMacros && category === 'vbaProject'
        ? 'removed'
        : PIPELINE_OWNED_CATEGORIES.has(category) || path === locations.workbookRelsPath
          ? 'rewritten'
          : 'verbatim';
    classified.push({
      path,
      category,
      description: DESCRIPTION[category],
      disposition,
    });
    if (disposition === 'rewritten') pipelineOwnedPartPaths.add(path);
    else if (disposition === 'verbatim') verbatimPartPaths.add(path);
    else removedPartPaths.add(path);
  }

  classified.sort((left, right) => left.path.localeCompare(right.path));
  return {
    format,
    sourceFormat,
    parts: classified,
    pipelineOwnedPartPaths,
    verbatimPartPaths,
    removedPartPaths,
  };
}

// ---------------------------------------------------------------------------
// Macro removal for an explicit `.xlsm` -> `.xlsx` Save As (file-compat §6.3)
// ---------------------------------------------------------------------------

export interface RemoveMacroPartsInput {
  readonly entries: ReadonlyMap<string, Uint8Array>;
  readonly relationships: ReadonlyMap<string, readonly Relationship[]>;
  readonly contentTypesPath: string;
  readonly workbookPartPath: string;
  readonly workbookRelsPath: string;
}

export interface RemoveMacroPartsResult {
  readonly entries: Map<string, Uint8Array>;
  readonly relationships: Map<string, Relationship[]>;
  readonly removedParts: readonly string[];
  readonly patchedParts: readonly string[];
  readonly changed: boolean;
}

/**
 * Drop the macro parts, their relationships and their content-type overrides,
 * and turn the workbook's main content type back into the plain `.xlsx` one.
 * Pure: the caller owns when this runs. The result is only for a confirmed
 * format change; the default save keeps the macro parts.
 */
export function removeMacroParts(input: RemoveMacroPartsInput): RemoveMacroPartsResult {
  const entries = new Map(input.entries);
  const relationships = new Map<string, Relationship[]>();
  for (const [partPath, list] of input.relationships) relationships.set(partPath, [...list]);
  const removed = new Set<string>(planPartPolicyRemovals(entries));

  const removedIdsByPart = new Map<string, Set<string>>();
  for (const [partPath, list] of input.relationships) {
    const ids = new Set<string>();
    for (const relationship of list) {
      if (
        relationship.type.endsWith('/vbaProject') ||
        relationship.type.endsWith('/vbaData') ||
        MACRO_PART.test(relationship.resolvedTarget) ||
        VBA_DATA_PART.test(relationship.resolvedTarget)
      ) {
        removed.add(relationship.resolvedTarget);
        ids.add(relationship.id);
      }
    }
    if (ids.size > 0) {
      removedIdsByPart.set(partPath, ids);
      relationships.set(
        partPath,
        list.filter((relationship) => !ids.has(relationship.id)),
      );
    }
  }

  for (const path of removed) {
    entries.delete(path);
    entries.delete(relsPathForPart(path));
  }

  const patchedParts: string[] = [];
  for (const [partPath, ids] of removedIdsByPart) {
    const bytes = entries.get(partPath);
    if (bytes === undefined) continue;
    let xml = decodeXml(bytes);
    for (const id of ids) {
      xml = xml.replace(
        new RegExp(`\\s*<Relationship\\b[^>]*\\bId\\s*=\\s*["']${escapeRegExp(id)}["'][^>]*/>\\s*`),
        '\n',
      );
    }
    if (!isWellFormedXml(xml)) {
      throw new WorkbookSaveError(
        'unsafe-part',
        `The relationships part ${partPath} could not drop the macro parts safely. The file was not written.`,
      );
    }
    entries.set(partPath, encodeXml(xml));
    patchedParts.push(partPath);
  }

  const contentTypes = entries.get(input.contentTypesPath);
  if (removed.size > 0 && contentTypes !== undefined) {
    let xml = decodeXml(contentTypes);
    for (const path of removed) {
      xml = xml.replace(
        new RegExp(
          `\\s*<(?:[\\w.-]+:)?Override\\b[^>]*\\bPartName\\s*=\\s*["']/${escapeRegExp(path)}["'][^>]*/>\\s*`,
          'g',
        ),
        '\n',
      );
    }
    // Turn the workbook's main content type back into the plain `.xlsx` one.
    // Rewriting the whole element leaves the attribute order irrelevant.
    xml = xml.replace(
      new RegExp(
        `<(?:[\\w.-]+:)?Override\\b[^>]*\\bPartName\\s*=\\s*["']/${escapeRegExp(
          input.workbookPartPath,
        )}["'][^>]*/>`,
      ),
      `<Override PartName="/${input.workbookPartPath}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>`,
    );
    if (!isWellFormedXml(xml)) {
      throw new WorkbookSaveError(
        'unsafe-part',
        `The content-types part ${input.contentTypesPath} could not drop the macro parts safely. The file was not written.`,
      );
    }
    entries.set(input.contentTypesPath, encodeXml(xml));
    if (!patchedParts.includes(input.contentTypesPath)) patchedParts.push(input.contentTypesPath);
  }

  return {
    entries,
    relationships,
    removedParts: [...removed].sort(),
    patchedParts,
    changed: removed.size > 0,
  };
}

function planPartPolicyRemovals(entries: ReadonlyMap<string, Uint8Array>): readonly string[] {
  const removed: string[] = [];
  for (const path of entries.keys()) {
    if (MACRO_PART.test(path) || VBA_DATA_PART.test(path)) removed.push(path);
  }
  return removed;
}

function relsPathForPart(partPath: string): string {
  const slash = partPath.lastIndexOf('/');
  const directory = slash < 0 ? '' : partPath.slice(0, slash + 1);
  const fileName = slash < 0 ? partPath : partPath.slice(slash + 1);
  return `${directory}_rels/${fileName}.rels`;
}

function escapeRegExp(text: string): string {
  return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

// ---------------------------------------------------------------------------
// Preserved-part listing for the compatibility and export reports (§7)
// ---------------------------------------------------------------------------

export type PartReportStatus = 'unchanged' | 'transformed' | 'stale' | 'removed' | 'warning';

export type PartReportReason =
  | 'pivot-cache-not-refreshed'
  | 'external-link-not-refreshed'
  | 'macros-removed'
  | 'digital-signature-invalidated';

export interface PreservedPartReport {
  readonly path: string;
  readonly category: PartCategory;
  readonly description: string;
  /** What calcplus will and will not do with the part (§7). */
  readonly note: string;
  readonly status: PartReportStatus;
  readonly reason: PartReportReason | null;
  /** The sheet the part is anchored on, when the relationship graph names one. */
  readonly sheetId?: SheetId;
}

export interface PreservedPartListingInput {
  readonly package: WorkbookPackage;
  readonly workbook: Workbook;
  readonly entries?: ReadonlyMap<string, Uint8Array>;
  /** Content changes since the last save; used for pivot-cache staleness. */
  readonly dirtyCells?: ReadonlyMap<SheetId, ReadonlySet<string>>;
  /** Parts the save rewrote (a chart's series, a drawing's anchors). */
  readonly transformedParts?: readonly string[];
  /** Defaults to the source package's own format. */
  readonly format?: SavedFormat;
}

/** Categories the reports list (§7: Tier B content, links, macros, signatures). */
const REPORTED_CATEGORIES: ReadonlySet<PartCategory> = new Set<PartCategory>([
  'chart',
  'chartSheet',
  'drawing',
  'vmlDrawing',
  'media',
  'oleObject',
  'pivotTable',
  'pivotCache',
  'externalLink',
  'vbaProject',
  'signature',
  'table',
  'slicer',
  'customXml',
  'customProperties',
]);

/**
 * The preserved parts, sorted by path, with the five §7 categories resolved.
 * A no-op save lists every part as `unchanged` except an external link, which
 * is always `stale` because calcplus never refreshes it (§6.2), and a
 * signature, which is a `warning` because any save invalidates it (§6.5).
 */
export function listPreservedParts(input: PreservedPartListingInput): PreservedPartReport[] {
  // Classify the union of the source package and the caller's entries: a part
  // the save removed is absent from the saved package but must still be
  // reported as `removed` (§7).
  const parts = new Map(input.package.entries);
  if (input.entries !== undefined) {
    for (const [path, bytes] of input.entries) parts.set(path, bytes);
  }
  const policy = planPartPolicy({ package: input.package, parts, format: input.format });
  const transformed = new Set(input.transformedParts ?? []);
  const pivotSources = pivotCacheSources(parts);
  const dirty = input.dirtyCells ?? new Map<SheetId, ReadonlySet<string>>();
  const listing: PreservedPartReport[] = [];

  for (const part of policy.parts) {
    if (!REPORTED_CATEGORIES.has(part.category)) continue;
    const reason = reportReason(part.path, part.category, input, policy, dirty, pivotSources);
    listing.push({
      path: part.path,
      category: part.category,
      description: part.description,
      note: NOTE[part.category],
      status: statusOf(part.category, part.disposition, transformed.has(part.path), reason),
      reason,
      ...sheetAnchorOf(part.path, part.category, input),
    });
  }
  return listing;
}

/** Every preserved part the listing resolves to `stale` (§7). */
export function listStaleParts(input: PreservedPartListingInput): PreservedPartReport[] {
  return listPreservedParts(input).filter((entry) => entry.status === 'stale');
}

function statusOf(
  category: PartCategory,
  disposition: PartDisposition,
  transformed: boolean,
  reason: PartReportReason | null,
): PartReportStatus {
  if (disposition === 'removed') return 'removed';
  if (category === 'signature') return 'warning';
  if (transformed) return 'transformed';
  if (reason !== null) return 'stale';
  return 'unchanged';
}

function reportReason(
  path: string,
  category: PartCategory,
  input: PreservedPartListingInput,
  policy: PartPolicy,
  dirty: ReadonlyMap<SheetId, ReadonlySet<string>>,
  pivotSources: readonly PivotCacheSource[],
): PartReportReason | null {
  if (policy.removedPartPaths.has(path)) return 'macros-removed';
  if (category === 'signature') return 'digital-signature-invalidated';
  if (category === 'externalLink') return 'external-link-not-refreshed';
  if (category === 'pivotCache') {
    for (const source of pivotSources) {
      if (source.path !== path) continue;
      if (pivotSourceChanged(source, input.workbook, dirty)) return 'pivot-cache-not-refreshed';
    }
  }
  return null;
}

interface PivotCacheSource {
  readonly path: string;
  readonly sheetName: string | null;
  readonly range: GridRange;
}

/** `<cacheSource><worksheetSource ref sheet/>` of every pivot cache definition. */
function pivotCacheSources(parts: ReadonlyMap<string, Uint8Array>): PivotCacheSource[] {
  const sources: PivotCacheSource[] = [];
  for (const [path, bytes] of parts) {
    if (!hasSegment(path, 'pivotCache')) continue;
    if (!path.endsWith('.xml') || path.includes('/_rels/')) continue;
    const xml = decodeXml(bytes);
    if (!xml.includes('worksheetSource')) continue;
    const pattern = /<(?:[A-Za-z_][\w.-]*:)?worksheetSource\b([^>]*?)\/?>/g;
    for (const match of xml.matchAll(pattern)) {
      const attributes = match[1] ?? '';
      const ref = xmlAttribute(attributes, 'ref');
      if (ref === null) continue;
      const range = parseRangeAttribute(ref);
      if (range === null) continue;
      sources.push({ path, sheetName: xmlAttribute(attributes, 'sheet'), range });
    }
  }
  return sources;
}

function pivotSourceChanged(
  source: PivotCacheSource,
  workbook: Workbook,
  dirty: ReadonlyMap<SheetId, ReadonlySet<string>>,
): boolean {
  const sheets =
    source.sheetName === null
      ? workbook.sheets
      : workbook.sheets.filter((sheet) => sheet.name === source.sheetName);
  for (const sheet of sheets) {
    const cells = dirty.get(sheet.id);
    if (cells === undefined) continue;
    for (const key of cells) {
      const address = parseCellKey(key);
      if (rangeContains(source.range, address.row, address.col)) return true;
    }
  }
  return false;
}

function sheetAnchorOf(
  path: string,
  category: PartCategory,
  input: PreservedPartListingInput,
): { readonly sheetId?: SheetId } {
  for (const sheet of input.workbook.sheets) {
    if (sheet.opaqueParts.has(path)) return { sheetId: sheet.id };
  }
  if (category !== 'pivotTable' && category !== 'externalLink') return {};
  for (const [pathKey, relationships] of input.package.relationships) {
    const owner = sheetForRelsPath(pathKey, input.workbook, input.package);
    if (owner === undefined) continue;
    if (relationships.some((relationship) => relationship.resolvedTarget === path)) {
      return { sheetId: owner };
    }
  }
  return {};
}

function sheetForRelsPath(
  relsPath: string,
  workbook: Workbook,
  pkg: WorkbookPackage,
): SheetId | undefined {
  for (const sheet of workbook.sheets) {
    const partPath = pkg.sheetPartPaths.get(sheet.id) ?? sheet.partPath;
    if (relsPathForPart(partPath) === relsPath) return sheet.id;
  }
  return undefined;
}

function xmlAttribute(attributes: string, name: string): string | null {
  const match = new RegExp(`(?:^|\\s)${name}\\s*=\\s*(["'])([\\s\\S]*?)\\1`).exec(attributes);
  return match === null ? null : decodeAttributeEntities(match[2] ?? '');
}

function decodeAttributeEntities(value: string): string {
  return value
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');
}

function parseRangeAttribute(text: string): GridRange | null {
  try {
    return parseRangeRef(text.replace(/\$/g, ''));
  } catch {
    return null;
  }
}
