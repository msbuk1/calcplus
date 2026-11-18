import type { DefinedName, Sheet, SheetId, Workbook } from '@model/types';
import type { ConditionalFormatRule } from '@model/conditional-format';
import { createDefaultStyle, createWorkbookId } from '@model/workbook';
import { isLocaleBuiltinNumberFormat } from '@model/style';
import { WorkbookOpenError } from './errors';
import { parseStylesXml } from './styles-xml';
import { parseSharedStringsXml } from './shared-strings';
import type { Relationship, SharedStringsTable, WorkbookPackage } from './types';
import { parseWorkbookXml } from './workbook-xml';
import { parseWorksheetXml } from './worksheet-xml';
import { parseRelationships, findRelationship, findRelationshipById } from './relationships';
import { readPackage } from './zip';

export interface CompatibilityInfo {
  readonly opaquePartPaths: readonly string[];
  readonly externalLinkParts: readonly string[];
  readonly hasMacros: boolean;
  readonly hasDigitalSignatures: boolean;
  readonly notes: readonly string[];
}

export interface OpenedWorkbook {
  readonly workbook: Workbook;
  readonly package: WorkbookPackage;
  readonly initialSheetId: string;
  readonly compatibility: CompatibilityInfo;
  /**
   * Each sheet's parsed conditional-format rules, base and x14 merged (§7).
   * `Sheet.conditionalFormats` holds the base records for byte preservation;
   * this map holds the rules the grid evaluates and the report lists.
   */
  readonly conditionalFormatRules: ReadonlyMap<SheetId, readonly ConditionalFormatRule[]>;
}

export function openWorkbook(bytes: Uint8Array, fileName?: string): OpenedWorkbook {
  if (fileName !== undefined) {
    const extension = fileName.slice(fileName.lastIndexOf('.')).toLowerCase();
    if (extension === '.xls' || extension === '.xlsb' || extension === '.ods') {
      throw new WorkbookOpenError(
        'unsupported-format',
        'This file format is not supported. Open it in Excel and save as .xlsx.',
      );
    }
  }
  const read = readPackage(bytes);
  const entries = read.entries;

  const contentTypesPath = '[Content_Types].xml';
  if (!entries.has(contentTypesPath)) {
    throw new WorkbookOpenError('corrupt', 'This file is damaged or is not a valid workbook.');
  }
  if (!entries.has('_rels/.rels')) {
    throw new WorkbookOpenError('corrupt', 'This file is damaged or is not a valid workbook.');
  }

  const relationshipsByPart = new Map<string, Relationship[]>();
  for (const [name, partBytes] of entries) {
    if (!name.endsWith('.rels')) continue;
    const basePart = relsBasePart(name);
    relationshipsByPart.set(name, parseRelationships(partBytes, name, basePart));
  }

  const rootRelationships = relationshipsByPart.get('_rels/.rels') ?? [];
  const officeDocument = findRelationship(rootRelationships, '/officeDocument');
  if (officeDocument === undefined || officeDocument.targetMode !== 'Internal') {
    throw new WorkbookOpenError('corrupt', 'This file is damaged or is not a valid workbook.');
  }
  const workbookPartPath = officeDocument.resolvedTarget;
  const workbookBytes = entries.get(workbookPartPath);
  if (workbookBytes === undefined) {
    throw new WorkbookOpenError('corrupt', 'The workbook part is missing from the package.');
  }

  const workbookRelsPath = relsPathForPart(workbookPartPath);
  const workbookRelationships = relationshipsByPart.get(workbookRelsPath) ?? [];
  const workbookInfo = parseWorkbookXml(workbookBytes, workbookPartPath);

  const stylesRelationship = findRelationship(workbookRelationships, '/styles');
  const stylesPath = stylesRelationship?.resolvedTarget ?? null;
  const stylesBytes = stylesPath === null ? undefined : entries.get(stylesPath);
  const stylesInfo =
    stylesBytes === undefined ? null : parseStylesXml(stylesBytes, stylesPath ?? 'styles.xml');

  const sharedStringsRelationship = findRelationship(workbookRelationships, '/sharedStrings');
  const sharedStringsPath = sharedStringsRelationship?.resolvedTarget ?? null;
  const sharedStringsBytes =
    sharedStringsPath === null ? undefined : entries.get(sharedStringsPath);
  let sharedStrings: SharedStringsTable | null = null;
  if (sharedStringsPath !== null && sharedStringsBytes !== undefined) {
    sharedStrings = parseSharedStringsXml(sharedStringsBytes, sharedStringsPath).table;
  }

  const themeRelationship = findRelationship(workbookRelationships, '/theme');
  const themePath = themeRelationship?.resolvedTarget ?? null;
  const themeBytes = themePath === null ? undefined : entries.get(themePath);

  const sheets: Sheet[] = [];
  const sheetPartPaths = new Map<string, string>();
  const sheetRelationshipIds = new Map<string, string>();
  const sheetIdByWorkbookIndex: string[] = [];
  const conditionalFormatRules = new Map<SheetId, readonly ConditionalFormatRule[]>();
  for (const [index, reference] of workbookInfo.sheets.entries()) {
    const relationship = findRelationshipById(workbookRelationships, reference.relationshipId);
    if (relationship === undefined || relationship.targetMode !== 'Internal') {
      throw new WorkbookOpenError(
        'corrupt',
        `The relationship for sheet ${reference.name} is missing or external.`,
      );
    }
    const partPath = relationship.resolvedTarget;
    const partBytes = entries.get(partPath);
    if (partBytes === undefined) {
      throw new WorkbookOpenError('corrupt', `The worksheet part ${partPath} is missing.`);
    }
    const parsed = parseWorksheetXml(partBytes, {
      sharedStrings,
      partPath,
      name: reference.name,
      indexedPalette: stylesInfo?.indexedPalette ?? null,
    });
    const sheet = parsed.sheet;
    const worksheetRelsPath = relsPathForPart(partPath);
    for (const relationship of relationshipsByPart.get(worksheetRelsPath) ?? []) {
      if (relationship.targetMode !== 'Internal') continue;
      const targetBytes = entries.get(relationship.resolvedTarget);
      if (targetBytes === undefined) continue;
      sheet.opaqueParts.set(relationship.resolvedTarget, {
        path: relationship.resolvedTarget,
        bytes: targetBytes,
        anchor: null,
      });
    }
    sheets.push(sheet);
    sheetPartPaths.set(sheet.id, partPath);
    sheetRelationshipIds.set(sheet.id, reference.relationshipId);
    sheetIdByWorkbookIndex[index] = sheet.id;
    conditionalFormatRules.set(sheet.id, parsed.conditionalFormatRules);
  }

  const workbookNames: DefinedName[] = [];
  for (const name of workbookInfo.definedNames) {
    const pending = name.sheetId;
    if (pending !== null && pending.startsWith('pending-sheet-index:')) {
      const index = Number(pending.slice('pending-sheet-index:'.length));
      const sheet = sheets[index];
      if (sheet === undefined) {
        // A malformed localSheetId keeps a workbook-scoped fallback so the
        // name text is never lost on save.
        workbookNames.push({ ...name, sheetId: null });
      } else {
        sheet.names.push({ ...name, sheetId: sheet.id });
      }
      continue;
    }
    workbookNames.push(name);
  }

  const workbook: Workbook = {
    id: createWorkbookId(),
    sheets,
    names: workbookNames,
    styles: stylesInfo?.styles ?? [createDefaultStyle(0)],
    theme: { path: themePath, bytes: themeBytes ?? null },
    namedStyles: stylesInfo?.namedStyles ?? [],
    protection: null,
    opaqueParts: new Map(),
    sourceParts: new Map(),
    indexedPalette: stylesInfo?.indexedPalette ?? null,
    date1904: workbookInfo.date1904,
  };

  const sourcePartPaths = new Set<string>([workbookPartPath]);
  if (stylesPath !== null) sourcePartPaths.add(stylesPath);
  if (sharedStringsPath !== null) sourcePartPaths.add(sharedStringsPath);
  if (themePath !== null) sourcePartPaths.add(themePath);
  for (const partPath of sheetPartPaths.values()) sourcePartPaths.add(partPath);

  for (const [name, partBytes] of entries) {
    if (sourcePartPaths.has(name)) {
      workbook.sourceParts.set(name, partBytes);
    } else {
      workbook.opaqueParts.set(name, { path: name, bytes: partBytes, anchor: null });
    }
  }

  const externalLinkParts = [...entries.keys()].filter(
    (name) =>
      name.startsWith('xl/externalLinks/') ||
      workbookRelationships.some(
        (relationship) =>
          relationship.type.endsWith('/externalLink') && relationship.resolvedTarget === name,
      ),
  );
  const hasMacros = [...entries.keys()].some(
    (name) =>
      name.toLowerCase().endsWith('vbaproject.bin') || name.toLowerCase().endsWith('vbadata.xml'),
  );
  const hasDigitalSignatures = [...entries.keys()].some(
    (name) => name.startsWith('_xmlsignatures/') || name.endsWith('.sigs'),
  );

  const compatibility: CompatibilityInfo = {
    opaquePartPaths: [...workbook.opaqueParts.keys()].sort(),
    externalLinkParts,
    hasMacros,
    hasDigitalSignatures,
    notes: buildCompatibilityNotes(workbook, hasMacros, hasDigitalSignatures, externalLinkParts),
  };

  const activeTab = Math.min(Math.max(0, workbookInfo.activeTab), Math.max(0, sheets.length - 1));
  const initialSheetId = sheetIdByWorkbookIndex[activeTab] ?? sheets[0]!.id;

  const pkg: WorkbookPackage = {
    originalBytes: bytes,
    entries,
    contentTypesPath,
    rootRelsPath: '_rels/.rels',
    workbookPartPath,
    workbookRelsPath,
    sharedStringsPath,
    stylesPath,
    themePath,
    sheetPartPaths,
    sheetRelationshipIds,
    relationships: relationshipsByPart,
    sharedStrings,
    externalLinkParts,
  };

  return {
    workbook,
    package: pkg,
    initialSheetId,
    compatibility,
    conditionalFormatRules,
  };
}

function buildCompatibilityNotes(
  workbook: Workbook,
  hasMacros: boolean,
  hasDigitalSignatures: boolean,
  externalLinkParts: readonly string[],
): string[] {
  const notes: string[] = [];
  const opaque = [...workbook.opaqueParts.keys()];
  if (opaque.some((path) => path.startsWith('xl/charts/'))) {
    notes.push('Charts are preserved but not rendered in this milestone.');
  }
  if (opaque.some((path) => path.startsWith('xl/drawings/'))) {
    notes.push('Drawings and images are preserved byte-for-byte.');
  }
  if (externalLinkParts.length > 0) {
    notes.push('External links are preserved and never refreshed.');
  }
  if (hasMacros) {
    notes.push('Macros are preserved byte-for-byte and never executed.');
  }
  if (hasDigitalSignatures) {
    notes.push('Digital signatures are preserved, but saving invalidates them.');
  }
  if (workbook.styles.some((style) => isLocaleBuiltinNumberFormat(style.numFmtId))) {
    notes.push(
      'Locale-specific built-in number formats are shown in their Gregorian en-US form; the original numFmtId is preserved.',
    );
  }
  return notes;
}

function relsPathForPart(partPath: string): string {
  const slash = partPath.lastIndexOf('/');
  const directory = slash < 0 ? '' : partPath.slice(0, slash + 1);
  const fileName = slash < 0 ? partPath : partPath.slice(slash + 1);
  return `${directory}_rels/${fileName}.rels`;
}

function relsBasePart(relsPath: string): string {
  // _rels/.rels -> ''
  // xl/_rels/workbook.xml.rels -> xl/workbook.xml
  const marker = '_rels/';
  const markerIndex = relsPath.lastIndexOf(marker);
  if (markerIndex < 0) return relsPath;
  const directory = relsPath.slice(0, markerIndex);
  const file = relsPath.slice(markerIndex + marker.length).replace(/\.rels$/, '');
  return `${directory}${file}`;
}
