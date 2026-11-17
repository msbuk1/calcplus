import { arrayFormulaForCell } from '@model/formula-ranges';
import type { Cell, Sheet, SheetId, Workbook } from '@model/types';
import {
  cellFormulaKey,
  cellKey,
  cellStateEquals,
  formatCellRef,
  getCell,
  getSheetById,
  parseCellKey,
} from '@model/types';
import type { PatchMode, PackagePatchReport, Relationship, WorkbookPackage } from './types';
import type {
  DirtyCellInput,
  DirtyDimensions,
  WorksheetSaveOptions,
} from './patch/worksheet-patch';
import { patchWorksheetXml } from './patch/worksheet-patch';
import {
  createSharedStringsAppender,
  type SharedStringsAppender,
} from './patch/shared-strings-patch';
import { adoptAppendedStyleIds, appendStylesXml } from './patch/styles-patch';
import { createWorksheetBytes } from './patch/new-worksheet';
import {
  parseWorkbookSheets,
  patchContentTypesForSheets,
  patchWorkbookRelationshipsForSheets,
  patchWorkbookSheetList,
  planWorkbookSheetChanges,
  removeContentTypeOverrides,
  type WorkbookSheetPlan,
} from './patch/workbook-sheets-patch';
import { parseRelationships } from './relationships';
import { parseSharedStringsXml } from './shared-strings';
import { parseWorksheetXml } from './worksheet-xml';
import { parseStylesXml } from './styles-xml';
import { parseWorkbookXml } from './workbook-xml';
import { WorkbookSaveError } from './errors';
import { validatePatchedPart } from './xml';
import { writePackage } from './zip';
import {
  patchWorkbookDefinedNames,
  removeCalcChain,
  setFullCalcOnLoad,
  setWorkbookDate1904,
} from './patch/workbook-patch';

export interface SaveWorkbookInput {
  readonly workbook: Workbook;
  readonly package: WorkbookPackage;
  /** Sheet id -> set of `row:col` keys changed since the last save. */
  readonly dirtyCells: ReadonlyMap<SheetId, ReadonlySet<string>>;
  /** Sheet id -> row and column records changed since the last save. */
  readonly dirtyDimensions?: ReadonlyMap<SheetId, DirtyDimensions>;
  /** Sheets whose merged ranges changed since the last save. */
  readonly dirtyMerges?: ReadonlySet<SheetId>;
  readonly options?: {
    readonly worksheet?: WorksheetSaveOptions;
  };
}

export interface SaveWorkbookResult {
  readonly ok: true;
  readonly bytes: Uint8Array;
  readonly noOp: boolean;
  readonly package: WorkbookPackage;
  readonly report: PackagePatchReport;
}

export function saveWorkbook(input: SaveWorkbookInput): SaveWorkbookResult {
  let entries = new Map(input.package.entries);
  const workbookPath = input.package.workbookPartPath;
  const stylesPath = input.package.stylesPath;
  const sourceStylesBytes = stylesPath === null ? undefined : entries.get(stylesPath);
  const sourceStyles =
    stylesPath === null || sourceStylesBytes === undefined
      ? null
      : parseStylesXml(sourceStylesBytes, stylesPath);
  const stylesNeedAppend =
    sourceStyles !== null &&
    input.workbook.styles.some((style) => style.id >= sourceStyles.sourceStyleCount);
  const sourceWorkbookBytes = entries.get(workbookPath);
  const dateSystemChanged =
    sourceWorkbookBytes !== undefined &&
    parseWorkbookXml(sourceWorkbookBytes, workbookPath).date1904 !== input.workbook.date1904;

  const packageSheetRelationshipIds =
    input.package.sheetRelationshipIds ?? new Map<SheetId, string>();
  const workbookRelsBytes = entries.get(input.package.workbookRelsPath);
  const workbookRelationships =
    input.package.relationships.get(input.package.workbookRelsPath) ??
    (workbookRelsBytes === undefined
      ? []
      : parseRelationships(workbookRelsBytes, input.package.workbookRelsPath, workbookPath));
  const parsedSheets =
    sourceWorkbookBytes === undefined
      ? null
      : parseWorkbookSheets(sourceWorkbookBytes, workbookPath);
  const sheetPlan: WorkbookSheetPlan | null =
    parsedSheets === null
      ? null
      : planWorkbookSheetChanges({
          workbook: input.workbook,
          parsed: parsedSheets,
          workbookRelationships,
          workbookPartPath: workbookPath,
          workbookRelsPath: input.package.workbookRelsPath,
          sheetPartPaths: input.package.sheetPartPaths,
          sheetRelationshipIds: packageSheetRelationshipIds,
        });
  // A new sheet is written whole and skips the patch path. A sheet that was
  // deleted earlier can never come back: the delete is terminal.
  const newSheetIds = new Set((sheetPlan?.additions ?? []).map((addition) => addition.sheetId));
  const sourceSheetIds = sourceSheetIdsByIndex(
    packageSheetRelationshipIds,
    input.workbook,
    parsedSheets,
  );
  const removedSheetIds = new Set(sheetPlan?.removals.map((removal) => removal.sheetId) ?? []);
  const definedNamesChanged =
    sourceWorkbookBytes !== undefined &&
    patchWorkbookDefinedNames({
      bytes: sourceWorkbookBytes,
      partName: workbookPath,
      workbook: input.workbook,
      sourceSheetIds,
      removedSheetIds,
    }) !== null;

  // A snapshot restored after a rename or delete carries the rewritten model
  // formulas but not the undo history's dirty markers. On a structural save,
  // compare every authored formula with its source part so the patch still
  // writes them.
  const dirtyCellsForSave = mergeDirtyCells(
    input.dirtyCells,
    sheetPlan?.changed === true ? structuralFormulaDirtyCells(input, newSheetIds) : new Map(),
  );
  const dirtyCellSheetIds = [...dirtyCellsForSave.entries()]
    .filter(([, cells]) => cells.size > 0)
    .map(([sheetId]) => sheetId);
  const dirtyDimensionSheetIds = [...(input.dirtyDimensions?.entries() ?? [])]
    .filter(([, dimensions]) => dimensions.rows.size > 0 || dimensions.columns.size > 0)
    .map(([sheetId]) => sheetId);
  const dirtyMergeSheetIds = [...(input.dirtyMerges ?? [])].filter((sheetId) =>
    input.workbook.sheets.some((sheet) => sheet.id === sheetId),
  );
  const dirtySheetIds = [
    ...new Set([...dirtyCellSheetIds, ...dirtyDimensionSheetIds, ...dirtyMergeSheetIds]),
  ].filter(
    (sheetId) =>
      !newSheetIds.has(sheetId) && input.workbook.sheets.some((sheet) => sheet.id === sheetId),
  );

  if (
    dirtySheetIds.length === 0 &&
    !dateSystemChanged &&
    !stylesNeedAppend &&
    sheetPlan?.changed !== true &&
    !definedNamesChanged
  ) {
    const bytes = input.package.originalBytes ?? writePackage(input.package.entries);
    return {
      ok: true,
      bytes,
      noOp: true,
      package: { ...input.package, originalBytes: bytes },
      report: {
        patchedParts: [],
        modeByPart: new Map<string, PatchMode>(),
        appendedSharedStrings: 0,
        removedParts: [],
      },
    };
  }

  assertStylesAreWritable(input, sourceStyles, dirtyCellsForSave);
  let relationships = new Map(input.package.relationships);
  const sharedStringsPath = input.package.sharedStringsPath;
  const sharedStringsBytes =
    sharedStringsPath === null ? undefined : entries.get(sharedStringsPath);
  let appender: SharedStringsAppender | null = null;
  if (sharedStringsPath !== null && sharedStringsBytes !== undefined) {
    const parsed = parseSharedStringsXml(sharedStringsBytes, sharedStringsPath);
    appender = createSharedStringsAppender(parsed.structure, sharedStringsPath);
  }

  const patchedParts: string[] = [];
  const modeByPart = new Map<string, PatchMode>();
  const removedParts: string[] = [];
  let formulaChanged = false;

  if (stylesNeedAppend) {
    if (stylesPath === null || sourceStylesBytes === undefined) {
      throw new WorkbookSaveError(
        'unsupported-edit',
        'The workbook has no styles part to receive the new formatting. The file was not written.',
      );
    }
    const appended = appendStylesXml({
      partName: stylesPath,
      sourceBytes: sourceStylesBytes,
      workbook: input.workbook,
    });
    entries.set(stylesPath, appended.bytes);
    adoptAppendedStyleIds(input.workbook, appended.appendedStyles);
    patchedParts.push(stylesPath);
    modeByPart.set(stylesPath, 'position');
  }

  for (const sheetId of dirtySheetIds) {
    const sheet = getSheetById(input.workbook, sheetId);
    const dirtyKeys = dirtyCellsForSave.get(sheetId) ?? new Set<string>();
    const dirtyCells: DirtyCellInput[] = [...dirtyKeys]
      .map((key) => {
        const address = parseCellKey(key);
        const cell = getCell(sheet, address.row, address.col);
        return { row: address.row, col: address.col, cell };
      })
      .sort((a, b) => a.row - b.row || a.col - b.col);
    const dirtyDimensions = input.dirtyDimensions?.get(sheetId);
    const syncMerges = input.dirtyMerges?.has(sheetId) === true;

    const sourceBytes = input.workbook.sourceParts.get(sheet.partPath);
    if (sourceBytes === undefined) {
      throw new WorkbookSaveError(
        'unsupported-edit',
        `The worksheet ${sheet.partPath} has no source skeleton to patch.`,
      );
    }

    for (const dirty of dirtyCells) {
      const arrayMaster = arrayFormulaForCell(sheet, dirty.row, dirty.col);
      if (arrayMaster !== null && arrayMaster !== dirty.cell) {
        throw new WorkbookSaveError(
          'unsupported-edit',
          'Legacy array formulas are read-only in this milestone; the file was not written.',
        );
      }
    }

    if (dirtyCells.length > 0) {
      formulaChanged =
        dirtyCellsChangeFormula(sourceBytes, sheet, dirtyCells, input) || formulaChanged;
    }

    const result = patchWorksheetXml({
      partName: sheet.partPath,
      sourceBytes,
      sheet,
      dirtyCells,
      dirtyDimensions,
      syncMerges,
      sharedStringsTable: input.package.sharedStrings,
      sharedStrings: appender,
      options: input.options?.worksheet,
    });
    entries.set(sheet.partPath, result.bytes);
    if (result.mode !== 'none') {
      patchedParts.push(sheet.partPath);
      modeByPart.set(sheet.partPath, result.mode);
    }
  }

  // Sheet insert, delete and rename are structural: new parts are written in
  // full, removed parts disappear with their package references, and the
  // workbook part is patched in place so untouched sheet elements, hidden
  // state and tab colours survive.
  if (sheetPlan !== null && sheetPlan.changed) {
    // calcChain stores sheet indices and cell addresses; any sheet-list change
    // makes it untrustworthy, even when no formula text changed.
    formulaChanged = true;
    const contentTypesPath = input.package.contentTypesPath;
    const contentTypesBefore = entries.get(contentTypesPath);
    const relationshipGraph = buildRelationshipGraph(input);
    let contentTypesDirty = false;
    const removedSheetPartPaths: string[] = [];

    for (const removal of sheetPlan.removals) {
      if (removal.partPath === null) continue;
      const paths = collectRemovedSheetParts(removal.partPath, workbookPath, relationshipGraph);
      const pathList = [...paths].sort();
      removedSheetPartPaths.push(...pathList);
      for (const path of pathList) {
        entries.delete(path);
        // The delete is terminal, so the removed bytes can never be needed by
        // an undo. Drop them from the model too instead of leaking a sheet-
        // sized skeleton into the next snapshot.
        input.workbook.sourceParts.delete(path);
        input.workbook.opaqueParts.delete(path);
        removedParts.push(path);
      }
      for (const path of pathList) {
        if (path.endsWith('.rels')) relationships.delete(path);
      }
      // Deleting a sheet removes its formula cells, so calcChain is stale.
      formulaChanged = true;
    }

    if (contentTypesBefore !== undefined && removedSheetPartPaths.length > 0) {
      const removed = removeContentTypeOverrides(
        contentTypesBefore,
        contentTypesPath,
        removedSheetPartPaths,
      );
      if (removed !== null) {
        entries.set(contentTypesPath, removed);
        contentTypesDirty = true;
      }
    }

    for (const addition of sheetPlan.additions) {
      const sheet = getSheetById(input.workbook, addition.sheetId);
      if (entries.has(sheet.partPath)) {
        throw new WorkbookSaveError(
          'unsupported-edit',
          `The worksheet part ${sheet.partPath} already exists. The file was not written.`,
        );
      }
      entries.set(sheet.partPath, createWorksheetBytes(sheet, appender));
      patchedParts.push(sheet.partPath);
      modeByPart.set(sheet.partPath, 'position');
    }

    entries.set(workbookPath, patchWorkbookSheetList(parsedSheets!, sheetPlan, workbookPath));
    if (!patchedParts.includes(workbookPath)) patchedParts.push(workbookPath);
    modeByPart.set(workbookPath, 'position');

    const relsBytes = entries.get(input.package.workbookRelsPath);
    const patchedRels =
      relsBytes === undefined
        ? null
        : patchWorkbookRelationshipsForSheets(
            relsBytes,
            input.package.workbookRelsPath,
            sheetPlan,
            workbookRelationships,
          );
    if (patchedRels !== null) {
      entries.set(input.package.workbookRelsPath, patchedRels);
      relationships = new Map(relationships).set(
        input.package.workbookRelsPath,
        parseRelationships(patchedRels, input.package.workbookRelsPath, workbookPath),
      );
      if (!patchedParts.includes(input.package.workbookRelsPath)) {
        patchedParts.push(input.package.workbookRelsPath);
      }
      modeByPart.set(input.package.workbookRelsPath, 'position');
    }

    const contentTypesBytes = entries.get(contentTypesPath);
    const patchedContentTypes =
      contentTypesBytes === undefined
        ? null
        : patchContentTypesForSheets(contentTypesBytes, contentTypesPath, sheetPlan);
    if (patchedContentTypes !== null) {
      entries.set(contentTypesPath, patchedContentTypes);
      contentTypesDirty = true;
    }
    if (contentTypesDirty) {
      if (!patchedParts.includes(contentTypesPath)) patchedParts.push(contentTypesPath);
      modeByPart.set(contentTypesPath, 'position');
    }
  }

  // Defined-name formulas follow a sheet rename or delete in the same
  // transaction. Patch the current workbook bytes so the sheet-list patch has
  // already settled; only the changed name text is replaced.
  {
    const workbookBytes = entries.get(workbookPath);
    if (workbookBytes !== undefined) {
      const patchedNames = patchWorkbookDefinedNames({
        bytes: workbookBytes,
        partName: workbookPath,
        workbook: input.workbook,
        sourceSheetIds,
        removedSheetIds,
      });
      if (patchedNames !== null) {
        entries.set(workbookPath, patchedNames);
        if (!patchedParts.includes(workbookPath)) patchedParts.push(workbookPath);
        modeByPart.set(workbookPath, 'position');
      }
    }
  }

  let sharedStringsTable = input.package.sharedStrings;
  let appendedSharedStrings = 0;
  if (appender !== null && appender.changed && sharedStringsPath !== null) {
    appendedSharedStrings = appender.addedCount;
    entries.set(sharedStringsPath, appender.bytes());
    sharedStringsTable = appender.table();
    patchedParts.push(sharedStringsPath);
    modeByPart.set(sharedStringsPath, 'position');
  }

  // The workbook date-system flag is document state; patch only the one
  // attribute and keep every other byte of the workbook part.
  if (dateSystemChanged) {
    const workbookBytes = entries.get(workbookPath);
    if (workbookBytes !== undefined) {
      const patchedWorkbook = setWorkbookDate1904(
        workbookBytes,
        workbookPath,
        input.workbook.date1904,
      );
      if (!bytesEqual(patchedWorkbook, workbookBytes)) {
        entries.set(workbookPath, patchedWorkbook);
        if (!patchedParts.includes(workbookPath)) patchedParts.push(workbookPath);
        modeByPart.set(workbookPath, 'position');
      }
    }
  }

  // File-compat §4/§6.1: an input edit can stale dependent formulas, so ask
  // Excel to recalculate while leaving their cached <v> results untouched.
  const hasFormulas = workbookHasFormulas(input.workbook);
  if (hasFormulas || formulaChanged) {
    const workbookBytes = entries.get(workbookPath);
    if (workbookBytes !== undefined) {
      const patchedWorkbook = setFullCalcOnLoad(workbookBytes, workbookPath);
      if (!bytesEqual(patchedWorkbook, workbookBytes)) {
        entries.set(workbookPath, patchedWorkbook);
        if (!patchedParts.includes(workbookPath)) patchedParts.push(workbookPath);
        modeByPart.set(workbookPath, 'position');
      }
    }
  }

  if (formulaChanged) {
    const calc = removeCalcChain({
      entries,
      relationships,
      contentTypesPath: input.package.contentTypesPath,
      workbookRelsPath: input.package.workbookRelsPath,
    });
    entries = calc.entries;
    relationships = calc.relationships;
    removedParts.push(...calc.removedParts);
    for (const path of calc.patchedParts) {
      if (!patchedParts.includes(path)) patchedParts.push(path);
      modeByPart.set(path, 'position');
    }
  }

  for (const path of new Set(patchedParts)) {
    const bytes = entries.get(path);
    if (bytes !== undefined) validatePatchedPart(bytes, path);
  }

  const sheetPartPaths = new Map<SheetId, string>();
  const sheetRelationshipIds = new Map<SheetId, string>();
  for (const sheet of input.workbook.sheets) {
    const addition = sheetPlan?.additions.find((entry) => entry.sheetId === sheet.id);
    sheetPartPaths.set(
      sheet.id,
      addition?.partPath ?? input.package.sheetPartPaths.get(sheet.id) ?? sheet.partPath,
    );
    const relationshipId = addition?.relationshipId ?? packageSheetRelationshipIds.get(sheet.id);
    if (relationshipId !== undefined) {
      sheetRelationshipIds.set(sheet.id, relationshipId);
    }
  }

  const bytes = writePackage(entries);
  const newPackage: WorkbookPackage = {
    ...input.package,
    originalBytes: bytes,
    entries,
    relationships,
    sheetPartPaths,
    sheetRelationshipIds,
    sharedStrings: sharedStringsTable,
  };

  return {
    ok: true,
    bytes,
    noOp: false,
    package: newPackage,
    report: {
      patchedParts,
      modeByPart,
      appendedSharedStrings,
      removedParts,
    },
  };
}

/**
 * Build the package relationship graph for a delete. The package map is used
 * when present; a restored snapshot carries no parsed map, so every `.rels`
 * part is parsed on demand.
 */
function buildRelationshipGraph(input: SaveWorkbookInput): Map<string, Relationship[]> {
  const graph = new Map<string, Relationship[]>();
  for (const [path, list] of input.package.relationships) graph.set(path, [...list]);
  for (const [path, bytes] of input.package.entries) {
    if (!path.endsWith('.rels') || graph.has(path)) continue;
    graph.set(path, parseRelationships(bytes, path, basePartForRelsPath(path)));
  }
  return graph;
}

/**
 * The worksheet part, its relationship part and every part reachable only
 * through them. A part still referenced by a surviving sheet is left in the
 * package, so shared drawings and media are never removed from another sheet.
 */
function collectRemovedSheetParts(
  worksheetPartPath: string,
  workbookPartPath: string,
  graph: ReadonlyMap<string, readonly Relationship[]>,
): Set<string> {
  const candidates = new Set<string>();
  const queue = [worksheetPartPath];
  while (queue.length > 0) {
    const part = queue.shift()!;
    if (candidates.has(part)) continue;
    candidates.add(part);
    const relsPath = relsPathForPart(part);
    const relationships = graph.get(relsPath);
    if (relationships === undefined) continue;
    candidates.add(relsPath);
    for (const relationship of relationships) {
      if (relationship.targetMode !== 'Internal') continue;
      if (!candidates.has(relationship.resolvedTarget)) {
        queue.push(relationship.resolvedTarget);
      }
    }
  }

  const reachable = reachableFromPackageRoot(workbookPartPath, worksheetPartPath, graph);
  const removed = new Set<string>([worksheetPartPath, relsPathForPart(worksheetPartPath)]);
  for (const path of candidates) {
    if (path === worksheetPartPath) continue;
    if (path.endsWith('.rels')) {
      if (removed.has(basePartForRelsPath(path))) removed.add(path);
      continue;
    }
    if (!reachable.has(path)) removed.add(path);
  }
  // A removed dependent part takes its own relationship part with it.
  let changed = true;
  while (changed) {
    changed = false;
    for (const path of [...removed]) {
      const relsPath = relsPathForPart(path);
      if (candidates.has(relsPath) && !removed.has(relsPath)) {
        removed.add(relsPath);
        changed = true;
      }
    }
  }
  return removed;
}

function reachableFromPackageRoot(
  workbookPartPath: string,
  removedWorksheetPartPath: string,
  graph: ReadonlyMap<string, readonly Relationship[]>,
): Set<string> {
  const reachable = new Set<string>();
  const visited = new Set<string>();
  const queue = [''];

  while (queue.length > 0) {
    const source = queue.shift()!;
    if (visited.has(source)) continue;
    visited.add(source);
    const relationships = graph.get(relsPathForPart(source)) ?? [];
    for (const relationship of relationships) {
      if (relationship.targetMode !== 'Internal') continue;
      const target = relationship.resolvedTarget;
      if (target === removedWorksheetPartPath) continue;
      if (reachable.has(target)) continue;
      reachable.add(target);
      queue.push(target);
    }
  }
  // The workbook itself is a target of the package root, but an empty or
  // unusual graph must not make the workbook look like a dependent part.
  reachable.add(workbookPartPath);
  return reachable;
}

function basePartForRelsPath(relsPath: string): string {
  const marker = '_rels/';
  const markerIndex = relsPath.lastIndexOf(marker);
  if (markerIndex < 0) return relsPath;
  const directory = relsPath.slice(0, markerIndex);
  const file = relsPath.slice(markerIndex + marker.length).replace(/\.rels$/, '');
  return `${directory}${file}`;
}

/**
 * A dirty cell may only reference a style record the model actually holds.
 * The styles writer appends every record above the source `cellXfs` count, so
 * a missing id would be written as a dangling reference and must be refused.
 */
function assertStylesAreWritable(
  input: SaveWorkbookInput,
  sourceStyles: ReturnType<typeof parseStylesXml> | null,
  dirtyCells: ReadonlyMap<SheetId, ReadonlySet<string>>,
): void {
  const styleCount = input.workbook.styles.length;
  for (const [sheetId, keys] of dirtyCells) {
    if (keys.size === 0) continue;
    const sheet = input.workbook.sheets.find((candidate) => candidate.id === sheetId);
    if (sheet === undefined) continue;
    for (const key of keys) {
      const address = parseCellKey(key);
      const cell = getCell(sheet, address.row, address.col);
      if (cell === null || cell.styleId === 0) continue;
      if (cell.styleId >= styleCount) {
        throw new WorkbookSaveError(
          'unsupported-edit',
          `The formatting at ${formatCellRef(address.row, address.col)} references a style that this workbook does not hold. The file was not written.`,
        );
      }
    }
  }
  if (
    sourceStyles === null &&
    input.workbook.styles.some((style) => style.id > 0) &&
    input.workbook.styles.length > 1
  ) {
    throw new WorkbookSaveError(
      'unsupported-edit',
      'The workbook has no readable style table to receive the new formatting. The file was not written.',
    );
  }
}

function mergeDirtyCells(
  base: ReadonlyMap<SheetId, ReadonlySet<string>>,
  extra: ReadonlyMap<SheetId, ReadonlySet<string>>,
): Map<SheetId, Set<string>> {
  const merged = new Map<SheetId, Set<string>>();
  for (const [sheetId, keys] of base) merged.set(sheetId, new Set(keys));
  for (const [sheetId, keys] of extra) {
    const set = merged.get(sheetId) ?? new Set<string>();
    for (const key of keys) set.add(key);
    merged.set(sheetId, set);
  }
  return merged;
}

/**
 * On a structural save, mark every formula whose authored text differs from
 * the source part. This is what makes rename/delete formulas survive a save
 * after a snapshot restore, where dirty markers are not persisted.
 */
function structuralFormulaDirtyCells(
  input: SaveWorkbookInput,
  newSheetIds: ReadonlySet<SheetId>,
): Map<SheetId, Set<string>> {
  const dirty = new Map<SheetId, Set<string>>();
  for (const sheet of input.workbook.sheets) {
    if (newSheetIds.has(sheet.id)) continue;
    const sourceBytes = input.workbook.sourceParts.get(sheet.partPath);
    if (sourceBytes === undefined) continue;
    let hasFormula = false;
    for (const cell of sheet.cells.values()) {
      if (cell.formula !== null) {
        hasFormula = true;
        break;
      }
    }
    if (!hasFormula) continue;
    const parsed = parseWorksheetXml(sourceBytes, {
      sharedStrings: input.package.sharedStrings,
      partPath: sheet.partPath,
      name: sheet.name,
    });
    for (const [key, cell] of sheet.cells) {
      if (cell.formula === null) continue;
      const original = parsed.sheet.cells.get(key) ?? null;
      if (cellFormulaKey(original) === cellFormulaKey(cell)) continue;
      const keys = dirty.get(sheet.id) ?? new Set<string>();
      keys.add(key);
      dirty.set(sheet.id, keys);
    }
  }
  return dirty;
}

function dirtyCellsChangeFormula(
  sourceBytes: Uint8Array,
  sheet: Sheet,
  dirtyCells: readonly DirtyCellInput[],
  input: SaveWorkbookInput,
): boolean {
  const parsed = parseWorksheetXml(sourceBytes, {
    sharedStrings: input.package.sharedStrings,
    partPath: sheet.partPath,
    name: sheet.name,
  });
  let formulaChanged = false;
  for (const dirty of dirtyCells) {
    const original = parsed.sheet.cells.get(cellKey(dirty.row, dirty.col)) ?? null;
    // File-compat §6.1: shared and legacy array groups are read-only in M1.
    // The command rejects content edits first; this save-level guard stops a
    // direct model edit from rewriting a group (including its t="str" cached
    // value). A style-only change keeps the group intact and is allowed.
    if (original?.formulaSource !== undefined && !sameContentIgnoringStyle(original, dirty.cell)) {
      throw new WorkbookSaveError(
        'unsupported-edit',
        original.formulaSource.kind === 'array'
          ? 'Legacy array formulas are read-only in this milestone; the file was not written.'
          : 'Shared formulas are read-only in this milestone; the file was not written.',
      );
    }
    // File-compat F17 / workbook-model §9: imported rich text is preserved
    // while untouched. The model cannot write rich runs back, so a dirty cell
    // that still holds the original rich text would be flattened silently.
    // The M1 editor blocks these edits first; this is the save-level guard.
    if (original?.richText !== null && original?.richText !== undefined) {
      throw new WorkbookSaveError(
        'unsupported-edit',
        `The rich text at ${formatCellRef(dirty.row, dirty.col)} cannot be rewritten in this milestone; the file was not written.`,
      );
    }
    if (cellFormulaKey(original) !== cellFormulaKey(dirty.cell)) formulaChanged = true;
  }
  return formulaChanged;
}

function sameContentIgnoringStyle(original: Cell | null, next: Cell | null): boolean {
  if (original === null || next === null) return original === next;
  return cellStateEquals(original, { ...next, styleId: original.styleId });
}

function workbookHasFormulas(workbook: Workbook): boolean {
  for (const sheet of workbook.sheets) {
    for (const cell of sheet.cells.values()) {
      if (cell.formula !== null || cell.formulaSource?.kind === 'shared-follower') return true;
    }
  }
  return false;
}

function bytesEqual(left: Uint8Array, right: Uint8Array): boolean {
  if (left.byteLength !== right.byteLength) return false;
  for (let index = 0; index < left.byteLength; index += 1) {
    if (left[index] !== right[index]) return false;
  }
  return true;
}

/**
 * Map each source `<sheet>` index to the stable sheet id it carried at open.
 * The fallback by order covers a snapshot written before sheet relationship
 * ids were stored; such a snapshot can only describe an unchanged sheet list.
 */
function sourceSheetIdsByIndex(
  sheetRelationshipIds: ReadonlyMap<SheetId, string>,
  workbook: Workbook,
  parsedSheets: ReturnType<typeof parseWorkbookSheets> | null,
): (SheetId | null)[] {
  if (parsedSheets === null) return [];
  const sheetIdByRelationshipId = new Map<string, SheetId>();
  for (const [sheetId, relationshipId] of sheetRelationshipIds) {
    sheetIdByRelationshipId.set(relationshipId, sheetId);
  }
  if (sheetIdByRelationshipId.size === 0) {
    return parsedSheets.elements.map((_, index) => workbook.sheets[index]?.id ?? null);
  }
  return parsedSheets.elements.map(
    (element) => sheetIdByRelationshipId.get(element.relationshipId) ?? null,
  );
}

function relsPathForPart(partPath: string): string {
  const slash = partPath.lastIndexOf('/');
  const directory = slash < 0 ? '' : partPath.slice(0, slash + 1);
  const fileName = slash < 0 ? partPath : partPath.slice(slash + 1);
  return `${directory}_rels/${fileName}.rels`;
}

export function adoptSavedPackage(workbook: Workbook, pkg: WorkbookPackage): void {
  const sourcePaths = new Set<string>([
    pkg.workbookPartPath,
    ...workbook.sheets.map((sheet) => sheet.partPath),
  ]);
  if (pkg.stylesPath !== null) sourcePaths.add(pkg.stylesPath);
  if (pkg.sharedStringsPath !== null) sourcePaths.add(pkg.sharedStringsPath);
  if (pkg.themePath !== null) sourcePaths.add(pkg.themePath);

  // A removed sheet's part is no longer a source part, and a newly inserted
  // sheet's part becomes one so later edits can patch it in place.
  for (const path of [...workbook.sourceParts.keys()]) {
    if (!sourcePaths.has(path) || !pkg.entries.has(path)) workbook.sourceParts.delete(path);
  }
  for (const path of sourcePaths) {
    const bytes = pkg.entries.get(path);
    if (bytes !== undefined) workbook.sourceParts.set(path, bytes);
  }

  workbook.opaqueParts.clear();
  for (const [path, bytes] of pkg.entries) {
    if (!workbook.sourceParts.has(path)) {
      workbook.opaqueParts.set(path, { path, bytes, anchor: null });
    }
  }
}

export function dirtyCellsFromChanges(
  changes: readonly { sheetId: SheetId; cells: readonly { row: number; col: number }[] }[],
): Map<SheetId, Set<string>> {
  const dirty = new Map<SheetId, Set<string>>();
  for (const change of changes) {
    const set = dirty.get(change.sheetId) ?? new Set<string>();
    for (const cell of change.cells) set.add(cellKey(cell.row, cell.col));
    dirty.set(change.sheetId, set);
  }
  return dirty;
}

export function cellMapForSheet(workbook: Workbook, sheetId: SheetId): Map<string, Cell> {
  return getSheetById(workbook, sheetId).cells;
}
