import { arrayFormulaForCell } from '@model/formula-ranges';
import type { Cell, Sheet, SheetId, Workbook } from '@model/types';
import {
  cellFormulaKey,
  cellKey,
  formatCellRef,
  getCell,
  getSheetById,
  parseCellKey,
} from '@model/types';
import type { PatchMode, PackagePatchReport, WorkbookPackage } from './types';
import type { DirtyCellInput, WorksheetSaveOptions } from './patch/worksheet-patch';
import { patchWorksheetXml } from './patch/worksheet-patch';
import {
  createSharedStringsAppender,
  type SharedStringsAppender,
} from './patch/shared-strings-patch';
import { parseSharedStringsXml } from './shared-strings';
import { parseWorksheetXml } from './worksheet-xml';
import { WorkbookSaveError } from './errors';
import { validatePatchedPart } from './xml';
import { writePackage } from './zip';
import { removeCalcChain, setFullCalcOnLoad } from './patch/workbook-patch';

export interface SaveWorkbookInput {
  readonly workbook: Workbook;
  readonly package: WorkbookPackage;
  /** Sheet id -> set of `row:col` keys changed since the last save. */
  readonly dirtyCells: ReadonlyMap<SheetId, ReadonlySet<string>>;
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
  const dirtySheetIds = [...input.dirtyCells.entries()]
    .filter(([, cells]) => cells.size > 0)
    .map(([sheetId]) => sheetId);

  if (dirtySheetIds.length === 0) {
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

  let entries = new Map(input.package.entries);
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

  for (const sheetId of dirtySheetIds) {
    const sheet = getSheetById(input.workbook, sheetId);
    const dirtyKeys = input.dirtyCells.get(sheetId) ?? new Set<string>();
    const dirtyCells: DirtyCellInput[] = [...dirtyKeys]
      .map((key) => {
        const address = parseCellKey(key);
        const cell = getCell(sheet, address.row, address.col);
        return { row: address.row, col: address.col, cell };
      })
      .sort((a, b) => a.row - b.row || a.col - b.col);

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

    formulaChanged =
      dirtyCellsChangeFormula(sourceBytes, sheet, dirtyCells, input) || formulaChanged;

    const result = patchWorksheetXml({
      partName: sheet.partPath,
      sourceBytes,
      sheet,
      dirtyCells,
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

  let sharedStringsTable = input.package.sharedStrings;
  let appendedSharedStrings = 0;
  if (appender !== null && appender.changed && sharedStringsPath !== null) {
    appendedSharedStrings = appender.addedCount;
    entries.set(sharedStringsPath, appender.bytes());
    sharedStringsTable = appender.table();
    patchedParts.push(sharedStringsPath);
    modeByPart.set(sharedStringsPath, 'position');
  }

  // File-compat §4/§6.1: an input edit can stale dependent formulas, so ask
  // Excel to recalculate while leaving their cached <v> results untouched.
  const hasFormulas = workbookHasFormulas(input.workbook);
  if (hasFormulas || formulaChanged) {
    const workbookPath = input.package.workbookPartPath;
    const workbookBytes = entries.get(workbookPath);
    if (workbookBytes !== undefined) {
      const patchedWorkbook = setFullCalcOnLoad(workbookBytes, workbookPath);
      if (!bytesEqual(patchedWorkbook, workbookBytes)) {
        entries.set(workbookPath, patchedWorkbook);
        patchedParts.push(workbookPath);
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

  const bytes = writePackage(entries);
  const newPackage: WorkbookPackage = {
    ...input.package,
    originalBytes: bytes,
    entries,
    relationships,
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
    // The command rejects them first; this save-level guard stops a direct
    // model edit from rewriting a group (including its t="str" cached value).
    if (original?.formulaSource !== undefined) {
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

export function adoptSavedPackage(workbook: Workbook, pkg: WorkbookPackage): void {
  for (const path of workbook.sourceParts.keys()) {
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
