import {
  planRangeTransform,
  type RangeTransformPlan,
  type StructuralEdit,
} from '@model/range-transform';
import type { GridRange, RowDimension, Sheet, SheetId } from '@model/types';
import { MAX_COLUMN, MAX_ROW } from '@model/references';
import {
  planOpaquePartPatches,
  renumberWorksheetPart,
  type StructuralPartPatch,
} from '@ooxml/structure-patch';
import type {
  CommandChange,
  CommandContext,
  CommandDefinition,
  CommandResult,
  StructureApplyChange,
  StructurePayload,
  StructureRowsChange,
} from './types';

/**
 * The six structural-edit commands (feature matrix "Insert/delete rows,
 * columns and cells"; command-system §2 and §4).
 *
 * Each command resolves its target from `args.range ?? context.selection ??
 * the active cell` (ticket 29), plans every range-bearing consequence through
 * the single range-transform module, and returns **one** transaction: one
 * `structure.apply` change plus one `name.set` per rewritten defined name.
 * `structure.apply` carries the moved, deleted and formula-rewritten cells, the
 * renumbered dimensions, merges, comments, hyperlinks, CF and validation
 * records, the AutoFilter and the renumbered source-part baseline, so undo
 * restores the exact prior state without recomputing a transform.
 *
 * Insert Cells & Shift Right / Down and Delete Cells & Shift Up / Left are
 * banded edits: the shift applies to the selected rows (for a column shift) or
 * columns (for a row shift). A banded edit that would split a merged cell or
 * the AutoFilter rectangle is blocked with Excel's message, exactly as Excel
 * refuses to change part of a merged cell.
 */

const INSERT_ROWS_KEYBINDING = 'Ctrl+Shift+=';
const DELETE_KEYBINDING = 'Ctrl+-';

export type CellShift = 'right' | 'down' | 'up' | 'left';

export interface StructureArgs {
  readonly sheetId?: SheetId;
  /** The target; defaults to the selection, then the active cell (ticket 29). */
  readonly range?: GridRange;
  /**
   * The Insert/Delete Cells direction. Insert defaults to Shift Down and
   * Delete to Shift Up, Excel's defaults in the Insert/Delete dialog.
   */
  readonly shift?: CellShift;
}

type EditKind =
  | 'insert-rows'
  | 'insert-columns'
  | 'insert-cells'
  | 'delete-rows'
  | 'delete-columns'
  | 'delete-cells';

export const structureInsertRowsCommand: CommandDefinition<StructureArgs> = {
  id: 'structure.insertRows',
  label: 'Insert rows',
  category: 'edit',
  keybindings: [INSERT_ROWS_KEYBINDING],
  isEnabled: hasSheet,
  run(context, args = {}) {
    return runStructure(context, args, 'insert-rows');
  },
};

export const structureInsertColumnsCommand: CommandDefinition<StructureArgs> = {
  id: 'structure.insertColumns',
  label: 'Insert columns',
  category: 'edit',
  keybindings: [INSERT_ROWS_KEYBINDING],
  isEnabled: hasSheet,
  run(context, args = {}) {
    return runStructure(context, args, 'insert-columns');
  },
};

export const structureInsertCellsCommand: CommandDefinition<StructureArgs> = {
  id: 'structure.insertCells',
  label: 'Insert cells',
  category: 'edit',
  keybindings: [INSERT_ROWS_KEYBINDING],
  isEnabled: hasSheet,
  run(context, args = {}) {
    return runStructure(context, args, 'insert-cells');
  },
};

export const structureDeleteRowsCommand: CommandDefinition<StructureArgs> = {
  id: 'structure.deleteRows',
  label: 'Delete rows',
  category: 'edit',
  keybindings: [DELETE_KEYBINDING],
  isEnabled: hasSheet,
  run(context, args = {}) {
    return runStructure(context, args, 'delete-rows');
  },
};

export const structureDeleteColumnsCommand: CommandDefinition<StructureArgs> = {
  id: 'structure.deleteColumns',
  label: 'Delete columns',
  category: 'edit',
  keybindings: [DELETE_KEYBINDING],
  isEnabled: hasSheet,
  run(context, args = {}) {
    return runStructure(context, args, 'delete-columns');
  },
};

export const structureDeleteCellsCommand: CommandDefinition<StructureArgs> = {
  id: 'structure.deleteCells',
  label: 'Delete cells',
  category: 'edit',
  keybindings: [DELETE_KEYBINDING],
  isEnabled: hasSheet,
  run(context, args = {}) {
    return runStructure(context, args, 'delete-cells');
  },
};

export const structureCommands = (): readonly CommandDefinition<StructureArgs>[] => [
  structureInsertRowsCommand,
  structureInsertColumnsCommand,
  structureInsertCellsCommand,
  structureDeleteRowsCommand,
  structureDeleteColumnsCommand,
  structureDeleteCellsCommand,
];

// ---------------------------------------------------------------------------
// Resolution and execution
// ---------------------------------------------------------------------------

function runStructure(context: CommandContext, args: StructureArgs, kind: EditKind): CommandResult {
  const sheet = resolveSheet(context, args.sheetId);
  if (sheet === undefined) return toast('Open a worksheet before editing its structure.');
  const selection = resolveSelection(context, args);
  if (selection === null) return toast('Select a cell or range first.');

  const shift = resolveShift(kind, args.shift);
  if (shift === null) {
    return blockResult(
      sheet.id,
      null,
      `"${String(args.shift)}" is not an Insert/Delete Cells direction.`,
    );
  }
  const edit = buildEdit(kind, sheet.id, selection, shift);
  const limit = edit.axis === 'row' ? MAX_ROW : MAX_COLUMN;
  if (edit.index + edit.count > limit) {
    return blockResult(sheet.id, edit, 'The edit would fall outside the worksheet.');
  }

  let plan: RangeTransformPlan;
  try {
    plan = planRangeTransform(context.workbook, edit);
  } catch (error) {
    return blockResult(sheet.id, edit, error instanceof Error ? error.message : String(error));
  }
  if (plan.block !== null) {
    return blockResult(sheet.id, edit, plan.block.message);
  }
  // The transform-or-block rule reaches the opaque parts the edit anchors
  // (charts, drawings, comments, pivots) and the sibling worksheets whose
  // hyperlinks point into the edited sheet. Their byte rewrites ride the same
  // transaction; undo writes the `before` bytes back (workbook-model §10).
  let opaquePatches: readonly StructuralPartPatch[];
  try {
    opaquePatches = planOpaquePartPatches(context.workbook, edit, plan).patches;
  } catch (error) {
    return blockResult(sheet.id, edit, error instanceof Error ? error.message : String(error));
  }

  const sourceBefore = context.workbook.sourceParts.get(sheet.partPath) ?? null;
  let sourceAfter: Uint8Array | null = null;
  if (sourceBefore !== null) {
    try {
      sourceAfter = renumberWorksheetPart({
        bytes: sourceBefore,
        partName: sheet.partPath,
        sheetName: sheet.name,
        edit,
        workbook: context.workbook,
      }).bytes;
    } catch (error) {
      return blockResult(sheet.id, edit, error instanceof Error ? error.message : String(error));
    }
  }

  const change = buildChange(sheet, edit, plan, sourceBefore, sourceAfter, opaquePatches);
  const changes: CommandChange[] = [change];
  for (const rewrite of plan.nameRewrites) {
    changes.push({
      type: 'name.set',
      scope: rewrite.scope,
      sheetId: rewrite.sheetId,
      name: rewrite.name,
      before: rewrite.before,
      after: rewrite.after,
    });
  }

  const result: CommandResult = {
    changes,
    selectionAfter: selectionAfter(kind, edit, selection),
    structure: payload(sheet.id, edit, plan),
  };
  // Broken-and-marked, one notice per edit (ticket 32): the edit applies, and
  // every reference whose target is gone already reads #REF! in the model and
  // in the saved parts. The user is told once, not blocked.
  if (plan.brokenReferences.length > 0) {
    const count = plan.brokenReferences.length;
    const first = plan.brokenReferences[0];
    return {
      ...result,
      ui: {
        type: 'toast',
        message: `${count} reference${count === 1 ? '' : 's'} could not follow the edit and now read #REF!. The first is ${first?.label ?? 'unknown'}.`,
      },
    };
  }
  return result;
}

function buildEdit(
  kind: EditKind,
  sheetId: SheetId,
  selection: GridRange,
  shift: CellShift,
): StructuralEdit {
  const rows = selection.endRow - selection.startRow + 1;
  const columns = selection.endCol - selection.startCol + 1;
  switch (kind) {
    case 'insert-rows':
      return { kind: 'insert', sheetId, axis: 'row', index: selection.startRow, count: rows };
    case 'insert-columns':
      return { kind: 'insert', sheetId, axis: 'column', index: selection.startCol, count: columns };
    case 'delete-rows':
      return { kind: 'delete', sheetId, axis: 'row', index: selection.startRow, count: rows };
    case 'delete-columns':
      return { kind: 'delete', sheetId, axis: 'column', index: selection.startCol, count: columns };
    case 'insert-cells':
      return shift === 'right'
        ? {
            kind: 'insert',
            sheetId,
            axis: 'column',
            index: selection.startCol,
            count: columns,
            band: [selection.startRow, selection.endRow],
          }
        : {
            kind: 'insert',
            sheetId,
            axis: 'row',
            index: selection.startRow,
            count: rows,
            band: [selection.startCol, selection.endCol],
          };
    case 'delete-cells':
      return shift === 'left'
        ? {
            kind: 'delete',
            sheetId,
            axis: 'column',
            index: selection.startCol,
            count: columns,
            band: [selection.startRow, selection.endRow],
          }
        : {
            kind: 'delete',
            sheetId,
            axis: 'row',
            index: selection.startRow,
            count: rows,
            band: [selection.startCol, selection.endCol],
          };
  }
}

function resolveShift(kind: EditKind, requested: CellShift | undefined): CellShift | null {
  if (kind === 'insert-cells') {
    if (requested === undefined) return 'down';
    return requested === 'right' || requested === 'down' ? requested : null;
  }
  if (kind === 'delete-cells') {
    if (requested === undefined) return 'up';
    return requested === 'up' || requested === 'left' ? requested : null;
  }
  return 'down';
}

// ---------------------------------------------------------------------------
// The one transaction
// ---------------------------------------------------------------------------

function buildChange(
  sheet: Sheet,
  edit: StructuralEdit,
  plan: RangeTransformPlan,
  sourceBefore: Uint8Array | null,
  sourceAfter: Uint8Array | null,
  parts: readonly StructuralPartPatch[],
): StructureApplyChange {
  const before = plan.mergeRewrites.map((rewrite) => rewrite.before);
  const after = plan.mergeRewrites
    .map((rewrite) => rewrite.after)
    .filter((range): range is GridRange => range !== null);
  return {
    type: 'structure.apply',
    sheetId: sheet.id,
    partPath: sheet.partPath,
    edit,
    parts,
    cells: plan.cellMoves,
    rows: rowsChange(plan),
    columns: { before: plan.columnDimensions.before, after: plan.columnDimensions.after },
    merges: { before, after },
    comments: plan.comments,
    hyperlinks: plan.hyperlinks,
    validation: {
      before: snapshotRecords(plan.validation.before),
      after: snapshotRecords(plan.validation.after),
    },
    conditionalFormats: {
      before: snapshotRecords(plan.conditionalFormats.before),
      after: snapshotRecords(plan.conditionalFormats.after),
    },
    autoFilter: { before: plan.autoFilter.before, after: plan.autoFilter.after },
    sourceBefore,
    sourceAfter,
  };
}

/**
 * A cleared filter restores the rows it hid: their hidden flag is the effective
 * saved flag, so the filter layer must give it back (sort-autofilter §3.3,
 * §3.5). A row with no other dimension fields loses its record.
 */
function rowsChange(plan: RangeTransformPlan): StructureRowsChange {
  const restored = new Set(plan.autoFilter.restoredHiddenRows);
  const after = new Map<number, RowDimension>();
  for (const [index, dimension] of plan.rowDimensions.after) {
    if (!restored.has(index) || dimension.hidden !== true) {
      after.set(index, dimension);
      continue;
    }
    const { hidden, ...rest } = dimension;
    if (hidden === true && Object.keys(rest).length > 0) after.set(index, rest);
  }
  const sorted = [...after.entries()]
    .sort((left, right) => left[0] - right[0])
    .map(([index, dimension]) => [index, dimension] as const);
  return { before: plan.rowDimensions.before, after: sorted };
}

/** Copy the range-bearing records so the transaction owns its before side. */
function snapshotRecords<T extends { readonly range: GridRange }>(
  records: readonly T[],
): readonly T[] {
  return records.map((record) => ({ ...record, range: { ...record.range } }));
}

function payload(
  sheetId: SheetId,
  edit: StructuralEdit,
  plan: RangeTransformPlan,
): StructurePayload {
  return {
    state: 'applied',
    sheetId,
    edit,
    cellCount: plan.cellMoves.length,
    formulaRewrites: plan.formulaRewrites.length,
    dimensionCount: plan.rowDimensions.after.length + plan.columnDimensions.after.length,
    filterCleared: plan.autoFilter.before !== null && plan.autoFilter.after === null,
  };
}

function blockResult(
  sheetId: SheetId,
  edit: StructuralEdit | null,
  message: string,
): CommandResult {
  return {
    changes: [],
    structure: {
      state: 'blocked',
      sheetId,
      edit,
      message,
      cellCount: 0,
      formulaRewrites: 0,
      dimensionCount: 0,
      filterCleared: false,
    },
    ui: { type: 'toast', message },
  };
}

// ---------------------------------------------------------------------------
// Selection
// ---------------------------------------------------------------------------

function selectionAfter(kind: EditKind, edit: StructuralEdit, selection: GridRange): GridRange {
  switch (kind) {
    case 'insert-rows':
      return {
        startRow: edit.index,
        startCol: selection.startCol,
        endRow: edit.index + edit.count - 1,
        endCol: selection.endCol,
      };
    case 'insert-columns':
      return {
        startRow: selection.startRow,
        startCol: edit.index,
        endRow: selection.endRow,
        endCol: edit.index + edit.count - 1,
      };
    case 'delete-rows':
      return {
        startRow: edit.index,
        startCol: selection.startCol,
        endRow: edit.index,
        endCol: selection.endCol,
      };
    case 'delete-columns':
      return {
        startRow: selection.startRow,
        startCol: edit.index,
        endRow: selection.endRow,
        endCol: edit.index,
      };
    case 'insert-cells':
    case 'delete-cells':
      return selection;
  }
}

function resolveSelection(context: CommandContext, args: StructureArgs): GridRange | null {
  if (args.range !== undefined) return args.range;
  if (context.selection !== null) return context.selection;
  const active = context.activeCell;
  if (active === null) return null;
  return { startRow: active.row, startCol: active.col, endRow: active.row, endCol: active.col };
}

function resolveSheet(context: CommandContext, sheetId?: SheetId): Sheet | undefined {
  const id = sheetId ?? context.activeSheetId;
  if (id === null || id === undefined) return undefined;
  return context.workbook.sheets.find((sheet) => sheet.id === id);
}

function hasSheet(context: CommandContext): boolean {
  return context.workbook.sheets.length > 0;
}

function toast(message: string): CommandResult {
  return { changes: [], ui: { type: 'toast', message } };
}
