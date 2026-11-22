import {
  planMovePartPatches,
  planOpaquePartPatches,
  type StructuralPartPatch,
} from '@ooxml/structure-patch';
import {
  planMoveTransform,
  planRangeTransform,
  moveDestinationRange,
  type MoveRangeEdit,
  type MoveRangePlan,
  type RangeTransformPlan,
  type StructuralEdit,
} from '@model/range-transform';
import { MAX_COLUMN, MAX_ROW } from '@model/references';
import type { CellAddress, GridRange, Sheet, SheetId } from '@model/types';
import type {
  CommandChange,
  CommandContext,
  CommandDefinition,
  CommandResult,
  StructureApplyChange,
  StructureRowsChange,
} from './types';

/**
 * The two drag gestures (feature matrix §1 "Drag-and-drop move; Shift-drag
 * insert"; ticket 49).
 *
 * `edit.dragMove` drops the selection at a new top-left: the cells move, the
 * source is cleared, the destination is overwritten and every reference that
 * points at a moved cell follows it through the range-transform module's move
 * rules (transform or block). `edit.dragInsert` drops the selection into a gap
 * the drop opens: the same move rides a banded Insert Cells & Shift edit, so
 * the cells at the drop shift down or right exactly like the structural
 * insert.
 *
 * Each gesture is one transaction: one `structure.apply` change carrying the
 * moved cells, the rewritten formulas and the moved merges, comments and
 * hyperlinks, plus one `name.set` per rewritten defined name. A drop back on
 * the source changes nothing.
 */

export interface DragMoveArgs {
  readonly sheetId?: SheetId;
  /** The dragged range. */
  readonly source: GridRange;
  /** Where the source's top-left lands. The source keeps its shape. */
  readonly to: CellAddress;
}

export interface DragInsertArgs extends DragMoveArgs {
  /** Which way the drop's gap shifts existing cells (ticket 49: down or right). */
  readonly shift: 'down' | 'right';
}

export const dragMoveCommand: CommandDefinition<DragMoveArgs> = {
  id: 'edit.dragMove',
  label: 'Drag and drop move',
  category: 'edit',
  keybindings: [],
  isEnabled: hasSheet,
  run(context, args) {
    return runDragMove(context, args);
  },
};

export const dragInsertCommand: CommandDefinition<DragInsertArgs> = {
  id: 'edit.dragInsert',
  label: 'Shift-drag insert',
  category: 'edit',
  keybindings: [],
  isEnabled: hasSheet,
  run(context, args) {
    return runDragInsert(context, args);
  },
};

export function moveCommands(): readonly CommandDefinition<never>[] {
  return [
    dragMoveCommand as unknown as CommandDefinition<never>,
    dragInsertCommand as unknown as CommandDefinition<never>,
  ];
}

function runDragMove(context: CommandContext, args: DragMoveArgs): CommandResult {
  const sheet = resolveSheet(context, args.sheetId);
  if (sheet === undefined) return toast('Open a worksheet before moving cells.');
  const source = args.source;
  if (!validRange(source)) return toast('Select a range to move first.');
  const to = args.to;
  if (
    to.row + (source.endRow - source.startRow) >= MAX_ROW ||
    to.col + (source.endCol - source.startCol) >= MAX_COLUMN
  ) {
    return blockResult(sheet.id, null, 'The move would fall outside the worksheet.');
  }
  // A drop back on the source is not a move (ticket 49).
  if (to.row === source.startRow && to.col === source.startCol) {
    return { changes: [], selectionAfter: source };
  }
  const edit: MoveRangeEdit = { kind: 'move', sheetId: sheet.id, source, to };
  return runMove(context, sheet, edit);
}

function runDragInsert(context: CommandContext, args: DragInsertArgs): CommandResult {
  const sheet = resolveSheet(context, args.sheetId);
  if (sheet === undefined) return toast('Open a worksheet before moving cells.');
  const source = args.source;
  if (!validRange(source)) return toast('Select a range to move first.');
  const to = args.to;
  const rows = source.endRow - source.startRow + 1;
  const columns = source.endCol - source.startCol + 1;
  if (to.row + rows > MAX_ROW || to.col + columns > MAX_COLUMN) {
    return blockResult(sheet.id, null, 'The move would fall outside the worksheet.');
  }
  const destination = moveDestinationRange({ kind: 'move', sheetId: sheet.id, source, to });
  if (!(
    destination.endRow < source.startRow ||
    source.endRow < destination.startRow ||
    destination.endCol < source.startCol ||
    source.endCol < destination.startCol
  )) {
    // The drop lands on the dragged range itself: nothing is inserted.
    return { changes: [], selectionAfter: source };
  }
  const prior: StructuralEdit =
    args.shift === 'right'
      ? {
          kind: 'insert',
          sheetId: sheet.id,
          axis: 'column',
          index: to.col,
          count: columns,
          band: [to.row, to.row + rows - 1],
        }
      : {
          kind: 'insert',
          sheetId: sheet.id,
          axis: 'row',
          index: to.row,
          count: rows,
          band: [to.col, to.col + columns - 1],
        };
  const edit: MoveRangeEdit = { kind: 'move', sheetId: sheet.id, source, to, prior };
  return runMove(context, sheet, edit, prior);
}

function runMove(
  context: CommandContext,
  sheet: Sheet,
  edit: MoveRangeEdit,
  prior?: StructuralEdit,
): CommandResult {
  let plan: MoveRangePlan;
  try {
    plan = planMoveTransform(context.workbook, edit);
  } catch (error) {
    return blockResult(sheet.id, edit, error instanceof Error ? error.message : String(error));
  }
  let priorPlan: RangeTransformPlan | null = null;
  if (prior !== undefined) {
    try {
      priorPlan = planRangeTransform(context.workbook, prior);
    } catch (error) {
      return blockResult(sheet.id, edit, error instanceof Error ? error.message : String(error));
    }
    if (priorPlan.block !== null) {
      return blockResult(sheet.id, edit, priorPlan.block.message);
    }
  }
  if (plan.block !== null) {
    return blockResult(sheet.id, edit, plan.block.message);
  }
  // A drop that moves nothing records nothing; the selection still follows the
  // drop, exactly as a click on the destination would leave it.
  if (!plan.changed && (priorPlan === null || !priorPlan.changed)) {
    return { changes: [], selectionAfter: moveDestinationRange(edit) };
  }

  // The byte rewrites ride the transaction: the moved cells are patched by the
  // save, and the hyperlink, rule and anchor parts travel as part patches so
  // undo restores their exact prior bytes (workbook-model §10).
  let parts: readonly StructuralPartPatch[];
  try {
    const movePatches = planMovePartPatches(context.workbook, edit, plan);
    if (movePatches.block !== null) {
      return blockResult(sheet.id, edit, movePatches.block.message);
    }
    parts = [
      ...movePatches.patches,
      ...(priorPlan === null
        ? []
        : planOpaquePartPatches(context.workbook, prior!, priorPlan).patches),
    ];
  } catch (error) {
    return blockResult(sheet.id, edit, error instanceof Error ? error.message : String(error));
  }

  const rows = rowsChange(sheet, priorPlan);
  const change: StructureApplyChange = {
    type: 'structure.apply',
    sheetId: sheet.id,
    partPath: sheet.partPath,
    edit,
    parts,
    cells: plan.cellMoves,
    rows,
    columns: {
      before: priorPlan?.columnDimensions.before ?? [...sheet.columnDimensions.entries()],
      after: priorPlan?.columnDimensions.after ?? [...sheet.columnDimensions.entries()],
    },
    merges: {
      before: plan.mergeRewrites.map((rewrite) => rewrite.before),
      after: plan.mergeRewrites
        .map((rewrite) => rewrite.after)
        .filter((range): range is GridRange => range !== null),
    },
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
    autoFilter: {
      before: priorPlan?.autoFilter.before ?? sheet.autoFilter,
      after: priorPlan?.autoFilter.after ?? sheet.autoFilter,
    },
    // The move renumbers no row or column index, so the saved baseline keeps
    // its bytes and the moved cells ride the transaction's dirty markers.
    sourceBefore: context.workbook.sourceParts.get(sheet.partPath) ?? null,
    sourceAfter: context.workbook.sourceParts.get(sheet.partPath) ?? null,
  };
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

  const destination = moveDestinationRange(edit);
  const result: CommandResult = {
    changes,
    selectionAfter: destination,
    structure: {
      state: 'applied',
      sheetId: sheet.id,
      edit,
      message: undefined,
      cellCount: plan.cellMoves.length,
      formulaRewrites: plan.formulaRewrites.length,
      dimensionCount: rows.after.length,
      filterCleared: false,
    },
  };
  if (plan.brokenReferences.length > 0) {
    const count = plan.brokenReferences.length;
    const first = plan.brokenReferences[0];
    return {
      ...result,
      ui: {
        type: 'toast',
        message: `${count} reference${count === 1 ? '' : 's'} could not follow the move and now read #REF!. The first is ${first?.label ?? 'unknown'}.`,
      },
    };
  }
  return result;
}

function rowsChange(sheet: Sheet, priorPlan: RangeTransformPlan | null): StructureRowsChange {
  const before =
    priorPlan?.rowDimensions.before ??
    [...sheet.rowDimensions.entries()].map(([index, value]) => [index, value] as const);
  const after =
    priorPlan?.rowDimensions.after ??
    [...sheet.rowDimensions.entries()].map(([index, value]) => [index, value] as const);
  return {
    before: before.map(([index, dimension]) => [index, { ...dimension }]),
    after: after.map(([index, dimension]) => [index, { ...dimension }]),
  };
}

/** Copy the range-bearing records so the transaction owns its before side. */
function snapshotRecords<T extends { readonly range: GridRange }>(
  records: readonly T[],
): readonly T[] {
  return records.map((record) => ({ ...record, range: { ...record.range } }));
}

function validRange(range: GridRange): boolean {
  return (
    Number.isInteger(range.startRow) &&
    Number.isInteger(range.startCol) &&
    Number.isInteger(range.endRow) &&
    Number.isInteger(range.endCol) &&
    range.startRow >= 0 &&
    range.startCol >= 0 &&
    range.endRow >= range.startRow &&
    range.endCol >= range.startCol &&
    range.endRow < MAX_ROW &&
    range.endCol < MAX_COLUMN
  );
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

function blockResult(sheetId: SheetId, edit: MoveRangeEdit | null, message: string): CommandResult {
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
