import { getCell, type GridRange, type StyleRecord, type Workbook } from '@model/types';
import type { StyleRecordInput } from '@model/workbook';
import { findOrAppendStyle } from '@model/style-table';
import {
  applyStyleIdToTarget,
  resolveFormatTarget,
  styleRecordById,
  targetError,
  type FormatTargetArgs,
  type ResolvedTarget,
} from './format';
import type { CommandContext, CommandDefinition, CommandResult, PainterSnapshot } from './types';

/**
 * Format Painter commands (feature matrix §3).
 *
 * `painter.pickup` reads the style record at the source range's anchor cell and
 * returns it as view state; it changes nothing and records no transaction.
 * `painter.apply` writes that snapshot over the target range as one
 * `style.set` transaction, reusing the append-only style table. Esc cancels a
 * painter by clearing the snapshot in the UI: there is no cancel command.
 */

export type PainterPickupArgs = FormatTargetArgs;

export interface PainterApplyArgs extends FormatTargetArgs {
  readonly snapshot?: PainterSnapshot | null;
}

interface PickupSource {
  readonly target: ResolvedTarget;
  readonly sourceRange: GridRange;
}

export const painterPickupCommand: CommandDefinition<PainterPickupArgs> = {
  id: 'painter.pickup',
  label: 'Pick Up Format',
  category: 'format',
  keybindings: [],
  isEnabled: workbookHasSheet,
  run(context, args = {}) {
    const source = resolvePickupSource(context, args);
    if (source === null) return toast('Select a cell or range before picking up formatting.');
    const anchor = getCell(
      source.target.sheet,
      source.target.range.startRow,
      source.target.range.startCol,
    );
    const style = styleRecordById(source.target.workbook, anchor?.styleId ?? 0);
    return {
      changes: [],
      painter: {
        sheetId: source.target.sheet.id,
        range: { ...source.sourceRange },
        styleId: style.id,
        style,
      },
    };
  },
};

export const painterApplyCommand: CommandDefinition<PainterApplyArgs> = {
  id: 'painter.apply',
  label: 'Apply Format',
  category: 'format',
  keybindings: [],
  isEnabled: workbookHasSheet,
  run(context, args) {
    if (!isPainterSnapshot(args?.snapshot)) {
      return toast('Pick up formatting with the Format Painter before applying it.');
    }
    if (resolveFormatTarget(context, args) === null) {
      return targetError(context, args, 'painting');
    }
    let styleId: number;
    try {
      styleId = styleIdForSnapshot(context.workbook, args.snapshot);
    } catch {
      return toast('The picked-up format is malformed and was not applied.');
    }
    return applyStyleIdToTarget(context, args, styleId);
  },
};

/** Shared by the registry and tests. */
export function painterCommands(): readonly CommandDefinition<never>[] {
  return [
    painterPickupCommand as unknown as CommandDefinition<never>,
    painterApplyCommand as unknown as CommandDefinition<never>,
  ];
}

/**
 * A pickup reads only the anchor cell, so a whole-row or whole-column
 * selection must not be rejected by the formatting budget that bounds a
 * `style.set` transaction. Resolve the collapsed anchor range and keep the
 * original selection for the snapshot's view state.
 */
function resolvePickupSource(
  context: CommandContext,
  args: PainterPickupArgs,
): PickupSource | null {
  const sourceRange = args.range ?? context.selection ?? null;
  const anchorRange =
    sourceRange === null
      ? null
      : {
          startRow: sourceRange.startRow,
          startCol: sourceRange.startCol,
          endRow: sourceRange.startRow,
          endCol: sourceRange.startCol,
        };
  const target = resolveFormatTarget(context, {
    ...(args.sheetId === undefined ? {} : { sheetId: args.sheetId }),
    ...(anchorRange === null ? {} : { range: anchorRange }),
  });
  if (target === null) return null;
  return { target, sourceRange: sourceRange ?? target.range };
}

function styleIdForSnapshot(workbook: Workbook, snapshot: PainterSnapshot): number {
  // Resolve the snapshot by its visual record, not just its id: the same
  // workbook reuses the source record, and a snapshot carried from elsewhere
  // still applies exactly what was picked up. Existing styles win; a missing
  // record appends without renumbering.
  return findOrAppendStyle(workbook, styleInputFrom(snapshot.style));
}

function styleInputFrom(style: StyleRecord): StyleRecordInput {
  return {
    numFmtId: style.numFmtId,
    fontId: style.fontId,
    fillId: style.fillId,
    borderId: style.borderId,
    xfId: style.xfId,
    rawXml: style.rawXml,
    font: style.font,
    fill: style.fill,
    border: style.border,
    alignment: style.alignment,
    numFmtCode: style.numFmtCode,
    numFmtKind: style.numFmtKind,
  };
}

function isPainterSnapshot(value: PainterSnapshot | null | undefined): value is PainterSnapshot {
  if (value === null || value === undefined || typeof value !== 'object') return false;
  if (!Number.isInteger(value.styleId) || value.styleId < 0) return false;
  const style = value.style;
  return (
    style !== null &&
    typeof style === 'object' &&
    style.font !== null &&
    typeof style.font === 'object' &&
    style.fill !== null &&
    typeof style.fill === 'object' &&
    style.border !== null &&
    typeof style.border === 'object' &&
    style.alignment !== null &&
    typeof style.alignment === 'object'
  );
}

function workbookHasSheet(context: CommandContext): boolean {
  return context.workbook.sheets.length > 0;
}

function toast(message: string): CommandResult {
  return { changes: [], ui: { type: 'toast', message } };
}
