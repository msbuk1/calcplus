import type {
  BorderEdgeStyle,
  BorderStyle,
  Cell,
  GridRange,
  Sheet,
  SheetId,
  StyleId,
  StyleRecord,
  UnderlineStyle,
  Workbook,
} from '@model/types';
import { cellIsEmpty, getCell } from '@model/types';
import { borderEdge } from '@model/style';
import { findOrAppendStyle, patchedStyleInput, type StylePatch } from '@model/style-table';
import type { StyleSetChange } from './types';
import type { CommandContext, CommandDefinition, CommandResult } from './types';

/**
 * Formatting commands. Each command is one transaction, whatever the range
 * size, and stores only before/after style ids per cell so undo is exact and a
 * large format does not copy cell contents into the history.
 *
 * A command may allocate an append-only style record while it computes the
 * change. The record is not undone, because style ids are stable and never
 * renumbered; undo restores the cell references, which is the document state a
 * user sees.
 */

type Mutable<T> = { -readonly [K in keyof T]: T[K] };

export interface FormatTargetArgs {
  readonly sheetId?: SheetId;
  readonly range?: GridRange;
}

export interface FontFormatArgs extends FormatTargetArgs {
  readonly name?: string;
  readonly size?: number;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: boolean;
  readonly underlineStyle?: UnderlineStyle;
  readonly strike?: boolean;
  readonly color?: string | null;
}

export interface FillFormatArgs extends FormatTargetArgs {
  readonly pattern?: string;
  readonly color?: string | null;
  readonly backgroundColor?: string | null;
}

export type BorderEdgeName = 'left' | 'right' | 'top' | 'bottom' | 'diagonal';

export interface BorderFormatArgs extends FormatTargetArgs {
  readonly preset?: 'all' | 'outline' | 'inside' | 'none';
  readonly style?: string;
  readonly color?: string | null;
  readonly edges?: Partial<Record<BorderEdgeName, BorderEdgeStyle>>;
  readonly diagonalUp?: boolean;
  readonly diagonalDown?: boolean;
}

export interface AlignmentFormatArgs extends FormatTargetArgs {
  readonly horizontal?: StyleRecord['alignment']['horizontal'];
  readonly vertical?: StyleRecord['alignment']['vertical'];
  readonly wrapText?: boolean;
  readonly indent?: number;
  readonly textRotation?: number;
  readonly shrinkToFit?: boolean;
}

export interface NumberFormatArgs extends FormatTargetArgs {
  readonly code: string | null;
}

interface ResolvedTarget {
  readonly workbook: Workbook;
  readonly sheet: Sheet;
  readonly range: GridRange;
}

/** A formatting command is one transaction; keep the before/after patch bounded. */
const MAX_FORMAT_CELLS = 100_000;

export const formatBoldCommand: CommandDefinition<FormatTargetArgs> = toggleCommand(
  'format.bold',
  'Bold',
  'Ctrl+B',
  (style) => style.font.bold,
  (style) => ({ font: { bold: !style.font.bold } }),
);

export const formatItalicCommand: CommandDefinition<FormatTargetArgs> = toggleCommand(
  'format.italic',
  'Italic',
  'Ctrl+I',
  (style) => style.font.italic,
  (style) => ({ font: { italic: !style.font.italic } }),
);

export const formatUnderlineCommand: CommandDefinition<FormatTargetArgs> = toggleCommand(
  'format.underline',
  'Underline',
  'Ctrl+U',
  (style) => style.font.underline,
  (style) => ({ font: { underline: !style.font.underline } }),
);

export const formatStrikeCommand: CommandDefinition<FormatTargetArgs> = toggleCommand(
  'format.strike',
  'Strikethrough',
  'Ctrl+5',
  (style) => style.font.strike,
  (style) => ({ font: { strike: !style.font.strike } }),
);

export const formatFontCommand: CommandDefinition<FontFormatArgs> = {
  id: 'format.font',
  label: 'Font',
  category: 'format',
  keybindings: [],
  isEnabled: workbookHasSheet,
  run(context, args = {}) {
    const underlineStyle =
      args.underlineStyle ??
      (args.underline === undefined ? undefined : args.underline ? 'single' : 'none');
    const font: NonNullable<StylePatch['font']> = {
      ...(args.name === undefined ? {} : { name: args.name }),
      ...(args.size === undefined ? {} : { size: args.size }),
      ...(args.bold === undefined ? {} : { bold: args.bold }),
      ...(args.italic === undefined ? {} : { italic: args.italic }),
      ...(underlineStyle === undefined
        ? {}
        : { underline: underlineStyle !== 'none', underlineStyle }),
      ...(args.strike === undefined ? {} : { strike: args.strike }),
      ...(args.color === undefined ? {} : { color: args.color }),
    };
    if (Object.keys(font).length === 0) return { changes: [] };
    return applyPatch(context, args, { font });
  },
};

export const formatFillCommand: CommandDefinition<FillFormatArgs> = {
  id: 'format.fill',
  label: 'Fill',
  category: 'format',
  keybindings: [],
  isEnabled: workbookHasSheet,
  run(context, args = {}) {
    if (
      args.pattern === undefined &&
      args.color === undefined &&
      args.backgroundColor === undefined
    ) {
      return { changes: [] };
    }
    const fill: NonNullable<StylePatch['fill']> = {
      ...(args.pattern === undefined ? {} : { pattern: args.pattern }),
      ...(args.color === undefined ? {} : { color: args.color }),
      ...(args.backgroundColor === undefined ? {} : { backgroundColor: args.backgroundColor }),
    };
    return applyPatch(context, args, { fill });
  },
};

export const formatBorderCommand: CommandDefinition<BorderFormatArgs> = {
  id: 'format.border',
  label: 'Borders',
  category: 'format',
  keybindings: [],
  isEnabled: workbookHasSheet,
  run(context, args = {}) {
    const target = resolveTarget(context, args);
    if (target === null) return noTarget();
    if (args.preset === 'none') {
      return applyPatch(context, args, {
        border: {
          left: borderEdge('none', null),
          right: borderEdge('none', null),
          top: borderEdge('none', null),
          bottom: borderEdge('none', null),
          diagonal: borderEdge('none', null),
          diagonalUp: false,
          diagonalDown: false,
        },
      });
    }
    if (
      args.preset === 'outline' ||
      args.preset === 'inside' ||
      args.preset === 'all' ||
      args.edges !== undefined ||
      args.style !== undefined
    ) {
      const edge = (): BorderEdgeStyle => borderEdge(args.style ?? 'thin', args.color ?? null);
      return applyPerCellPatch(context, args, (row, col) => {
        const border: Mutable<Partial<BorderStyle>> = {};
        const preset = args.preset;
        if (preset === 'outline' || preset === 'all') {
          if (row === target.range.startRow) border.top = edge();
          if (row === target.range.endRow) border.bottom = edge();
          if (col === target.range.startCol) border.left = edge();
          if (col === target.range.endCol) border.right = edge();
        }
        if (preset === 'inside' || preset === 'all') {
          if (row > target.range.startRow) border.top = edge();
          if (row < target.range.endRow) border.bottom = edge();
          if (col > target.range.startCol) border.left = edge();
          if (col < target.range.endCol) border.right = edge();
        }
        if (preset === undefined) {
          Object.assign(border, {
            left: edge(),
            right: edge(),
            top: edge(),
            bottom: edge(),
          });
        }
        if (args.edges !== undefined) Object.assign(border, args.edges);
        if (args.diagonalUp !== undefined) border.diagonalUp = args.diagonalUp;
        if (args.diagonalDown !== undefined) border.diagonalDown = args.diagonalDown;
        return Object.keys(border).length === 0 ? null : { border };
      });
    }
    if (args.diagonalUp !== undefined || args.diagonalDown !== undefined) {
      return applyPatch(context, args, {
        border: {
          ...(args.diagonalUp === undefined ? {} : { diagonalUp: args.diagonalUp }),
          ...(args.diagonalDown === undefined ? {} : { diagonalDown: args.diagonalDown }),
        },
      });
    }
    return { changes: [] };
  },
};

export const formatAlignmentCommand: CommandDefinition<AlignmentFormatArgs> = {
  id: 'format.alignment',
  label: 'Alignment',
  category: 'format',
  keybindings: [],
  isEnabled: workbookHasSheet,
  run(context, args = {}) {
    return applyPatch(context, args, {
      alignment: {
        ...(args.horizontal === undefined ? {} : { horizontal: args.horizontal }),
        ...(args.vertical === undefined ? {} : { vertical: args.vertical }),
        ...(args.wrapText === undefined ? {} : { wrapText: args.wrapText }),
        ...(args.indent === undefined ? {} : { indent: args.indent }),
        ...(args.textRotation === undefined ? {} : { textRotation: args.textRotation }),
        ...(args.shrinkToFit === undefined ? {} : { shrinkToFit: args.shrinkToFit }),
      },
    });
  },
};

export const formatNumberFormatCommand: CommandDefinition<NumberFormatArgs> = {
  id: 'format.numberFormat',
  label: 'Number format',
  category: 'format',
  keybindings: [],
  isEnabled: workbookHasSheet,
  run(context, args = { code: null }) {
    return applyPatch(context, args, {
      numFmtCode: args.code,
    });
  },
};

export const formatClearCommand: CommandDefinition<FormatTargetArgs> = {
  id: 'format.clear',
  label: 'Clear formatting',
  category: 'format',
  keybindings: [],
  isEnabled: workbookHasSheet,
  run(context, args = {}) {
    const target = resolveTarget(context, args);
    if (target === null) return noTarget();
    const change = buildChange(target, () => 0);
    return change === null ? { changes: [] } : { changes: [change] };
  },
};

function toggleCommand(
  id: string,
  label: string,
  keybinding: string,
  checked: (style: StyleRecord) => boolean,
  patchFor: (style: StyleRecord) => StylePatch,
): CommandDefinition<FormatTargetArgs> {
  return {
    id,
    label,
    category: 'format',
    keybindings: [keybinding],
    isEnabled: workbookHasSheet,
    isChecked(context) {
      const target = resolveTarget(context, {});
      if (target === null) return false;
      return checked(baseStyle(target, context.activeCell));
    },
    run(context, args = {}) {
      return applyPatchForCell(context, args, (cell) =>
        patchFor(styleRecord(context.workbook, cell?.styleId ?? 0)),
      );
    },
  };
}

function applyPatch(
  context: CommandContext,
  args: FormatTargetArgs,
  patch: StylePatch,
): CommandResult {
  return applyPatchForCell(context, args, () => patch);
}

function applyPerCellPatch(
  context: CommandContext,
  args: FormatTargetArgs,
  patchFor: (row: number, col: number) => StylePatch | null,
): CommandResult {
  return applyPatchForCell(context, args, (_cell, row, col) => patchFor(row, col));
}

function applyPatchForCell(
  context: CommandContext,
  args: FormatTargetArgs,
  patchFor: (cell: Cell | null, row: number, col: number) => StylePatch | null,
): CommandResult {
  const target = resolveTarget(context, args);
  if (target === null) return noTarget();
  const change = buildChange(target, (cell, row, col) => {
    const patch = patchFor(cell, row, col);
    if (patch === null || Object.keys(patch).length === 0) return null;
    return findOrAppendStyle(
      target.workbook,
      patchedStyleInput(styleRecordForCell(target, cell), patch),
    );
  });
  return change === null ? { changes: [] } : { changes: [change] };
}

function buildChange(
  target: ResolvedTarget,
  styleFor: (cell: Cell | null, row: number, col: number) => StyleId | null,
): StyleSetChange | null {
  const before: (StyleId | null)[][] = [];
  const after: (StyleId | null)[][] = [];
  let changed = false;
  for (let row = target.range.startRow; row <= target.range.endRow; row += 1) {
    const beforeRow: (StyleId | null)[] = [];
    const afterRow: (StyleId | null)[] = [];
    for (let col = target.range.startCol; col <= target.range.endCol; col += 1) {
      const cell = getCell(target.sheet, row, col);
      const beforeId =
        cell === null || (cell.styleId === 0 && cellIsEmpty(cell)) ? null : cell.styleId;
      const targetId = styleFor(cell, row, col) ?? 0;
      const afterId = styleForCellAfter(cell, targetId);
      beforeRow.push(beforeId);
      afterRow.push(afterId);
      if (beforeId !== afterId) changed = true;
    }
    before.push(beforeRow);
    after.push(afterRow);
  }
  if (!changed) return null;
  return {
    type: 'style.set',
    sheetId: target.sheet.id,
    range: target.range,
    before,
    after,
  };
}

function styleForCellAfter(cell: Cell | null, targetId: StyleId): StyleId | null {
  if (cell === null) return targetId === 0 ? null : targetId;
  if (targetId === 0 && cellIsEmpty(cell)) return null;
  return targetId;
}

function styleRecordForCell(target: ResolvedTarget, cell: Cell | null): StyleRecord {
  return styleRecord(target.workbook, cell?.styleId ?? 0);
}

function resolveTarget(context: CommandContext, args: FormatTargetArgs): ResolvedTarget | null {
  const sheetId = args.sheetId ?? context.activeSheetId;
  if (sheetId === null) return null;
  const sheet = context.workbook.sheets.find((candidate) => candidate.id === sheetId);
  if (sheet === undefined) return null;
  const range = args.range ?? context.selection ?? activeRange(context);
  if (range === null) return null;
  if (
    range.startRow < 0 ||
    range.startCol < 0 ||
    range.endRow < range.startRow ||
    range.endCol < range.startCol ||
    range.endRow >= 1_048_576 ||
    range.endCol >= 16_384
  ) {
    return null;
  }
  const cellCount = (range.endRow - range.startRow + 1) * (range.endCol - range.startCol + 1);
  if (cellCount > MAX_FORMAT_CELLS) return null;
  return { workbook: context.workbook, sheet, range };
}

function activeRange(context: CommandContext): GridRange | null {
  const active = context.activeCell;
  if (active === null) return null;
  return {
    startRow: active.row,
    startCol: active.col,
    endRow: active.row,
    endCol: active.col,
  };
}

function baseStyle(
  target: ResolvedTarget,
  activeCell: CommandContext['activeCell'] = null,
): StyleRecord {
  if (
    activeCell !== null &&
    activeCell.row >= target.range.startRow &&
    activeCell.row <= target.range.endRow &&
    activeCell.col >= target.range.startCol &&
    activeCell.col <= target.range.endCol
  ) {
    const active = getCell(target.sheet, activeCell.row, activeCell.col);
    if (active !== null) return styleRecord(target.workbook, active.styleId);
  }
  for (let row = target.range.startRow; row <= target.range.endRow; row += 1) {
    for (let col = target.range.startCol; col <= target.range.endCol; col += 1) {
      const cell = getCell(target.sheet, row, col);
      if (cell !== null) return styleRecord(target.workbook, cell.styleId);
    }
  }
  return styleRecord(target.workbook, 0);
}

function styleRecord(workbook: Workbook, styleId: StyleId): StyleRecord {
  return workbook.styles.find((style) => style.id === styleId) ?? workbook.styles[0]!;
}

function workbookHasSheet(context: CommandContext): boolean {
  return context.workbook.sheets.length > 0;
}

function noTarget(): CommandResult {
  return {
    changes: [],
    ui: {
      type: 'toast',
      message: 'Select a cell or a range of at most 100,000 cells before formatting.',
    },
  };
}

/** Shared by the registry and tests; typed as the narrow formatting surface. */
export function formatCommands(): readonly CommandDefinition<never>[] {
  return [
    formatBoldCommand as unknown as CommandDefinition<never>,
    formatItalicCommand as unknown as CommandDefinition<never>,
    formatUnderlineCommand as unknown as CommandDefinition<never>,
    formatStrikeCommand as unknown as CommandDefinition<never>,
    formatFontCommand as unknown as CommandDefinition<never>,
    formatFillCommand as unknown as CommandDefinition<never>,
    formatBorderCommand as unknown as CommandDefinition<never>,
    formatAlignmentCommand as unknown as CommandDefinition<never>,
    formatNumberFormatCommand as unknown as CommandDefinition<never>,
    formatClearCommand as unknown as CommandDefinition<never>,
  ];
}
