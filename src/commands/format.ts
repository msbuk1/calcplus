import { isValidFormat } from 'numfmt';
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
import { normalizeFormatCode } from '@view/number-format';
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

export interface ResolvedTarget {
  readonly workbook: Workbook;
  readonly sheet: Sheet;
  readonly range: GridRange;
}

type NormalizedColor =
  { readonly ok: true; readonly color: string | null } | { readonly ok: false };

const BORDER_EDGE_NAMES: readonly BorderEdgeName[] = ['left', 'right', 'top', 'bottom', 'diagonal'];

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

/**
 * Wrap Text (Alignment: horizontal/vertical, indent, wrap text). A toggle like
 * Bold: one transaction, checked from the active cell's style, and the patch is
 * applied per cell so a mixed selection flips each cell's own flag.
 */
export const formatWrapTextCommand: CommandDefinition<FormatTargetArgs> = toggleCommand(
  'format.wrapText',
  'Wrap Text',
  null,
  (style) => style.alignment.wrapText,
  (style) => ({ alignment: { wrapText: !style.alignment.wrapText } }),
);

export const formatFontCommand: CommandDefinition<FontFormatArgs> = {
  id: 'format.font',
  label: 'Font',
  category: 'format',
  keybindings: [],
  isEnabled: workbookHasSheet,
  run(context, args = {}) {
    const color = args.color === undefined ? null : normalizeColorInput(args.color);
    if (color !== null && !color.ok) return invalidColor();
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
      ...(color === null ? {} : { color: color.color }),
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
    const color = args.color === undefined ? null : normalizeColorInput(args.color);
    if (color !== null && !color.ok) return invalidColor();
    const backgroundColor =
      args.backgroundColor === undefined ? null : normalizeColorInput(args.backgroundColor);
    if (backgroundColor !== null && !backgroundColor.ok) return invalidColor();
    const fill: NonNullable<StylePatch['fill']> = {
      ...(args.pattern === undefined ? {} : { pattern: args.pattern }),
      ...(color === null ? {} : { color: color.color }),
      ...(backgroundColor === null ? {} : { backgroundColor: backgroundColor.color }),
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
    const target = resolveFormatTarget(context, args);
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
    const preset = args.preset;
    const hasCustomEdges = args.edges !== undefined;
    if (preset !== undefined || hasCustomEdges || args.style !== undefined) {
      // An explicit custom edge is authoritative. Edges that the caller did not
      // name stay untouched: a custom top edge must not quietly add a thin
      // border to the other three sides.
      let customEdges: Partial<Record<BorderEdgeName, BorderEdgeStyle>> | undefined;
      if (args.edges !== undefined) {
        customEdges = {};
        for (const name of BORDER_EDGE_NAMES) {
          const custom = args.edges[name];
          if (custom === undefined) continue;
          const customColor = normalizeColorInput(custom.color);
          if (!customColor.ok) return invalidColor();
          customEdges[name] = borderEdge(custom.style, customColor.color);
        }
      }
      const color = args.color === undefined ? null : normalizeColorInput(args.color);
      if (color !== null && !color.ok) return invalidColor();
      const edge = (): BorderEdgeStyle =>
        borderEdge(args.style ?? 'thin', color === null ? null : color.color);
      return applyPerCellPatch(context, args, (row, col) => {
        const border: Mutable<Partial<BorderStyle>> = {};
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
        if (preset === undefined && customEdges === undefined && args.style !== undefined) {
          Object.assign(border, {
            left: edge(),
            right: edge(),
            top: edge(),
            bottom: edge(),
          });
        }
        if (customEdges !== undefined) Object.assign(border, customEdges);
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
    const code = args.code ?? null;
    if (code !== null && !numberFormatIsValid(code)) return invalidNumberFormat();
    return applyPatch(context, args, {
      numFmtCode: code,
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
    return applyStyleIdToTarget(context, args, 0);
  },
};

/**
 * Increase/Decrease Decimal (Number formats: built-in categories + full custom
 * grammar). Excel adjusts the format code of the selection by one decimal
 * place, so the value is untouched and the cached result is unchanged; the
 * command reads the active cell's code and applies one new code to every cell
 * in the target, as one transaction.
 */
export const formatIncreaseDecimalCommand: CommandDefinition<FormatTargetArgs> = decimalCommand(
  'format.increaseDecimal',
  'Increase decimal',
  1,
);

export const formatDecreaseDecimalCommand: CommandDefinition<FormatTargetArgs> = decimalCommand(
  'format.decreaseDecimal',
  'Decrease decimal',
  -1,
);

/**
 * Chrome-owned dialog opener. The dialog itself submits `format.font`,
 * `format.fill`, `format.border`, `format.alignment` and `format.numberFormat`
 * with a payload (ui-chrome §4.1); the registry owns the command id and
 * enablement, and the run is a no-op like the other chrome commands.
 */
export const formatCellsDialogCommand: CommandDefinition<void> = {
  id: 'format.cellsDialog',
  label: 'Format Cells',
  category: 'format',
  keybindings: ['Ctrl+1'],
  isEnabled(context) {
    return context.activeSheetId !== null && context.busy !== true;
  },
  run() {
    return { changes: [] };
  },
};

function toggleCommand(
  id: string,
  label: string,
  keybinding: string | null,
  checked: (style: StyleRecord) => boolean,
  patchFor: (style: StyleRecord) => StylePatch,
): CommandDefinition<FormatTargetArgs> {
  return {
    id,
    label,
    category: 'format',
    keybindings: keybinding === null ? [] : [keybinding],
    isEnabled: workbookHasSheet,
    isChecked(context) {
      const target = resolveFormatTarget(context, {});
      if (target === null) return false;
      return checked(baseStyle(target, context.activeCell));
    },
    run(context, args = {}) {
      return applyPatchForCell(context, args, (cell) =>
        patchFor(styleRecordById(context.workbook, cell?.styleId ?? 0)),
      );
    },
  };
}

function decimalCommand(
  id: string,
  label: string,
  delta: 1 | -1,
): CommandDefinition<FormatTargetArgs> {
  return {
    id,
    label,
    category: 'format',
    keybindings: [],
    isEnabled: workbookHasSheet,
    run(context, args = {}) {
      const target = resolveFormatTarget(context, args);
      if (target === null) return noTarget();
      const current = styleRecordForCell(target, activeCellOf(target, context)).numFmtCode;
      const next = adjustDecimalPlaces(current, delta);
      if (next === null) return { changes: [] };
      if (!numberFormatIsValid(next)) return invalidNumberFormat();
      return applyPatch(context, args, { numFmtCode: next });
    },
  };
}

/** The cell whose format seeds a selection-wide number-format step. */
function activeCellOf(target: ResolvedTarget, context: CommandContext): Cell | null {
  const active = context.activeCell;
  if (
    active !== null &&
    active.row >= target.range.startRow &&
    active.row <= target.range.endRow &&
    active.col >= target.range.startCol &&
    active.col <= target.range.endCol
  ) {
    const cell = getCell(target.sheet, active.row, active.col);
    if (cell !== null) return cell;
  }
  return getCell(target.sheet, target.range.startRow, target.range.startCol);
}

/**
 * One decimal-place step over the first section of a number-format code. The
 * general format gains two places on the first increase (`0.00`) and a plain
 * integer on the first decrease; a patterned code gains or loses one `0`
 * after its decimal point. `null` means "nothing to remove".
 */
export function adjustDecimalPlaces(code: string | null, delta: 1 | -1): string | null {
  const source = (code ?? '').trim();
  if (source === '' || /^general$/i.test(source) || source === '@') {
    return delta > 0 ? '0.00' : '0';
  }
  const [first = '', ...rest] = source.split(';');
  if (first === '') return null;
  const point = first.indexOf('.');
  const next = point < 0 ? addDecimalPoint(first, delta) : stepDecimalRun(first, point, delta);
  if (next === null || next === first) return null;
  return [next, ...rest].join(';');
}

function addDecimalPoint(section: string, delta: 1 | -1): string | null {
  if (delta < 0) return null;
  // Place the point after the last digit placeholder: `0%` becomes `0.0%`.
  for (let index = section.length - 1; index >= 0; index -= 1) {
    if ('0#?'.includes(section[index]!)) {
      return `${section.slice(0, index + 1)}.0${section.slice(index + 1)}`;
    }
  }
  return '0.00';
}

function stepDecimalRun(section: string, point: number, delta: 1 | -1): string | null {
  let runEnd = point + 1;
  while (runEnd < section.length && '0#?'.includes(section[runEnd]!)) runEnd += 1;
  const runLength = runEnd - point - 1;
  if (delta > 0) {
    return `${section.slice(0, runEnd)}0${section.slice(runEnd)}`;
  }
  if (runLength === 0) {
    // `0.` has no place left to remove: drop the point itself.
    return `${section.slice(0, point)}${section.slice(point + 1)}`;
  }
  const trimmed = `${section.slice(0, runEnd - 1)}${section.slice(runEnd)}`;
  return runLength === 1 ? `${trimmed.slice(0, point)}${trimmed.slice(point + 1)}` : trimmed;
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
  const target = resolveFormatTarget(context, args);
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
  pruneUnstyledBlanks = false,
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
      const afterId = styleForCellAfter(cell, targetId, pruneUnstyledBlanks);
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

function styleForCellAfter(
  cell: Cell | null,
  targetId: StyleId,
  pruneUnstyledBlanks: boolean,
): StyleId | null {
  if (cell === null) return targetId === 0 ? null : targetId;
  // A formatting patch may bring a blank cell back to the default style (for
  // example clearing the last underline or colour). The cell record stays so
  // the caller still sees the style it just applied. Only an explicit style-id
  // apply (Clear Formats, painter) removes a blank that ends up unstyled.
  if (pruneUnstyledBlanks && targetId === 0 && cellIsEmpty(cell)) return null;
  return targetId;
}

function styleRecordForCell(target: ResolvedTarget, cell: Cell | null): StyleRecord {
  return styleRecordById(target.workbook, cell?.styleId ?? 0);
}

/**
 * Apply one existing style id to a target range as a single `style.set`
 * change. The Format Painter uses this with its picked-up record; no style is
 * appended and undo restores the exact prior style ids.
 */
export function applyStyleIdToTarget(
  context: CommandContext,
  args: FormatTargetArgs,
  styleId: StyleId,
): CommandResult {
  const target = resolveFormatTarget(context, args);
  if (target === null) return noTarget();
  const change = buildChange(target, () => styleId, true);
  return change === null ? { changes: [] } : { changes: [change] };
}

export function resolveFormatTarget(
  context: CommandContext,
  args: FormatTargetArgs,
): ResolvedTarget | null {
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
    if (active !== null) return styleRecordById(target.workbook, active.styleId);
  }
  for (let row = target.range.startRow; row <= target.range.endRow; row += 1) {
    for (let col = target.range.startCol; col <= target.range.endCol; col += 1) {
      const cell = getCell(target.sheet, row, col);
      if (cell !== null) return styleRecordById(target.workbook, cell.styleId);
    }
  }
  return styleRecordById(target.workbook, 0);
}

export function styleRecordById(workbook: Workbook, styleId: StyleId): StyleRecord {
  return workbook.styles.find((style) => style.id === styleId) ?? workbook.styles[0]!;
}

function workbookHasSheet(context: CommandContext): boolean {
  return context.workbook.sheets.length > 0;
}

function normalizeColorInput(value: string | null): NormalizedColor {
  if (value === null) return { ok: true, color: null };
  if (typeof value !== 'string') return { ok: false };
  // The dialog accepts Excel's #AARRGGBB spelling. The model paints CSS
  // #RRGGBB, so the alpha byte is dropped rather than passed to the renderer
  // (where an eight-digit value would be read as #RRGGBBAA).
  const match = /^#?([0-9a-f]{3}|[0-9a-f]{6}|[0-9a-f]{8})$/i.exec(value.trim());
  if (match === null) return { ok: false };
  const hex = match[1]!;
  const rgb = hex.length === 8 ? hex.slice(2) : hex;
  const expanded =
    rgb.length === 3
      ? rgb
          .split('')
          .map((character) => character + character)
          .join('')
      : rgb;
  return { ok: true, color: `#${expanded.toUpperCase()}` };
}

function numberFormatIsValid(code: string): boolean {
  if (typeof code !== 'string') return false;
  const trimmed = code.trim();
  if (trimmed === '' || trimmed === 'General') return true;
  const normalized = normalizeFormatCode(code);
  if (normalized.trim() === '') return false;
  try {
    return isValidFormat(normalized);
  } catch {
    return false;
  }
}

function invalidColor(): CommandResult {
  return toast('That colour is not valid. Use a hex colour such as #FF0000.');
}

function invalidNumberFormat(): CommandResult {
  return toast('That is not a valid number format. The cells were not changed.');
}

function toast(message: string): CommandResult {
  return { changes: [], ui: { type: 'toast', message } };
}

function noTarget(): CommandResult {
  return toast('Select a cell or a range of at most 100,000 cells before formatting.');
}

/** Shared by the registry and tests; typed as the narrow formatting surface. */
export function formatCommands(): readonly CommandDefinition<never>[] {
  return [
    formatBoldCommand as unknown as CommandDefinition<never>,
    formatItalicCommand as unknown as CommandDefinition<never>,
    formatUnderlineCommand as unknown as CommandDefinition<never>,
    formatStrikeCommand as unknown as CommandDefinition<never>,
    formatWrapTextCommand as unknown as CommandDefinition<never>,
    formatFontCommand as unknown as CommandDefinition<never>,
    formatFillCommand as unknown as CommandDefinition<never>,
    formatBorderCommand as unknown as CommandDefinition<never>,
    formatAlignmentCommand as unknown as CommandDefinition<never>,
    formatNumberFormatCommand as unknown as CommandDefinition<never>,
    formatIncreaseDecimalCommand as unknown as CommandDefinition<never>,
    formatDecreaseDecimalCommand as unknown as CommandDefinition<never>,
    formatClearCommand as unknown as CommandDefinition<never>,
    formatCellsDialogCommand as unknown as CommandDefinition<never>,
  ];
}
