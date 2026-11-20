import { MAX_COLUMN, MAX_ROW, formatAreaReference, parseWholeArea } from '@model/references';
import { columnToLetters } from '@model/types';
import type {
  GridRange,
  PageMarginsRecord,
  PrintOptionsRecord,
  PrintSetupRecord,
  Sheet,
  SheetId,
} from '@model/types';
import {
  PRINT_AREA_NAME,
  isPrintAreaName,
  printAreaOf,
  printSetupStateKey,
  splitPrintAreas,
} from '@ooxml/print-setup-xml';
import type {
  CommandChange,
  CommandContext,
  CommandDefinition,
  CommandResult,
  NameSetChange,
} from './types';

/**
 * Page setup (print-pdf §4, ticket 44).
 *
 * The paper list in the print preview (ticket 42) is one small part of the
 * sheet's page setup; this module owns the rest of the v1 settings and the
 * print area:
 *
 * - `print.setup` writes the preview's paper choice.
 * - `print.applySetup` writes the Page Setup dialog's draft as one
 *   transaction: orientation, scaling, margins, paper, print options and the
 *   print area all travel together, so one Apply is one undo step.
 * - `print.setArea` / `print.clearArea` write the `_xlnm.Print_Area` system
 *   name. print-pdf §4 rules that setting or clearing it is not defined-name
 *   authoring.
 * - `print.gridlines` / `print.headings` author `printOptions`, the printed
 *   output only. The view toggles (`view.gridlines`/`view.headings`) stay
 *   screen-only chrome state.
 * - `print.pageSetup` is the dialog's opener; the chrome owns the surface.
 *
 * Every command returns the whole record on both sides of one `printSetup.set`
 * change (or one `name.set` for the print area), so undo restores the exact
 * prior state and a save rewrites only the elements that changed
 * (`print-setup-xml.ts`).
 */

export interface PrintSetupArgs {
  readonly sheetId?: SheetId;
  /** The OOXML `pageSetup@paperSize` code. */
  readonly paperSize: number;
}

/** A print setup record carrying only a paper size; every other field absent. */
export const EMPTY_PRINT_SETUP: PrintSetupRecord = {
  margins: null,
  options: null,
  orientation: null,
  fitToWidth: null,
  fitToHeight: null,
  scale: null,
  fitToPage: null,
  paperSize: null,
};

/** Excel's print-scaling bounds (print-pdf research §1.2). */
export const MIN_PRINT_SCALE_PERCENT = 10;
export const MAX_PRINT_SCALE_PERCENT = 400;
/** Excel's fit-to-page bounds; 0 means automatic. */
export const MAX_FIT_PAGES = 32767;
/** Excel's margin bounds, in inches. */
export const MAX_PRINT_MARGIN_INCHES = 49;

/** Excel's defaults for a sheet that declares no `<pageMargins>`. */
export const DEFAULT_PRINT_MARGINS: PageMarginsRecord = {
  left: 0.7,
  right: 0.7,
  top: 0.75,
  bottom: 0.75,
  header: 0.3,
  footer: 0.3,
};

/** The modelled page-setup values the dialog edits. */
export interface PageSetupPatch {
  readonly orientation?: 'portrait' | 'landscape' | null;
  /** `true` selects fit-to-width/height; `false` selects the scale percentage. */
  readonly fitToPage?: boolean | null;
  readonly scale?: number | null;
  readonly fitToWidth?: number | null;
  readonly fitToHeight?: number | null;
  readonly paperSize?: number | null;
  readonly margins?: PageMarginsRecord | null;
  /** Writes both `gridLines` and `gridLinesSet`, the printed gridlines pair. */
  readonly printGridlines?: boolean;
  readonly printHeadings?: boolean;
}

/**
 * The Page Setup dialog's submission. `printArea` is the authored text (a
 * range such as `A1:C5`, or blank to clear); `undefined` leaves the name
 * alone.
 */
export interface PageSetupPayload {
  readonly setup?: PageSetupPatch;
  readonly printArea?: string | null;
}

export interface PrintApplySetupArgs extends PageSetupPayload {
  readonly sheetId?: SheetId;
}

export type PageSetupField =
  'orientation' | 'scale' | 'fitToWidth' | 'fitToHeight' | 'paperSize' | 'margins';

/** A validation message per modelled field; empty when the patch is valid. */
export type PageSetupPatchErrors = Partial<Record<PageSetupField, string>>;

export interface PageSetupBuild {
  readonly errors: PageSetupPatchErrors;
  readonly setup: PrintSetupRecord;
}

export const printSetupCommand: CommandDefinition<PrintSetupArgs> = {
  id: 'print.setup',
  label: 'Paper size',
  category: 'view',
  keybindings: [],
  isEnabled(context) {
    return context.workbook.sheets.length > 0;
  },
  run(context, args): CommandResult {
    const sheet = resolveSheet(context, args.sheetId);
    if (sheet === undefined) return toast('Open a worksheet before choosing a paper size.');
    if (!Number.isInteger(args.paperSize) || args.paperSize <= 0) {
      return toast('That paper size is not valid. No page setup was changed.');
    }
    const before = sheet.printSetup ?? null;
    if (before?.paperSize === args.paperSize) {
      return toast('The sheet already uses that paper size.');
    }
    // Keep every imported attribute and the parsed source in `raw`, so the save
    // writer rewrites one value in place instead of replacing the element.
    const after: PrintSetupRecord = {
      ...EMPTY_PRINT_SETUP,
      ...(before ?? {}),
      paperSize: args.paperSize,
    };
    return {
      changes: [{ type: 'printSetup.set', sheetId: sheet.id, before, after }],
      ui: { type: 'toast', message: 'Changed the paper size.' },
    };
  },
};

/**
 * Chrome-owned dialog opener (the Format Cells pattern). The dialog submits
 * `print.applySetup`; the registry owns the id and enablement.
 */
export const printPageSetupCommand: CommandDefinition<void> = {
  id: 'print.pageSetup',
  label: 'Page Setup',
  category: 'view',
  keybindings: [],
  isEnabled(context) {
    return context.activeSheetId !== null && context.busy !== true;
  },
  run(): CommandResult {
    return { changes: [], ui: { type: 'dialog' } };
  },
};

/**
 * The Page Setup dialog's submission: one transaction for the whole draft.
 * The print area rides in the same transaction as a `name.set` change, so a
 * dialog that edits both undoes in one step.
 */
export const printApplySetupCommand: CommandDefinition<PrintApplySetupArgs> = {
  id: 'print.applySetup',
  label: 'Apply page setup',
  category: 'view',
  keybindings: [],
  isEnabled(context) {
    return context.activeSheetId !== null && context.busy !== true;
  },
  run(context, args = {}): CommandResult {
    const sheet = resolveSheet(context, args.sheetId);
    if (sheet === undefined) return toast('Open a worksheet before changing its page setup.');

    const before = sheet.printSetup ?? null;
    const built = buildPageSetup(before, args.setup ?? {});
    const problem = firstError(built.errors);
    if (problem !== null) return toast(problem);

    const changes: CommandChange[] = [];
    if (setupKey(before) !== setupKey(built.setup)) {
      changes.push({ type: 'printSetup.set', sheetId: sheet.id, before, after: built.setup });
    }
    if (args.printArea !== undefined) {
      const area = printAreaChange(sheet, args.printArea ?? '');
      if (!area.ok) return toast(area.message);
      if (area.change !== null) changes.push(area.change);
    }
    if (changes.length === 0) return { changes: [] };
    return {
      changes,
      ui: { type: 'toast', message: 'Changed the page setup.' },
    };
  },
};

export interface PrintAreaArgs {
  readonly sheetId?: SheetId;
  /** The area to print; defaults to the selection, then the active cell. */
  readonly range?: GridRange;
}

/** Set Print Area: the print area is the selected range on the active sheet. */
export const printSetAreaCommand: CommandDefinition<PrintAreaArgs> = {
  id: 'print.setArea',
  label: 'Set Print Area',
  category: 'view',
  keybindings: [],
  isEnabled(context) {
    return context.activeSheetId !== null;
  },
  run(context, args = {}): CommandResult {
    const sheet = resolveSheet(context, args.sheetId);
    if (sheet === undefined) return toast('Open a worksheet before setting a print area.');
    const range = args.range ?? context.selection ?? cellRange(context.activeCell);
    if (range === undefined) return toast('Select the cells to print first.');
    return applyPrintArea(sheet, printAreaTextForRange(range));
  },
};

/** Clear Print Area: removes the sheet's `_xlnm.Print_Area` name. */
export const printClearAreaCommand: CommandDefinition<PrintAreaArgs> = {
  id: 'print.clearArea',
  label: 'Clear Print Area',
  category: 'view',
  keybindings: [],
  isEnabled(context) {
    return context.activeSheetId !== null;
  },
  run(context, args = {}): CommandResult {
    const sheet = resolveSheet(context, args.sheetId);
    if (sheet === undefined) return toast('Open a worksheet before clearing its print area.');
    return applyPrintArea(sheet, '');
  },
};

/**
 * Print Gridlines. `printOptions@gridLines` and `@gridLinesSet` are both set,
 * because the planner prints gridlines only when both are true. The screen
 * toggle is a different command (`view.gridlines`) and is never touched.
 */
export const printGridlinesCommand: CommandDefinition<PrintAreaArgs> = {
  id: 'print.gridlines',
  label: 'Print Gridlines',
  category: 'view',
  keybindings: [],
  isEnabled(context) {
    return context.activeSheetId !== null;
  },
  isChecked(context) {
    const options = activePrintOptions(context);
    return options?.gridLines === true && options.gridLinesSet === true;
  },
  run(context, args = {}): CommandResult {
    const sheet = resolveSheet(context, args.sheetId);
    if (sheet === undefined) return toast('Open a worksheet before changing print options.');
    const options = sheet.printSetup?.options ?? null;
    const checked = options?.gridLines === true && options.gridLinesSet === true;
    return applyPrintOptions(
      sheet,
      { printGridlines: !checked },
      checked ? 'Turned off printed gridlines.' : 'Turned on printed gridlines.',
    );
  },
};

/** Print Headings: `printOptions@headings`, printed output only. */
export const printHeadingsCommand: CommandDefinition<PrintAreaArgs> = {
  id: 'print.headings',
  label: 'Print Headings',
  category: 'view',
  keybindings: [],
  isEnabled(context) {
    return context.activeSheetId !== null;
  },
  isChecked(context) {
    return activePrintOptions(context)?.headings === true;
  },
  run(context, args = {}): CommandResult {
    const sheet = resolveSheet(context, args.sheetId);
    if (sheet === undefined) return toast('Open a worksheet before changing print options.');
    const checked = sheet.printSetup?.options?.headings === true;
    return applyPrintOptions(
      sheet,
      { printHeadings: !checked },
      checked ? 'Turned off printed headings.' : 'Turned on printed headings.',
    );
  },
};

/**
 * Resolve a patch against the current record. The `raw` source is carried
 * through so a save rewrites only the elements whose state key moved.
 */
export function buildPageSetup(
  before: PrintSetupRecord | null,
  patch: PageSetupPatch,
): PageSetupBuild {
  const base = before ?? EMPTY_PRINT_SETUP;
  const errors: PageSetupPatchErrors = {};

  if (
    patch.orientation !== undefined &&
    patch.orientation !== null &&
    patch.orientation !== 'portrait' &&
    patch.orientation !== 'landscape'
  ) {
    errors.orientation = 'Choose portrait or landscape.';
  }
  if (
    patch.scale !== undefined &&
    patch.scale !== null &&
    !isIntegerInRange(patch.scale, MIN_PRINT_SCALE_PERCENT, MAX_PRINT_SCALE_PERCENT)
  ) {
    errors.scale = `Scale must be a whole number from ${MIN_PRINT_SCALE_PERCENT} to ${MAX_PRINT_SCALE_PERCENT}.`;
  }
  if (
    patch.fitToWidth !== undefined &&
    patch.fitToWidth !== null &&
    !isIntegerInRange(patch.fitToWidth, 0, MAX_FIT_PAGES)
  ) {
    errors.fitToWidth = `Fit to width must be a whole number from 0 to ${MAX_FIT_PAGES}.`;
  }
  if (
    patch.fitToHeight !== undefined &&
    patch.fitToHeight !== null &&
    !isIntegerInRange(patch.fitToHeight, 0, MAX_FIT_PAGES)
  ) {
    errors.fitToHeight = `Fit to height must be a whole number from 0 to ${MAX_FIT_PAGES}.`;
  }
  if (
    patch.paperSize !== undefined &&
    patch.paperSize !== null &&
    (!Number.isInteger(patch.paperSize) || patch.paperSize <= 0)
  ) {
    errors.paperSize = 'Choose a paper size from the list.';
  }
  if (patch.margins !== undefined && patch.margins !== null) {
    for (const value of Object.values(patch.margins)) {
      if (!Number.isFinite(value) || value < 0 || value > MAX_PRINT_MARGIN_INCHES) {
        errors.margins = `Margins must be between 0 and ${MAX_PRINT_MARGIN_INCHES} inches.`;
        break;
      }
    }
  }

  const options = printOptionsAfter(base.options ?? null, patch);
  const setup: PrintSetupRecord = {
    margins: patch.margins !== undefined ? patch.margins : base.margins,
    options: options === base.options ? base.options : options,
    orientation: patch.orientation !== undefined ? patch.orientation : base.orientation,
    fitToWidth: patch.fitToWidth !== undefined ? patch.fitToWidth : base.fitToWidth,
    fitToHeight: patch.fitToHeight !== undefined ? patch.fitToHeight : base.fitToHeight,
    scale: patch.scale !== undefined ? patch.scale : base.scale,
    fitToPage: patch.fitToPage !== undefined ? patch.fitToPage : base.fitToPage,
    paperSize: patch.paperSize !== undefined ? patch.paperSize : base.paperSize,
    ...(base.raw === undefined ? {} : { raw: base.raw }),
  };
  return { errors, setup };
}

/** The page-setup values the dialog opens with, from the sheet's model state. */
export interface PageSetupValues {
  readonly orientation: 'portrait' | 'landscape';
  readonly fitToPage: boolean;
  readonly fitToWidth: number;
  readonly fitToHeight: number;
  readonly scale: number;
  readonly paperSize: number | null;
  readonly margins: PageMarginsRecord;
  readonly printAreaText: string;
  readonly printGridlines: boolean;
  readonly printHeadings: boolean;
}

/**
 * The dialog's starting values. Fit-to-page wins over the scale when the sheet
 * declares it, exactly as the planner reads the pair.
 */
export function pageSetupValuesOf(sheet: Sheet): PageSetupValues {
  const setup = sheet.printSetup ?? null;
  const options = setup?.options ?? null;
  return {
    orientation: setup?.orientation ?? 'portrait',
    fitToPage: setup?.fitToPage === true,
    fitToWidth: setup?.fitToWidth ?? 1,
    fitToHeight: setup?.fitToHeight ?? 0,
    scale: setup?.scale ?? 100,
    paperSize: setup?.paperSize ?? null,
    margins: setup?.margins ?? DEFAULT_PRINT_MARGINS,
    printAreaText: formatPrintAreaRanges(printAreaOf(sheet)),
    printGridlines: options?.gridLines === true && options.gridLinesSet === true,
    printHeadings: options?.headings === true,
  };
}

export type PrintAreaTextResult =
  | { readonly ok: true; readonly formula: string | null; readonly ranges: readonly GridRange[] }
  | { readonly ok: false; readonly message: string };

/**
 * Normalise the dialog's print-area text to the name's formula. A blank field
 * clears the name. One or more comma-separated areas are accepted, each on
 * this sheet, which is the shape Excel writes for a multi-area name (imported
 * files already carry it; print-pdf §4 prints each area as its own page set).
 */
export function normalisePrintAreaText(text: string, sheetName: string): PrintAreaTextResult {
  const trimmed = text.trim();
  if (trimmed === '') return { ok: true, formula: null, ranges: [] };
  const parts = splitPrintAreas(trimmed);
  const ranges: GridRange[] = [];
  const bodies: string[] = [];
  for (const part of parts) {
    const area = parseWholeArea(part);
    if (area === null) {
      return {
        ok: false,
        message:
          'Type one or more ranges, for example A1:C5, or leave the box blank to print the used range.',
      };
    }
    if (area.sheetName !== null && area.sheetName.toUpperCase() !== sheetName.toUpperCase()) {
      return { ok: false, message: `The print area must be on ${sheetName}.` };
    }
    ranges.push({
      startRow: area.startRow,
      startCol: area.startCol,
      endRow: area.endRow,
      endCol: area.endCol,
    });
    bodies.push(formatAreaReference({ ...area, prefix: null }, true));
  }
  const prefix = `'${sheetName.replaceAll("'", "''")}'!`;
  return { ok: true, formula: bodies.map((body) => `${prefix}${body}`).join(','), ranges };
}

/**
 * A range as Excel's Print area text: `$A$1:$C$5`, `$A:$C` for whole columns
 * and `$1:$5` for whole rows. No sheet prefix.
 */
export function printAreaTextForRange(range: GridRange): string {
  const wholeColumns = range.startRow === 0 && range.endRow === MAX_ROW - 1;
  const wholeRows = range.startCol === 0 && range.endCol === MAX_COLUMN - 1;
  if (wholeColumns) return `$${columnToLetters(range.startCol)}:$${columnToLetters(range.endCol)}`;
  if (wholeRows) return `$${range.startRow + 1}:$${range.endRow + 1}`;
  const start = `$${columnToLetters(range.startCol)}$${range.startRow + 1}`;
  const end = `$${columnToLetters(range.endCol)}$${range.endRow + 1}`;
  return start === end ? start : `${start}:${end}`;
}

/** One or more print areas as the dialog's field text, without the sheet prefix. */
export function formatPrintAreaRanges(ranges: readonly GridRange[]): string {
  return ranges.map((range) => printAreaTextForRange(range)).join(', ');
}

// ---------------------------------------------------------------------------
// Internals
// ---------------------------------------------------------------------------

function resolveSheet(
  context: Pick<CommandContext, 'workbook' | 'activeSheetId'>,
  sheetId: SheetId | undefined,
): Sheet | undefined {
  const id = sheetId ?? context.activeSheetId;
  if (id === null) return undefined;
  return context.workbook.sheets.find((candidate) => candidate.id === id);
}

function activePrintOptions(
  context: Pick<CommandContext, 'workbook' | 'activeSheetId'>,
): PrintOptionsRecord | null {
  const sheet = resolveSheet(context, undefined);
  return sheet?.printSetup?.options ?? null;
}

function cellRange(
  cell: { readonly row: number; readonly col: number } | null | undefined,
): GridRange | undefined {
  if (cell === null || cell === undefined) return undefined;
  return { startRow: cell.row, startCol: cell.col, endRow: cell.row, endCol: cell.col };
}

function printOptionsAfter(
  before: PrintOptionsRecord | null,
  patch: PageSetupPatch,
): PrintOptionsRecord | null {
  if (patch.printGridlines === undefined && patch.printHeadings === undefined) return before;
  const base: PrintOptionsRecord = {
    gridLines: before?.gridLines ?? false,
    gridLinesSet: before?.gridLinesSet ?? false,
    headings: before?.headings ?? false,
    horizontalCentered: before?.horizontalCentered ?? false,
    verticalCentered: before?.verticalCentered ?? false,
  };
  return {
    ...base,
    gridLines: patch.printGridlines ?? base.gridLines,
    gridLinesSet: patch.printGridlines ?? base.gridLinesSet,
    headings: patch.printHeadings ?? base.headings,
  };
}

function setupKey(setup: PrintSetupRecord | null): string {
  return setup === null ? 'none' : printSetupStateKey(setup);
}

function firstError(errors: PageSetupPatchErrors): string | null {
  for (const message of Object.values(errors)) {
    if (message !== undefined) return message;
  }
  return null;
}

function isIntegerInRange(value: number, min: number, max: number): boolean {
  return Number.isInteger(value) && value >= min && value <= max;
}

type PrintAreaChange =
  | { readonly ok: true; readonly change: NameSetChange | null }
  | { readonly ok: false; readonly message: string };

function printAreaChange(sheet: Sheet, text: string): PrintAreaChange {
  const existing = sheet.names.find((candidate) => isPrintAreaName(candidate.name));
  const before = existing?.formula ?? null;
  const normalised = normalisePrintAreaText(text, sheet.name);
  if (!normalised.ok) return { ok: false, message: normalised.message };
  if (normalised.formula === before) return { ok: true, change: null };
  return {
    ok: true,
    change: {
      type: 'name.set',
      scope: 'sheet',
      sheetId: sheet.id,
      name: existing?.name ?? PRINT_AREA_NAME,
      before,
      after: normalised.formula,
    },
  };
}

function applyPrintArea(sheet: Sheet, text: string): CommandResult {
  const result = printAreaChange(sheet, text);
  if (!result.ok) return toast(result.message);
  if (result.change === null) {
    return toast(
      text.trim() === '' ? 'The sheet has no print area.' : 'That is already the print area.',
    );
  }
  return {
    changes: [result.change],
    ui: {
      type: 'toast',
      message: result.change.after === null ? 'Cleared the print area.' : 'Set the print area.',
    },
  };
}

function applyPrintOptions(sheet: Sheet, patch: PageSetupPatch, message: string): CommandResult {
  const before = sheet.printSetup ?? null;
  const built = buildPageSetup(before, patch);
  if (setupKey(before) === setupKey(built.setup)) {
    return toast('The sheet already uses that print setting.');
  }
  return {
    changes: [{ type: 'printSetup.set', sheetId: sheet.id, before, after: built.setup }],
    ui: { type: 'toast', message },
  };
}

function toast(message: string): CommandResult {
  return { changes: [], ui: { type: 'toast', message } };
}
