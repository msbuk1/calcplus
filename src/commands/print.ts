import type { PrintSetupRecord, SheetId } from '@model/types';
import type { CommandDefinition, CommandResult } from './types';

/**
 * `print.setup` — the print preview's paper list (ticket 42).
 *
 * Excel's print view persists the chosen paper as the sheet's page setup, so
 * the choice is document state and reaches the file through `<pageSetup>` on
 * save (file-compat §3, `src/ooxml/print-setup-xml.ts`). One choice is one
 * transaction: the change carries the whole record on both sides, so undo puts
 * the previous paper back and the preview re-plans from it.
 *
 * The command only owns the paper size. Orientation, scale, margins and the
 * print area keep their authored values; a sheet that declared no print setup
 * at all gets a record with only `paperSize` set, which the save writer adds to
 * the worksheet part as a fresh `<pageSetup>` element.
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

export const printSetupCommand: CommandDefinition<PrintSetupArgs> = {
  id: 'print.setup',
  label: 'Page Setup',
  category: 'view',
  keybindings: [],
  isEnabled(context) {
    return context.workbook.sheets.length > 0;
  },
  run(context, args): CommandResult {
    const sheetId = args.sheetId ?? context.activeSheetId;
    const sheet =
      sheetId === null || sheetId === undefined
        ? undefined
        : context.workbook.sheets.find((candidate) => candidate.id === sheetId);
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

function toast(message: string): CommandResult {
  return { changes: [], ui: { type: 'toast', message } };
}
