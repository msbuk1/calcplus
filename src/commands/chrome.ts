import { downloadBytes } from '@io/file-system';
import type { PrintPlan } from '../print/types';
import type { CommandDefinition } from './types';

/**
 * Chrome-only commands. Their `run` is a no-op: the application shell owns the
 * browser file APIs and transient view state, and dispatches these IDs through
 * its command host. They still live in the command registry so enablement and
 * checked state have one source (command-system §1).
 */

const fileNewCommand: CommandDefinition<void> = {
  id: 'file.new',
  label: 'New',
  category: 'file',
  keybindings: [],
  isEnabled(context) {
    return context.busy !== true;
  },
  run() {
    return { changes: [] };
  },
};

const fileOpenCommand: CommandDefinition<void> = {
  id: 'file.open',
  label: 'Open',
  category: 'file',
  keybindings: ['Ctrl+O'],
  isEnabled(context) {
    return context.busy !== true;
  },
  run() {
    return { changes: [] };
  },
};

const fileSaveCommand: CommandDefinition<void> = {
  id: 'file.save',
  label: 'Save',
  category: 'file',
  keybindings: ['Ctrl+S'],
  isEnabled(context) {
    return context.workbook.sheets.length > 0 && context.busy !== true;
  },
  run() {
    return { changes: [] };
  },
};

const fileSaveAsCommand: CommandDefinition<void> = {
  id: 'file.saveAs',
  label: 'Save As',
  category: 'file',
  keybindings: ['Ctrl+Shift+S'],
  isEnabled(context) {
    return context.workbook.sheets.length > 0 && context.busy !== true;
  },
  run() {
    return { changes: [] };
  },
};

const fileCloseCommand: CommandDefinition<void> = {
  id: 'file.close',
  label: 'Close',
  category: 'file',
  keybindings: [],
  isEnabled(context) {
    return context.workbook.sheets.length > 0 && context.busy !== true;
  },
  run() {
    return { changes: [] };
  },
};

/**
 * File → Print opens the preview window. Like the other file commands, its
 * `run` is a no-op: the shell builds the plan from the live display records and
 * owns the preview dialog (print-pdf §2).
 */
const filePrintCommand: CommandDefinition<void> = {
  id: 'file.print',
  label: 'Print',
  category: 'file',
  keybindings: ['Ctrl+P'],
  isEnabled(context) {
    return context.workbook.sheets.length > 0 && context.busy !== true;
  },
  run() {
    return { changes: [] };
  },
};

/**
 * File → Export PDF. Like Print, the command's `run` is a no-op: the shell
 * owns the live display records and the plan, and calls `exportPdfPlan` below.
 * The map entry lands before any UI (ui-chrome §4.1).
 */
const fileExportPdfCommand: CommandDefinition<void> = {
  id: 'file.exportPdf',
  label: 'Export PDF',
  category: 'file',
  keybindings: [],
  isEnabled(context) {
    return context.workbook.sheets.length > 0 && context.busy !== true;
  },
  run() {
    return { changes: [] };
  },
};

const viewGridlinesCommand: CommandDefinition<void> = {
  id: 'view.gridlines',
  label: 'Gridlines',
  category: 'view',
  keybindings: [],
  isEnabled(context) {
    return context.workbook.sheets.length > 0;
  },
  isChecked(context) {
    return context.view?.showGridLines ?? false;
  },
  run() {
    return { changes: [] };
  },
};

const viewHeadingsCommand: CommandDefinition<void> = {
  id: 'view.headings',
  label: 'Headings',
  category: 'view',
  keybindings: [],
  isEnabled(context) {
    return context.workbook.sheets.length > 0;
  },
  isChecked(context) {
    return context.view?.showHeadings ?? false;
  },
  run() {
    return { changes: [] };
  },
};

const viewZoom100Command: CommandDefinition<void> = {
  id: 'view.zoom100',
  label: 'Zoom 100%',
  category: 'view',
  keybindings: [],
  isEnabled(context) {
    return context.workbook.sheets.length > 0;
  },
  run() {
    return { changes: [] };
  },
};

const nameBoxNavigateCommand: CommandDefinition<void> = {
  id: 'nameBox.navigate',
  label: 'Go to address',
  category: 'view',
  keybindings: [],
  isEnabled(context) {
    return context.workbook.sheets.length > 0 && context.activeSheetId !== null;
  },
  run() {
    return { changes: [] };
  },
};

const helpAboutCommand: CommandDefinition<void> = {
  id: 'help.about',
  label: 'About calcplus',
  category: 'help',
  keybindings: [],
  isEnabled() {
    return true;
  },
  run() {
    return { changes: [] };
  },
};

const fileCompatibilityReportCommand: CommandDefinition<void> = {
  id: 'file.compatibilityReport',
  label: 'Compatibility Report',
  category: 'file',
  keybindings: [],
  isEnabled(context) {
    return context.workbook.sheets.length > 0;
  },
  run() {
    return { changes: [] };
  },
};

const fileExportReportCommand: CommandDefinition<void> = {
  id: 'file.exportReport',
  label: 'Export Report',
  category: 'file',
  keybindings: [],
  isEnabled(context) {
    return context.workbook.sheets.length > 0;
  },
  run() {
    return { changes: [] };
  },
};

/**
 * Persistence §3: the recent list is a chrome-owned dialog. The command is
 * always enabled; the dialog reports an empty list rather than hiding the File
 * menu entry, which is how Excel's Recent files behaves.
 */
const fileRecentCommand: CommandDefinition<void> = {
  id: 'file.recent',
  label: 'Recent files',
  category: 'file',
  keybindings: [],
  isEnabled() {
    return true;
  },
  run() {
    return { changes: [] };
  },
};

/**
 * Persistence §2: the recovery prompt is chrome-owned. It is not in a menu;
 * the host offers a copy on startup and exposes the command for a manual
 * check of the pending list.
 */
const fileRecoverCommand: CommandDefinition<void> = {
  id: 'file.recover',
  label: 'Recover unsaved work',
  category: 'file',
  keybindings: [],
  isEnabled() {
    return true;
  },
  run() {
    return { changes: [] };
  },
};

/**
 * The formula autocomplete surface is chrome-owned: the editor and the formula
 * bar render the suggestion list themselves from the published catalogue. The
 * registry entry keeps the command addressable and enablement has one source.
 */
const formulaAutocompleteCommand: CommandDefinition<void> = {
  id: 'formula.autocomplete',
  label: 'Function autocomplete and signature hints',
  category: 'formula',
  keybindings: [],
  isEnabled(context) {
    return context.workbook.sheets.length > 0;
  },
  run() {
    return { changes: [] };
  },
};

/**
 * File drop (feature matrix §7; ticket 49). The drop handler reads the file
 * and dispatches this ID through the command host, exactly as File → Open
 * dispatches `file.open`; the shell owns the prompts and the open path. The
 * supported extensions and the refusal message live here so the rule has one
 * source.
 */
const fileOpenDroppedCommand: CommandDefinition<void> = {
  id: 'file.openDropped',
  label: 'Open dropped file',
  category: 'file',
  keybindings: [],
  isEnabled(context) {
    return context.busy !== true;
  },
  run() {
    return { changes: [] };
  },
};

/** The v1 dropped-file types: `.xlsx`, `.xlsm` and `.csv` (ticket 49). */
const DROPPED_FILE_EXTENSIONS = ['.xlsx', '.xlsm', '.csv'] as const;

export const UNSUPPORTED_DROP_MESSAGE =
  'calcplus opens dropped files that are Excel workbooks (`.xlsx`, `.xlsm`) or CSV (`.csv`). Nothing was opened.';

/**
 * The refusal message for a dropped file calcplus cannot open, or `null` when
 * the file is supported. An extension-less or oddly named file is refused too:
 * v1 does not sniff contents.
 */
export function droppedFileProblem(name: string): string | null {
  const lower = name.toLowerCase();
  return DROPPED_FILE_EXTENSIONS.some((extension) => lower.endsWith(extension))
    ? null
    : UNSUPPORTED_DROP_MESSAGE;
}

/**
 * The result File → Export PDF reports to the status bar.
 *
 * `fallbackPages` names the pages that were rasterised at 300 DPI because
 * their text could not embed (print-pdf §2, §7). An empty list is the normal
 * vector export.
 */
export interface PdfExportSummary {
  readonly fileName: string;
  readonly pageCount: number;
  readonly fallbackPages: readonly number[];
}

/**
 * File → Export PDF: the plan becomes a PDF download.
 *
 * The plan is the single source of truth, built by the shell from the same
 * display records the grid and the preview use (print-pdf §2), so the PDF and
 * the preview cannot disagree. PDFKit and the canvas fallback are imported on
 * demand: the ~592 KB PDFKit chunk must stay out of the startup bundle
 * (research/26 §7.1).
 *
 * Every page is vector unless `preflightPrintPlan` routed it to the 300 DPI
 * fallback. A fallback page is rasterised from the same page, never redrawn
 * differently, and the download carries the plan's page boxes.
 */
export async function exportPdfPlan(
  plan: PrintPlan,
  workbookFileName: string | null,
): Promise<PdfExportSummary> {
  const [{ renderPdf }, { createCanvasRasterizer }] = await Promise.all([
    import('../print/pdf-renderer'),
    import('../print/canvas-fallback'),
  ]);
  const result = await renderPdf(plan, { rasterize: createCanvasRasterizer() });
  const fileName = pdfFileName(workbookFileName, plan.sheetName);
  downloadBytes(result.bytes, fileName, PDF_MIME);
  return {
    fileName,
    pageCount: result.pageCount,
    fallbackPages: result.fallbackPages,
  };
}

const PDF_MIME = 'application/pdf';

/**
 * The download name: the workbook's stem with a `.pdf` extension, or the sheet
 * name for a workbook that has never been named.
 */
export function pdfFileName(workbookFileName: string | null, sheetName: string): string {
  const name = workbookFileName?.trim() ?? '';
  const stem = name.replace(/\.(xlsx|xlsm)$/i, '');
  return `${stem === '' ? sheetName : stem}.pdf`;
}

export function chromeCommands(): readonly CommandDefinition<never>[] {
  return [
    fileNewCommand,
    fileOpenCommand,
    fileSaveCommand,
    fileSaveAsCommand,
    fileCloseCommand,
    filePrintCommand,
    fileExportPdfCommand,
    viewGridlinesCommand,
    viewHeadingsCommand,
    viewZoom100Command,
    nameBoxNavigateCommand,
    fileCompatibilityReportCommand,
    fileExportReportCommand,
    fileRecentCommand,
    fileRecoverCommand,
    fileOpenDroppedCommand,
    formulaAutocompleteCommand,
    helpAboutCommand,
  ] as unknown as readonly CommandDefinition<never>[];
}
