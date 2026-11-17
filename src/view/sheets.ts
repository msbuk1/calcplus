import type { Sheet, Workbook } from '@model/types';
import { isSheetVisible } from './sheet-visibility';

/**
 * Chrome-facing sheet metadata.
 *
 * The tab strip and the status bar read plain records from this module, so
 * they never import the workbook model. The projection owns the cells; this
 * module owns only the ordered tab list and each sheet's view defaults.
 */
export interface SheetTabInfo {
  readonly id: string;
  readonly name: string;
}

/** Live view state for one sheet: never written to the workbook. */
export interface SheetViewState {
  /** Display zoom as an Excel percentage, 10-400. */
  readonly zoom: number;
  readonly showGridLines: boolean;
  readonly showHeadings: boolean;
}

export const DEFAULT_SHEET_VIEW: SheetViewState = Object.freeze({
  zoom: 100,
  showGridLines: true,
  showHeadings: true,
});

export function clampZoomPercent(value: number): number {
  if (!Number.isFinite(value)) return 100;
  return Math.min(400, Math.max(10, Math.round(value)));
}

/**
 * The tab order is the workbook's ordered visible sheet list, never an ID
 * order. Imported hidden and veryHidden sheets stay in the workbook but are
 * not offered as ordinary editable tabs; hide/unhide authoring is V2.
 */
export function sheetTabList(workbook: Workbook): SheetTabInfo[] {
  return workbook.sheets
    .filter((sheet) => isSheetVisible(workbook, sheet.id))
    .map((sheet) => ({ id: sheet.id, name: sheet.name }));
}

/** The imported view settings, used to seed the live view state on open. */
export function viewForSheet(sheet: Sheet): SheetViewState {
  return {
    zoom: clampZoomPercent(sheet.view.zoom),
    showGridLines: sheet.view.showGridLines,
    showHeadings: sheet.view.showHeadings,
  };
}

export function initialSheetViews(workbook: Workbook): Record<string, SheetViewState> {
  const views: Record<string, SheetViewState> = {};
  for (const sheet of workbook.sheets) views[sheet.id] = viewForSheet(sheet);
  return views;
}
