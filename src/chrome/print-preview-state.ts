import { DEFAULT_FONT } from '@model/style';
import { createPrintPlan } from '../print/planner';
import type { PrintPlan, PrintSheetSnapshot } from '../print/types';
import { EMPTY_PRINT_SETUP } from '../commands/print';

/**
 * View state for the print preview window (print-pdf §2, ui-chrome §4.1).
 *
 * Pure helpers only: the dialog owns React state, and these functions own the
 * keyboard map, the one-line summary and the paper list. Keeping them here
 * means the contracts are unit-tested without a browser.
 */

/** One choice in the preview's media list: an OOXML paper code and its name. */
export interface PrintPaperOption {
  readonly code: number;
  readonly name: string;
}

/**
 * Ecma-376 `pageSetup@paperSize` codes run 1 … 118. The planner resolves each
 * code against its own table and names an unknown one `Custom`, so probing the
 * table through the public planner keeps the list and the table from drifting:
 * `src/print/planner.ts` owns the only copy.
 */
const HIGHEST_PAPER_CODE = 118;

const PAPER_PROBE: PrintSheetSnapshot = {
  sheetId: 'paper-list',
  sheetName: '',
  printAreas: [],
  usedRange: null,
  printSetup: null,
  rowHeight: () => 0,
  columnWidth: () => 0,
  columnLabel: () => '',
  rowLabel: () => '',
  headingFont: DEFAULT_FONT,
  cellAt: () => null,
  cellsIn: () => [],
  mergesIn: () => [],
};

let cachedPaperOptions: readonly PrintPaperOption[] | null = null;

/**
 * Every paper the planner knows, in code order. A plan with no ranges returns
 * before it reads any sheet geometry, so the probe is cheap; the list is built
 * once per session.
 */
export function printPaperOptions(): readonly PrintPaperOption[] {
  if (cachedPaperOptions !== null) return cachedPaperOptions;
  const options: PrintPaperOption[] = [];
  for (let code = 1; code <= HIGHEST_PAPER_CODE; code += 1) {
    const paper = createPrintPlan({
      ...PAPER_PROBE,
      printSetup: { ...EMPTY_PRINT_SETUP, paperSize: code },
    }).paper;
    // A code the table does not name resolves to `Custom`; it is not a choice.
    if (paper.name !== 'Custom') options.push({ code, name: paper.name });
  }
  cachedPaperOptions = options;
  return options;
}

/**
 * The media list's selected option for a plan. A sheet that declares no paper
 * uses the planner's default (Letter), and an unknown imported code has no
 * option, so the dialog shows a Custom entry instead.
 */
export function printSelectedPaperCode(
  plan: PrintPlan,
  options: readonly PrintPaperOption[],
): number | null {
  const exact = options.find((option) => option.code === plan.paper.code);
  if (exact !== undefined) return exact.code;
  return options.find((option) => option.name === plan.paper.name)?.code ?? null;
}

/** What a key press in the preview should do, or `null` to leave it alone. */
export type PrintPreviewKeyAction =
  { readonly type: 'page'; readonly page: number } | { readonly type: 'print' } | null;

export interface PrintPreviewKeyInput {
  readonly key: string;
  readonly ctrlKey?: boolean;
  readonly metaKey?: boolean;
  readonly altKey?: boolean;
  readonly shiftKey?: boolean;
}

/**
 * The preview's keyboard map: arrows and Page Up/Down move one page, Home and
 * End jump to the ends, and Ctrl+P prints through the same path as the button.
 * `page` is 0-based. A key that would not change the page returns `null`, so it
 * is not swallowed.
 */
export function printPreviewKeyAction(
  input: PrintPreviewKeyInput,
  page: number,
  pageCount: number,
): PrintPreviewKeyAction {
  if (pageCount <= 0) return null;
  const last = pageCount - 1;
  const current = clamp(Math.round(page), 0, last);
  if ((input.ctrlKey === true || input.metaKey === true) && !input.altKey && !input.shiftKey) {
    return input.key === 'p' || input.key === 'P' ? { type: 'print' } : null;
  }
  if (input.altKey === true || input.shiftKey === true) return null;
  switch (input.key) {
    case 'ArrowLeft':
    case 'PageUp':
      return current > 0 ? { type: 'page', page: current - 1 } : null;
    case 'ArrowRight':
    case 'PageDown':
      return current < last ? { type: 'page', page: current + 1 } : null;
    case 'Home':
      return current !== 0 ? { type: 'page', page: 0 } : null;
    case 'End':
      return current !== last ? { type: 'page', page: last } : null;
    default:
      return null;
  }
}

/**
 * The paper, scale and page count line above the preview. It names what the
 * plan resolved, not what the view shows: orientation, the uniform scale, the
 * fit-to-page setting and the number of pages.
 */
export function printPreviewSummary(plan: PrintPlan): string {
  const { paper } = plan;
  const parts = [
    plan.sheetName,
    `${paper.name === 'Custom' ? 'Custom paper' : paper.name} ${paper.orientation}`,
    `${Math.round(plan.scale * 100)}%${plan.fitToPage ? ' (fit to page)' : ''}`,
  ];
  const count = plan.pages.length;
  parts.push(count === 1 ? '1 page' : `${count} pages`);
  return parts.join(' · ');
}

/** "Page 2 of 3" for the live region that announces a page change. */
export function printPreviewPageText(page: number, pageCount: number): string {
  const current = clamp(page, 0, Math.max(0, pageCount - 1)) + 1;
  return `Page ${current} of ${pageCount}`;
}

function clamp(value: number, low: number, high: number): number {
  if (!Number.isFinite(value)) return low;
  return Math.min(high, Math.max(low, value));
}
