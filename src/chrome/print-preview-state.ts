import type { PrintPlan } from '../print/types';

/**
 * View state for the print preview window (print-pdf §2, ui-chrome §4.1).
 *
 * Pure helpers only: the dialog owns React state, and these functions own the
 * keyboard map and the one-line summary. Keeping them here means the keyboard
 * contract is unit-tested without a browser.
 */

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
