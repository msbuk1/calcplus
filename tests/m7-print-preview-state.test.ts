import { describe, expect, it } from 'vitest';
import {
  printPreviewKeyAction,
  printPreviewPageText,
  printPreviewSummary,
} from '../src/chrome/print-preview-state';
import type { PrintPaper, PrintPlan } from '../src/print/types';

/**
 * M7 lane 2 — the preview window's keyboard map and summary line.
 *
 * The dialog shell owns Escape and focus; this module owns the paging keys, so
 * the contract is unit-tested without a browser.
 */

const A4: PrintPaper = {
  name: 'A4',
  code: 9,
  orientation: 'landscape',
  width: 841.89,
  height: 595.28,
};

function planOf(pageCount: number, overrides: Partial<PrintPlan> = {}): PrintPlan {
  return {
    sheetId: 'sheet-1',
    sheetName: 'Sheet1',
    paper: A4,
    margins: { top: 54, right: 50.4, bottom: 54, left: 50.4 },
    scale: 0.85,
    fitToPage: true,
    showGridLines: true,
    showHeadings: true,
    horizontalCentered: false,
    verticalCentered: false,
    pageSets: [],
    pages: Array.from({ length: pageCount }, (_, index) => ({
      index,
      setIndex: 0,
      pageInSetIndex: index,
      paper: A4,
      box: { x: 0, y: 0, width: A4.width, height: A4.height },
      body: { x: 50.4, y: 54, width: 741.09, height: 487.28 },
      band: { startRow: 0, endRow: 0, startCol: 0, endCol: 0 },
      columns: [],
      rows: [],
      items: [],
      clipped: false,
    })),
    ...overrides,
  };
}

describe('the preview keyboard map', () => {
  it('moves one page with the arrows', () => {
    expect(printPreviewKeyAction({ key: 'ArrowRight' }, 0, 3)).toEqual({ type: 'page', page: 1 });
    expect(printPreviewKeyAction({ key: 'ArrowLeft' }, 1, 3)).toEqual({ type: 'page', page: 0 });
  });

  it('moves one page with Page Up and Page Down', () => {
    expect(printPreviewKeyAction({ key: 'PageDown' }, 0, 3)).toEqual({ type: 'page', page: 1 });
    expect(printPreviewKeyAction({ key: 'PageUp' }, 2, 3)).toEqual({ type: 'page', page: 1 });
  });

  it('jumps to the ends with Home and End', () => {
    expect(printPreviewKeyAction({ key: 'End' }, 0, 4)).toEqual({ type: 'page', page: 3 });
    expect(printPreviewKeyAction({ key: 'Home' }, 3, 4)).toEqual({ type: 'page', page: 0 });
  });

  it('clamps at the ends instead of wrapping', () => {
    expect(printPreviewKeyAction({ key: 'ArrowLeft' }, 0, 3)).toBeNull();
    expect(printPreviewKeyAction({ key: 'PageDown' }, 2, 3)).toBeNull();
    expect(printPreviewKeyAction({ key: 'Home' }, 0, 3)).toBeNull();
    expect(printPreviewKeyAction({ key: 'End' }, 2, 3)).toBeNull();
  });

  it('ignores keys it does not own, and any key with no pages', () => {
    expect(printPreviewKeyAction({ key: 'Tab' }, 0, 3)).toBeNull();
    expect(printPreviewKeyAction({ key: 'a' }, 0, 3)).toBeNull();
    expect(printPreviewKeyAction({ key: 'ArrowRight', shiftKey: true }, 0, 3)).toBeNull();
    expect(printPreviewKeyAction({ key: 'ArrowRight', altKey: true }, 0, 3)).toBeNull();
    expect(printPreviewKeyAction({ key: 'ArrowRight' }, 0, 0)).toBeNull();
  });

  it('prints on Ctrl+P and Meta+P', () => {
    expect(printPreviewKeyAction({ key: 'p', ctrlKey: true }, 0, 3)).toEqual({ type: 'print' });
    expect(printPreviewKeyAction({ key: 'P', metaKey: true }, 1, 3)).toEqual({ type: 'print' });
    expect(printPreviewKeyAction({ key: 'p', ctrlKey: true, shiftKey: true }, 0, 3)).toBeNull();
  });
});

describe('the summary line', () => {
  it('names the paper, orientation, scale and page count', () => {
    expect(printPreviewSummary(planOf(2))).toBe(
      'Sheet1 · A4 landscape · 85% (fit to page) · 2 pages',
    );
  });

  it('reads naturally for one page and a custom paper', () => {
    const plan = planOf(1, {
      scale: 1,
      fitToPage: false,
      paper: { name: 'Custom', code: 256, orientation: 'portrait', width: 400, height: 500 },
    });
    expect(printPreviewSummary(plan)).toBe('Sheet1 · Custom paper portrait · 100% · 1 page');
  });

  it('announces the page number without the sheet name', () => {
    expect(printPreviewPageText(0, 3)).toBe('Page 1 of 3');
    expect(printPreviewPageText(2, 3)).toBe('Page 3 of 3');
    expect(printPreviewPageText(9, 3)).toBe('Page 3 of 3');
  });
});
