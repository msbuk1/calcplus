import { useCallback, useEffect, useMemo, useState } from 'react';
import type { JSX } from 'react';
import type { PrintPlan } from '../print/types';
import { renderPrintPlan } from '../print/svg-renderer';
import type { RenderedPrintDocument } from '../print/svg-renderer';
import { Dialog } from './Dialog';
import type { DialogCloseReason } from './Dialog';
import {
  printPreviewKeyAction,
  printPreviewPageText,
  printPreviewSummary,
} from './print-preview-state';

export interface PrintPreviewDialogProps {
  readonly plan: PrintPlan;
  readonly onClose: (reason: DialogCloseReason) => void;
  /** Injected by tests; the window's own print dialog by default. */
  readonly onPrint?: () => void;
}

/**
 * The print preview window (print-pdf §2, ui-chrome §4.1).
 *
 * It renders from the frozen `PrintPlan` only: one SVG per page, in page
 * points, so what the preview shows is what `window.print()` lays on paper and
 * what the PDF export writes. Page navigation is keyboard-first (arrows, Page
 * Up/Down, Home/End) and announced through a live region; the SVG itself is a
 * labelled image ("Page 1 of 3, Sheet1").
 *
 * The print-only stack is a sibling of the dialog, not part of it: under
 * `@media print` the shell hides everything except that stack, so the browser
 * prints the plan's pages instead of the chrome. Its pages are sized to the
 * browser's real page box (spike §1), which keeps N planned pages at N printed
 * pages.
 */
export function PrintPreviewDialog({
  plan,
  onClose,
  onPrint,
}: PrintPreviewDialogProps): JSX.Element {
  const rendered = useMemo(() => renderPrintPlan(plan), [plan]);
  const pageCount = rendered.pageCount;
  const [page, setPage] = useState(0);
  const current = rendered.pages[Math.min(page, Math.max(0, pageCount - 1))];

  const print = useCallback((): void => {
    if (pageCount === 0) return;
    if (onPrint !== undefined) onPrint();
    else if (typeof window !== 'undefined') window.print();
  }, [onPrint, pageCount]);

  // The preview owns the paging keys wherever focus sits in the modal, so the
  // keys work from the Print button as well as from the preview itself.
  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent): void => {
      if (event.defaultPrevented) return;
      const target = event.target;
      if (
        target instanceof HTMLElement &&
        (target.isContentEditable || /^(INPUT|TEXTAREA|SELECT)$/.test(target.tagName))
      ) {
        return;
      }
      const action = printPreviewKeyAction(
        {
          key: event.key,
          ctrlKey: event.ctrlKey,
          metaKey: event.metaKey,
          altKey: event.altKey,
          shiftKey: event.shiftKey,
        },
        page,
        pageCount,
      );
      if (action === null) return;
      event.preventDefault();
      if (action.type === 'print') print();
      else setPage(action.page);
    };
    document.addEventListener('keydown', onKeyDown);
    return () => document.removeEventListener('keydown', onKeyDown);
  }, [page, pageCount, print]);

  return (
    <>
      <Dialog
        title="Print Preview"
        className="print-preview-dialog"
        onClose={onClose}
        defaultAction={pageCount > 0 ? { label: 'Print', onSelect: print } : undefined}
        secondaryAction={{ label: 'Close', onSelect: () => onClose('dismiss') }}
      >
        <div className="print-preview" data-testid="print-preview">
          <p className="dialog-hint print-preview-summary" data-testid="print-preview-summary">
            {printPreviewSummary(plan)}
          </p>
          {current === undefined ? (
            <p className="dialog-hint print-preview-empty" data-testid="print-preview-empty">
              There is nothing to print. The sheet has no used cells and no print area.
            </p>
          ) : (
            <>
              <div className="print-preview-stage">
                <div
                  className="print-preview-page"
                  data-testid="print-preview-page"
                  style={{ aspectRatio: `${current.box.width} / ${current.box.height}` }}
                  // The markup is built by `renderPrintPlan`, which escapes every
                  // sheet-sourced string and validates every colour.
                  dangerouslySetInnerHTML={{ __html: current.svg }}
                />
              </div>
              <div className="print-preview-footer">
                <button
                  type="button"
                  className="dialog-button"
                  data-testid="print-preview-previous"
                  aria-label="Previous page"
                  disabled={page <= 0}
                  onClick={() => setPage((value) => Math.max(0, value - 1))}
                >
                  Previous
                </button>
                <p
                  className="print-preview-count"
                  data-testid="print-preview-count"
                  role="status"
                  aria-live="polite"
                  aria-atomic="true"
                >
                  {printPreviewPageText(page, pageCount)}
                </p>
                <button
                  type="button"
                  className="dialog-button"
                  data-testid="print-preview-next"
                  aria-label="Next page"
                  disabled={page >= pageCount - 1}
                  onClick={() => setPage((value) => Math.min(pageCount - 1, value + 1))}
                >
                  Next
                </button>
              </div>
              <p className="dialog-hint print-preview-hint">
                Arrow keys and Page Up/Down change the page. The browser print dialog controls the
                printer, the paper and whether background graphics print.
              </p>
            </>
          )}
        </div>
      </Dialog>
      <PrintPages rendered={rendered} />
    </>
  );
}

/**
 * The print-only stack: one page-sized element per plan page. It is hidden on
 * screen and is the only thing the shell leaves visible under `@media print`.
 */
export function PrintPages({
  rendered,
}: {
  readonly rendered: RenderedPrintDocument;
}): JSX.Element {
  const { paper } = rendered;
  return (
    <div className="print-only-pages" data-testid="print-pages" aria-hidden="true">
      {/*
        The plan's margins are already page coordinates inside the SVG, so the
        page box itself has no CSS margin; only its size comes from the plan's
        paper and orientation. `@page` cannot read a custom property, so the
        rule is written here from the plan's own numbers.
      */}
      <style>{`@page { size: ${paper.width}pt ${paper.height}pt; margin: 0; }`}</style>
      {rendered.pages.map((page) => (
        <div
          key={page.index}
          className="print-page"
          data-testid="print-page"
          data-page-number={page.number}
          style={{ width: `${page.printBox.width}pt`, height: `${page.printBox.height}pt` }}
          dangerouslySetInnerHTML={{ __html: page.svg }}
        />
      ))}
    </div>
  );
}
