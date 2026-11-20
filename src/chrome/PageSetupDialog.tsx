import { useId, useMemo, useState } from 'react';
import type { JSX } from 'react';
import type { PageSetupPayload, PageSetupValues } from '@commands/print';
import { Dialog, DialogTabs } from './Dialog';
import type { DialogCloseReason, DialogTab } from './Dialog';
import {
  firstPageSetupError,
  hasPageSetupErrors,
  pageSetupDraftFrom,
  pageSetupDraftPayload,
  pageSetupPlanText,
  pageSetupValuesFromDraft,
} from './page-setup-state';
import type { PageSetupDraft, PageSetupDraftErrors, PageSetupPlanCheck } from './page-setup-state';
import { printPaperOptions as allPaperOptions } from './print-preview-state';
import type { PrintPaperOption } from './print-preview-state';

export interface PageSetupOutcome {
  readonly ok: boolean;
  readonly message: string | null;
}

export interface PageSetupDialogProps {
  /** The active sheet's current page setup; the draft starts here. */
  readonly initial: PageSetupValues;
  /** The active sheet's name, for the print-area hint and validation. */
  readonly sheetName: string;
  readonly onClose: (reason: DialogCloseReason) => void;
  /**
   * Writes the changed fields as one `print.applySetup` transaction. The shell
   * owns the dispatch, so the dialog itself never touches the model.
   */
  readonly onApply: (payload: PageSetupPayload) => PageSetupOutcome;
  /** Re-plans the draft, so the page count follows every edit (print-pdf §2). */
  readonly onPlanPreview?: (payload: PageSetupPayload) => PageSetupPlanCheck | null;
  /** The planner's media table; injected for tests. */
  readonly paperOptions?: readonly PrintPaperOption[];
}

const TABS: readonly DialogTab[] = [
  { id: 'page', label: 'Page' },
  { id: 'margins', label: 'Margins' },
  { id: 'sheet', label: 'Sheet' },
];

/**
 * Page Setup (print-pdf §4, ticket 44).
 *
 * The dialog owns the draft and its validation only. Page carries the
 * orientation, the scale or fit-to option and the paper size; Margins carries
 * the six inch values; Sheet carries the print area and the printed gridlines
 * and headings. One OK or Apply submits every changed field as one
 * `print.applySetup` transaction, and the live plan line re-plans the draft
 * through the shell's planner, so the page count follows the edit.
 *
 * The print-area box is the `_xlnm.Print_Area` system name. It is the one
 * authored defined name in v1 (print-pdf §4: a system name, not defined-name
 * authoring), and it travels in the same transaction as the rest of the draft.
 */
export function PageSetupDialog({
  initial,
  sheetName,
  onClose,
  onApply,
  onPlanPreview,
  paperOptions,
}: PageSetupDialogProps): JSX.Element {
  const [draft, setDraft] = useState<PageSetupDraft>(() => pageSetupDraftFrom(initial));
  const [committed, setCommitted] = useState<PageSetupValues>(initial);
  const [activeTab, setActiveTab] = useState('page');
  const [errors, setErrors] = useState<PageSetupDraftErrors>({});
  const [message, setMessage] = useState<string | null>(null);

  const ids = {
    scale: useId(),
    fitWidth: useId(),
    fitHeight: useId(),
    paper: useId(),
    printArea: useId(),
    top: useId(),
    right: useId(),
    bottom: useId(),
    left: useId(),
    header: useId(),
    footer: useId(),
  };

  const papers = paperOptions ?? allPaperOptions();
  const result = pageSetupDraftPayload(draft, committed, sheetName);
  const plan = useMemo(() => {
    if (onPlanPreview === undefined || hasPageSetupErrors(result.errors)) return null;
    return onPlanPreview(result.payload);
  }, [onPlanPreview, result]);

  const update = (patch: Partial<PageSetupDraft>): void => {
    setDraft((current) => ({ ...current, ...patch }));
    setErrors({});
    setMessage(null);
  };

  const updateMargin = (edge: keyof PageSetupDraft['margins'], value: string): void => {
    setDraft((current) => ({ ...current, margins: { ...current.margins, [edge]: value } }));
    setErrors({});
    setMessage(null);
  };

  const submit = (closeAfter: boolean): void => {
    const submitted = pageSetupDraftPayload(draft, committed, sheetName);
    if (hasPageSetupErrors(submitted.errors)) {
      setErrors(submitted.errors);
      setMessage(null);
      return;
    }
    setErrors({});
    if (submitted.payload.setup === undefined && submitted.payload.printArea === undefined) {
      if (closeAfter) onClose('dismiss');
      else setMessage('No page setup change.');
      return;
    }
    const outcome = onApply(submitted.payload);
    if (!outcome.ok) {
      setMessage(outcome.message ?? 'The page setup could not be applied.');
      return;
    }
    const next = pageSetupValuesFromDraft(draft, sheetName);
    if (next !== null) setCommitted(next);
    if (closeAfter) {
      onClose('dismiss');
      return;
    }
    setMessage(outcome.message ?? 'Page setup applied.');
  };

  const fieldError = (error: string | undefined, testId: string): JSX.Element | null =>
    error === undefined ? null : (
      <p className="page-setup-error" role="alert" data-testid={testId}>
        {error}
      </p>
    );

  return (
    <Dialog
      title="Page Setup"
      onClose={onClose}
      defaultAction={{ label: 'OK', onSelect: () => submit(true) }}
      secondaryAction={{ label: 'Cancel', onSelect: () => onClose('dismiss') }}
      className="page-setup-dialog"
    >
      <DialogTabs tabs={TABS} activeId={activeTab} onChange={setActiveTab} />

      {activeTab === 'page' ? (
        <div className="page-setup-panel" role="tabpanel" aria-label="Page options">
          <fieldset className="page-setup-group">
            <legend>Orientation</legend>
            <label className="page-setup-radio">
              <input
                type="radio"
                name="page-setup-orientation"
                value="portrait"
                checked={draft.orientation === 'portrait'}
                data-testid="page-setup-portrait"
                onChange={() => update({ orientation: 'portrait' })}
              />
              Portrait
            </label>
            <label className="page-setup-radio">
              <input
                type="radio"
                name="page-setup-orientation"
                value="landscape"
                checked={draft.orientation === 'landscape'}
                data-testid="page-setup-landscape"
                onChange={() => update({ orientation: 'landscape' })}
              />
              Landscape
            </label>
          </fieldset>

          <fieldset className="page-setup-group">
            <legend>Scaling</legend>
            <div className="page-setup-scale-row">
              <label className="page-setup-radio">
                <input
                  type="radio"
                  name="page-setup-scaling"
                  value="scale"
                  checked={draft.scaling === 'scale'}
                  data-testid="page-setup-mode-scale"
                  onChange={() => update({ scaling: 'scale' })}
                />
                Adjust to
              </label>
              <input
                id={ids.scale}
                type="number"
                className="find-input page-setup-number"
                min={10}
                max={400}
                step={1}
                aria-label="Scale percentage"
                data-testid="page-setup-scale"
                disabled={draft.scaling !== 'scale'}
                value={draft.scale}
                onChange={(event) => update({ scale: event.target.value })}
              />
              <span>% normal size</span>
            </div>
            {fieldError(errors.scale, 'page-setup-scale-error')}
            <div className="page-setup-scale-row">
              <label className="page-setup-radio">
                <input
                  type="radio"
                  name="page-setup-scaling"
                  value="fit"
                  checked={draft.scaling === 'fit'}
                  data-testid="page-setup-mode-fit"
                  onChange={() => update({ scaling: 'fit' })}
                />
                Fit to
              </label>
              <input
                id={ids.fitWidth}
                type="number"
                className="find-input page-setup-number"
                min={0}
                max={32767}
                step={1}
                aria-label="Fit to width in pages"
                data-testid="page-setup-fit-width"
                disabled={draft.scaling !== 'fit'}
                value={draft.fitToWidth}
                onChange={(event) => update({ fitToWidth: event.target.value })}
              />
              <span>page(s) wide by</span>
              <input
                id={ids.fitHeight}
                type="number"
                className="find-input page-setup-number"
                min={0}
                max={32767}
                step={1}
                aria-label="Fit to height in pages"
                data-testid="page-setup-fit-height"
                disabled={draft.scaling !== 'fit'}
                value={draft.fitToHeight}
                onChange={(event) => update({ fitToHeight: event.target.value })}
              />
              <span>tall</span>
            </div>
            {fieldError(errors.fitToWidth, 'page-setup-fit-width-error')}
            {fieldError(errors.fitToHeight, 'page-setup-fit-height-error')}
          </fieldset>

          <div className="page-setup-row">
            <label className="find-label" htmlFor={ids.paper}>
              Paper size
            </label>
            <select
              id={ids.paper}
              className="find-select"
              data-testid="page-setup-paper"
              value={draft.paperSize}
              onChange={(event) => update({ paperSize: event.target.value })}
            >
              {papers.some((paper) => String(paper.code) === draft.paperSize) ||
              draft.paperSize === '' ? null : (
                <option value={draft.paperSize}>Custom (code {draft.paperSize})</option>
              )}
              {papers.map((paper) => (
                <option key={paper.code} value={String(paper.code)}>
                  {paper.name}
                </option>
              ))}
            </select>
          </div>
          {fieldError(errors.paperSize, 'page-setup-paper-error')}
        </div>
      ) : null}

      {activeTab === 'margins' ? (
        <div className="page-setup-panel" role="tabpanel" aria-label="Margin options">
          <p className="dialog-hint">Margins are in inches, like the file stores them.</p>
          <div className="page-setup-margins">
            {(
              [
                ['top', 'Top', ids.top],
                ['bottom', 'Bottom', ids.bottom],
                ['left', 'Left', ids.left],
                ['right', 'Right', ids.right],
                ['header', 'Header', ids.header],
                ['footer', 'Footer', ids.footer],
              ] as const
            ).map(([edge, label, id]) => (
              <div className="page-setup-margin" key={edge}>
                <label className="find-label" htmlFor={id}>
                  {label}
                </label>
                <input
                  id={id}
                  type="number"
                  className="find-input page-setup-number"
                  min={0}
                  max={49}
                  step={0.05}
                  data-testid={`page-setup-margin-${edge}`}
                  value={draft.margins[edge]}
                  onChange={(event) => updateMargin(edge, event.target.value)}
                />
              </div>
            ))}
          </div>
          {fieldError(errors.margins, 'page-setup-margins-error')}
          <fieldset className="page-setup-group">
            <legend>Centre on page</legend>
            <label className="page-setup-check">
              <input
                type="checkbox"
                data-testid="page-setup-centre-horizontal"
                checked={draft.horizontalCentered}
                onChange={(event) => update({ horizontalCentered: event.target.checked })}
              />
              Horizontally
            </label>
            <label className="page-setup-check">
              <input
                type="checkbox"
                data-testid="page-setup-centre-vertical"
                checked={draft.verticalCentered}
                onChange={(event) => update({ verticalCentered: event.target.checked })}
              />
              Vertically
            </label>
          </fieldset>
        </div>
      ) : null}

      {activeTab === 'sheet' ? (
        <div className="page-setup-panel" role="tabpanel" aria-label="Sheet options">
          <div className="page-setup-row">
            <label className="find-label" htmlFor={ids.printArea}>
              Print area
            </label>
            <input
              id={ids.printArea}
              type="text"
              className="find-input"
              placeholder={`For example A1:C5 on ${sheetName}`}
              data-testid="page-setup-print-area"
              value={draft.printArea}
              onChange={(event) => update({ printArea: event.target.value })}
            />
          </div>
          {fieldError(errors.printArea, 'page-setup-print-area-error')}
          <p className="dialog-hint">
            Ranges on this sheet, separated by commas. Leave the box blank to print the used range.
            The Print Area drop-down on the Page Layout tab sets it from the selection.
          </p>
          <fieldset className="page-setup-group">
            <legend>Print</legend>
            <label className="page-setup-check">
              <input
                type="checkbox"
                data-testid="page-setup-print-gridlines"
                checked={draft.printGridlines}
                onChange={(event) => update({ printGridlines: event.target.checked })}
              />
              Gridlines
            </label>
            <label className="page-setup-check">
              <input
                type="checkbox"
                data-testid="page-setup-print-headings"
                checked={draft.printHeadings}
                onChange={(event) => update({ printHeadings: event.target.checked })}
              />
              Row and column headings
            </label>
          </fieldset>
        </div>
      ) : null}

      <p className="dialog-hint page-setup-plan" data-testid="page-setup-plan" role="status">
        {plan === null ? '' : pageSetupPlanText(plan)}
      </p>
      {message === null ? null : (
        <p className="dialog-message" role="status" data-testid="page-setup-message">
          {message}
        </p>
      )}
      {Object.keys(errors).length === 0 ? null : (
        <p className="dialog-error" role="alert" data-testid="page-setup-error">
          {firstPageSetupError(errors)}
        </p>
      )}

      <div className="format-actions">
        <button
          type="button"
          className="dialog-button"
          data-testid="page-setup-apply"
          onClick={() => submit(false)}
        >
          Apply
        </button>
      </div>
    </Dialog>
  );
}
