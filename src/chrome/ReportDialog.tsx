import { useState } from 'react';
import type { JSX } from 'react';
import { Dialog } from './Dialog';
import type { DialogCloseReason } from './Dialog';
import type { ReportSectionView } from './report-view';

export interface ReportDialogProps {
  readonly title: string;
  readonly summary: string;
  readonly sections: readonly ReportSectionView[];
  readonly notes?: readonly string[];
  /** Shown instead of the section list when there is nothing to list. */
  readonly emptyMessage: string;
  /** The text the Copy button puts on the clipboard. */
  readonly copyText: string;
  readonly onClose: (reason: DialogCloseReason) => void;
}

/**
 * The Compatibility Report and Export Report panel (file-compat §7,
 * ui-chrome §4.1). Read-only: it lists what calcplus preserved, transformed,
 * left stale, removed or warns about, and offers the text form on the
 * clipboard. It never mutates the model and never contacts a network.
 */
export function ReportDialog({
  title,
  summary,
  sections,
  notes = [],
  emptyMessage,
  copyText,
  onClose,
}: ReportDialogProps): JSX.Element {
  const [copied, setCopied] = useState<'idle' | 'copied' | 'failed'>('idle');

  const copy = (): void => {
    const clipboard = typeof navigator === 'undefined' ? undefined : navigator.clipboard;
    if (clipboard === undefined) {
      setCopied('failed');
      return;
    }
    void clipboard.writeText(copyText).then(
      () => setCopied('copied'),
      () => setCopied('failed'),
    );
  };

  return (
    <Dialog
      title={title}
      onClose={onClose}
      defaultAction={{ label: 'Close', onSelect: () => onClose('dismiss') }}
      className="report-dialog"
    >
      <p className="dialog-hint report-summary" data-testid="report-summary">
        {summary}
      </p>

      {sections.length === 0 ? (
        <p className="dialog-hint report-empty" data-testid="report-empty">
          {emptyMessage}
        </p>
      ) : (
        <div className="report-sections" data-testid="report-sections">
          {sections.map((section) => (
            <section key={section.id} className="report-section">
              <h3 className="report-section-title">{section.title}</h3>
              <ul className="report-list">
                {section.entries.map((entry) => (
                  <li
                    key={entry.id}
                    className="report-entry"
                    data-testid={`report-entry-${entry.id}`}
                  >
                    <span className="report-entry-title">{entry.title}</span>
                    {entry.where === null ? null : (
                      <span className="report-entry-where"> — {entry.where}</span>
                    )}
                    <p className="report-entry-detail">{entry.detail}</p>
                  </li>
                ))}
              </ul>
            </section>
          ))}
        </div>
      )}

      {notes.length === 0 ? null : (
        <section className="report-notes">
          <h3 className="report-section-title">Notes</h3>
          <ul className="report-note-list">
            {notes.map((note) => (
              <li key={note}>{note}</li>
            ))}
          </ul>
        </section>
      )}

      <div className="report-actions">
        <button type="button" className="dialog-button" data-testid="report-copy" onClick={copy}>
          Copy as text
        </button>
        {copied === 'copied' ? (
          <span className="report-copied" role="status">
            Copied.
          </span>
        ) : null}
        {copied === 'failed' ? (
          <span className="report-copied" role="status">
            The clipboard is unavailable.
          </span>
        ) : null}
      </div>
    </Dialog>
  );
}
