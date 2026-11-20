import type { JSX } from 'react';
import type { RecentEntry } from '@io/recent';
import { Dialog } from './Dialog';
import type { DialogCloseReason } from './Dialog';

export interface RecentDialogProps {
  readonly entries: readonly RecentEntry[];
  readonly onOpen: (entry: RecentEntry) => void;
  readonly onRemove: (entry: RecentEntry) => void;
  readonly onClose: (reason: DialogCloseReason) => void;
}

/**
 * File → Recent files (persistence §3, ui-chrome §8). The list shows a display
 * name and an origin label only; it never shows an absolute path. Reopening an
 * entry asks the host to re-request permission and re-read the file.
 */
export function RecentDialog({
  entries,
  onOpen,
  onRemove,
  onClose,
}: RecentDialogProps): JSX.Element {
  return (
    <Dialog
      title="Recent files"
      onClose={onClose}
      defaultAction={{ label: 'Close', onSelect: () => onClose('dismiss') }}
      className="recent-dialog"
    >
      {entries.length === 0 ? (
        <p className="dialog-hint" data-testid="recent-empty">
          There are no recent files on this origin yet.
        </p>
      ) : (
        <ul className="recent-list" data-testid="recent-list">
          {entries.map((entry) => (
            <li key={entry.id} className="recent-entry" data-testid={`recent-entry-${entry.id}`}>
              <button
                type="button"
                className="recent-open"
                data-testid={`recent-open-${entry.id}`}
                onClick={() => onOpen(entry)}
              >
                <span className="recent-name">{entry.name}</span>
                <span className="recent-detail">
                  {entry.handle === null ? 'Name only' : 'Saved copy'} · last opened{' '}
                  {new Date(entry.lastOpenedAt).toLocaleString()}
                </span>
              </button>
              <button
                type="button"
                className="recent-remove"
                data-testid={`recent-remove-${entry.id}`}
                aria-label={`Remove ${entry.name} from the recent list`}
                onClick={() => onRemove(entry)}
              >
                Remove
              </button>
            </li>
          ))}
        </ul>
      )}
    </Dialog>
  );
}
