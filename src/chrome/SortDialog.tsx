import { useState } from 'react';
import type { JSX } from 'react';
import { MAX_SORT_KEYS } from '@model/sort';
import type { SortKey } from '@model/sort';
import { Dialog } from './Dialog';
import type { DialogCloseReason } from './Dialog';
import {
  SORT_DIRECTION_OPTIONS,
  SORT_ON_LABEL,
  addSortLevel,
  copySortLevel,
  createSortDraft,
  deleteSortLevel,
  moveSortLevel,
  setSortHeaders,
  sortColumnLabel,
  sortDraftProblem,
  sortKeys,
  sortLevelLabel,
  updateSortLevel,
} from './sort-dialog-state';
import type { SortDialogColumn, SortDraft } from './sort-dialog-state';

export interface SortSubmit {
  readonly keys: readonly SortKey[];
  readonly hasHeaders: boolean;
}

export interface SortDialogOutcome {
  readonly ok: boolean;
  readonly message: string | null;
}

export interface SortDialogProps {
  /** The target range as a read-only label (sort-autofilter §2.1). */
  readonly targetLabel: string;
  /** The columns of the target range, with the header text the command guessed. */
  readonly columns: readonly SortDialogColumn[];
  /** The command's header guess; the checkbox overrides it (§2.2). */
  readonly initialHasHeaders: boolean;
  readonly onClose: (reason: DialogCloseReason) => void;
  readonly onSubmit: (request: SortSubmit) => SortDialogOutcome;
}

/**
 * The Sort dialog (design-refs m5-menus-dialogs §1): title "Customised Sort",
 * a level toolbar, the "My data has headers" override, the key table and the
 * read-only target range. Blueprint scope only — colour and icon sorts, the
 * left-to-right orientation and case sensitivity are v2, so the Order row is
 * the only choice per level and the Options menu is omitted rather than
 * rendered empty (ui-chrome §4.1: out-of-scope entries are hidden).
 *
 * The dialog collects keys; `data.sort` owns the transaction, so confirming
 * here is one undo step and undo restores the exact prior rows (§2.7).
 */
export function SortDialog({
  targetLabel,
  columns,
  initialHasHeaders,
  onClose,
  onSubmit,
}: SortDialogProps): JSX.Element {
  const [draft, setDraft] = useState<SortDraft>(() => createSortDraft(columns, initialHasHeaders));
  const [selected, setSelected] = useState(0);
  const [error, setError] = useState<string | null>(null);
  const [dispatchError, setDispatchError] = useState<string | null>(null);

  const change = (next: SortDraft, select?: number): void => {
    setDraft(next);
    setError(null);
    setDispatchError(null);
    if (select !== undefined) {
      setSelected(Math.min(Math.max(select, 0), Math.max(0, next.levels.length - 1)));
    } else if (selected >= next.levels.length) {
      setSelected(Math.max(0, next.levels.length - 1));
    }
  };

  const submit = (): void => {
    const problem = sortDraftProblem(draft, columns);
    if (problem !== null) {
      setError(problem);
      setDispatchError(null);
      return;
    }
    const outcome = onSubmit({ keys: sortKeys(draft), hasHeaders: draft.hasHeaders });
    if (!outcome.ok) {
      setDispatchError(outcome.message ?? 'The sort could not be applied.');
      return;
    }
    onClose('dismiss');
  };

  const lastIndex = draft.levels.length - 1;

  return (
    <Dialog
      title="Customised Sort"
      onClose={onClose}
      defaultAction={{ label: 'OK', onSelect: submit }}
      secondaryAction={{ label: 'Cancel', onSelect: () => onClose('dismiss') }}
      className="sort-dialog"
    >
      <div className="sort-toolbar">
        <button
          type="button"
          className="sort-tool"
          data-testid="sort-add"
          disabled={draft.levels.length >= MAX_SORT_KEYS}
          onClick={() => change(addSortLevel(draft, columns), draft.levels.length)}
        >
          <span className="sort-tool-icon" aria-hidden="true">
            +
          </span>
          Add
        </button>
        <button
          type="button"
          className="sort-tool"
          data-testid="sort-delete"
          disabled={draft.levels.length <= 1}
          onClick={() => change(deleteSortLevel(draft, selected))}
        >
          Delete
        </button>
        <button
          type="button"
          className="sort-tool"
          data-testid="sort-copy"
          disabled={draft.levels.length >= MAX_SORT_KEYS}
          onClick={() => change(copySortLevel(draft, selected), selected + 1)}
        >
          Copy
        </button>
        <button
          type="button"
          className="sort-tool"
          aria-label="Move level up"
          data-testid="sort-up"
          disabled={selected <= 0}
          onClick={() => change(moveSortLevel(draft, selected, -1), selected - 1)}
        >
          <span aria-hidden="true">↑</span>
        </button>
        <button
          type="button"
          className="sort-tool"
          aria-label="Move level down"
          data-testid="sort-down"
          disabled={selected >= lastIndex}
          onClick={() => change(moveSortLevel(draft, selected, 1), selected + 1)}
        >
          <span aria-hidden="true">↓</span>
        </button>
        <label className="sort-headers">
          <input
            type="checkbox"
            checked={draft.hasHeaders}
            data-testid="sort-headers"
            onChange={(event) => change(setSortHeaders(draft, event.target.checked))}
          />
          My data has headers
        </label>
      </div>

      <table className="sort-table" aria-label="Sort keys" data-testid="sort-keys">
        <thead>
          <tr>
            <th scope="col" className="sort-level-head">
              <span className="visually-hidden">Level</span>
            </th>
            <th scope="col">Column</th>
            <th scope="col">Sort On</th>
            <th scope="col">Order</th>
          </tr>
        </thead>
        <tbody>
          {draft.levels.map((level, index) => (
            <tr
              key={index}
              className={index === selected ? 'sort-key-row sort-key-row-selected' : 'sort-key-row'}
              data-testid={`sort-key-row-${index}`}
              onClick={() => setSelected(index)}
            >
              <th scope="row" className="sort-level-label">
                {sortLevelLabel(index)}
              </th>
              <td>
                <select
                  className="find-select sort-control"
                  aria-label={`Sort column ${index + 1}`}
                  data-testid={`sort-key-${index}-column`}
                  value={String(level.column)}
                  onFocus={() => setSelected(index)}
                  onChange={(event) =>
                    change(
                      updateSortLevel(draft, index, { column: Number(event.target.value) }),
                      index,
                    )
                  }
                >
                  {columns.map((column) => (
                    <option key={column.column} value={String(column.column)}>
                      {sortColumnLabel(columns, column.column, draft.hasHeaders)}
                    </option>
                  ))}
                </select>
              </td>
              <td>
                <select
                  className="find-select sort-control"
                  aria-label={`Sort on ${index + 1}`}
                  data-testid={`sort-key-${index}-on`}
                  value="cell-values"
                  onFocus={() => setSelected(index)}
                  onChange={() => undefined}
                >
                  <option value="cell-values">{SORT_ON_LABEL}</option>
                </select>
              </td>
              <td>
                <select
                  className="find-select sort-control"
                  aria-label={`Order ${index + 1}`}
                  data-testid={`sort-key-${index}-order`}
                  value={level.direction}
                  onFocus={() => setSelected(index)}
                  onChange={(event) =>
                    change(
                      updateSortLevel(draft, index, {
                        direction: event.target.value === 'descending' ? 'descending' : 'ascending',
                      }),
                      index,
                    )
                  }
                >
                  {SORT_DIRECTION_OPTIONS.map((option) => (
                    <option key={option.value} value={option.value}>
                      {option.label}
                    </option>
                  ))}
                </select>
              </td>
            </tr>
          ))}
        </tbody>
      </table>

      <p className="dialog-hint sort-target" data-testid="sort-target">
        Sorts the range {targetLabel}. Change the selection to sort a different range.
      </p>

      {error === null ? null : (
        <p className="dialog-error" role="alert" data-testid="sort-error">
          {error}
        </p>
      )}
      {dispatchError === null ? null : (
        <p className="dialog-error" role="alert" data-testid="sort-dispatch-error">
          {dispatchError}
        </p>
      )}
    </Dialog>
  );
}
