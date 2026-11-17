import { useEffect, useId, useRef, useState } from 'react';
import type { JSX } from 'react';
import { Dialog } from './Dialog';
import type { DialogAction, DialogCloseReason } from './Dialog';

export type FindScope = 'sheet' | 'workbook';

/** The exact payload `find.search` accepts (command-system §6, matrix row 1). */
export interface FindQuery {
  readonly query: string;
  readonly scope: FindScope;
  readonly matchCase: boolean;
  readonly wholeCell: boolean;
  readonly inFormulas: boolean;
  readonly inValues: boolean;
}

/**
 * Replace commands take the same search payload plus the replacement text. An
 * empty replacement is valid: it clears the matched content.
 */
export interface FindReplaceQuery extends FindQuery {
  readonly replacement: string;
}

/**
 * One row of the Find All list. The command result carries the match location;
 * the sheet name and address are presentation text, never model access.
 */
export interface FindMatchView {
  readonly key: string;
  readonly sheetId: string;
  readonly sheetName: string;
  readonly row: number;
  readonly col: number;
  readonly address: string;
  readonly preview: string;
}

/**
 * What a find or replace command reports to the dialog. `error` is reserved for
 * a dispatch failure; a query the user can fix is rejected before dispatch.
 */
export interface FindOutcome {
  readonly error: string | null;
  readonly message: string | null;
  readonly matches: readonly FindMatchView[];
  readonly total: number;
  readonly replaced: number;
}

export interface FindDialogProps {
  /** The launcher owns the window title: Find or Find and Replace. */
  readonly title: string;
  readonly onClose: (reason: DialogCloseReason) => void;
  readonly onSearch: (query: FindQuery, all: boolean) => FindOutcome;
  readonly onReplace: (query: FindReplaceQuery, all: boolean) => FindOutcome;
  readonly onActivateMatch: (match: FindMatchView) => void;
}

type SubmitValidation = { readonly query: FindQuery } | { readonly error: string };

/** A Find All list is a view: keep a large match set from building a huge DOM. */
const FIND_RESULT_LIMIT = 500;

/**
 * Find & Replace. The dialog owns the form state and submits one command per
 * action; it never searches the workbook itself. Find Next and Find All share
 * `find.search`; Replace and Replace All submit `find.replaceOne` and
 * `find.replaceAll`. The Dialog shell owns the focus trap, Esc and the restored
 * invoker focus.
 */
export function FindDialog({
  title,
  onClose,
  onSearch,
  onReplace,
  onActivateMatch,
}: FindDialogProps): JSX.Element {
  const [query, setQuery] = useState('');
  const [replacement, setReplacement] = useState('');
  const [scope, setScope] = useState<FindScope>('sheet');
  const [matchCase, setMatchCase] = useState(false);
  const [wholeCell, setWholeCell] = useState(false);
  const [inFormulas, setInFormulas] = useState(true);
  const [inValues, setInValues] = useState(true);
  const [outcome, setOutcome] = useState<FindOutcome | null>(null);
  const [error, setError] = useState<string | null>(null);
  // Enter follows the last field the user touched: the query searches, the
  // replacement field replaces. The shell's default action carries the label.
  const [enterIntent, setEnterIntent] = useState<'find' | 'replace'>('find');
  const queryRef = useRef<HTMLInputElement | null>(null);
  const queryId = useId();
  const replaceId = useId();
  const scopeId = useId();

  // The shell focuses its default action first. A feature dialog puts the
  // caret in its primary field instead, after that effect has run.
  useEffect(() => {
    const handle = requestAnimationFrame(() => {
      const input = queryRef.current;
      if (input === null) return;
      input.focus();
      input.select();
    });
    return () => cancelAnimationFrame(handle);
  }, []);

  const resetResult = (): void => {
    setOutcome(null);
    setError(null);
  };

  const validate = (): SubmitValidation => {
    if (query === '') {
      return { error: 'Enter a value or formula to find.' };
    }
    if (!inFormulas && !inValues) {
      return { error: 'Choose formulas, values, or both.' };
    }
    return {
      query: { query, scope, matchCase, wholeCell, inFormulas, inValues },
    };
  };

  const submitSearch = (all: boolean): void => {
    const validation = validate();
    if ('error' in validation) {
      setError(validation.error);
      setOutcome(null);
      return;
    }
    setError(null);
    setOutcome(onSearch(validation.query, all));
  };

  const submitReplace = (all: boolean): void => {
    const validation = validate();
    if ('error' in validation) {
      setError(validation.error);
      setOutcome(null);
      return;
    }
    setError(null);
    setOutcome(onReplace({ ...validation.query, replacement }, all));
  };

  const defaultAction: DialogAction =
    enterIntent === 'replace'
      ? { label: 'Replace', onSelect: () => submitReplace(false) }
      : { label: 'Find Next', onSelect: () => submitSearch(false) };

  const displayedError = error ?? outcome?.error ?? null;
  const summary = outcome === null || displayedError !== null ? null : outcomeSummary(outcome);
  const visibleMatches = outcome === null ? [] : outcome.matches.slice(0, FIND_RESULT_LIMIT);

  return (
    <Dialog
      title={title}
      onClose={onClose}
      defaultAction={defaultAction}
      secondaryAction={{ label: 'Close', onSelect: () => onClose('dismiss') }}
      className="find-dialog"
    >
      <div className="find-form">
        <div className="find-row">
          <label className="find-label" htmlFor={queryId}>
            Find what
          </label>
          <input
            ref={queryRef}
            id={queryId}
            type="text"
            className="find-input"
            value={query}
            spellCheck={false}
            autoComplete="off"
            data-testid="find-query"
            onFocus={() => setEnterIntent('find')}
            onChange={(event) => {
              setQuery(event.target.value);
              resetResult();
            }}
          />
        </div>

        <div className="find-row">
          <label className="find-label" htmlFor={replaceId}>
            Replace with
          </label>
          <input
            id={replaceId}
            type="text"
            className="find-input"
            value={replacement}
            spellCheck={false}
            autoComplete="off"
            data-testid="find-replacement"
            onFocus={() => setEnterIntent('replace')}
            onChange={(event) => setReplacement(event.target.value)}
          />
        </div>

        <div className="find-row">
          <label className="find-label" htmlFor={scopeId}>
            Within
          </label>
          <select
            id={scopeId}
            className="find-select"
            value={scope}
            data-testid="find-scope"
            onChange={(event) => {
              setScope(event.target.value === 'workbook' ? 'workbook' : 'sheet');
              resetResult();
            }}
          >
            <option value="sheet">Sheet</option>
            <option value="workbook">Workbook</option>
          </select>
        </div>

        <fieldset className="find-options">
          <legend>Options</legend>
          <label className="find-check">
            <input
              type="checkbox"
              checked={matchCase}
              data-testid="find-match-case"
              onChange={(event) => {
                setMatchCase(event.target.checked);
                resetResult();
              }}
            />
            Match case
          </label>
          <label className="find-check">
            <input
              type="checkbox"
              checked={wholeCell}
              data-testid="find-whole-cell"
              onChange={(event) => {
                setWholeCell(event.target.checked);
                resetResult();
              }}
            />
            Match entire cell contents
          </label>
          <label className="find-check">
            <input
              type="checkbox"
              checked={inFormulas}
              data-testid="find-in-formulas"
              onChange={(event) => {
                setInFormulas(event.target.checked);
                resetResult();
              }}
            />
            In formulas
          </label>
          <label className="find-check">
            <input
              type="checkbox"
              checked={inValues}
              data-testid="find-in-values"
              onChange={(event) => {
                setInValues(event.target.checked);
                resetResult();
              }}
            />
            In values
          </label>
        </fieldset>

        <div className="find-actions">
          {enterIntent === 'replace' ? (
            <button type="button" onClick={() => submitSearch(false)} data-testid="find-next">
              Find Next
            </button>
          ) : null}
          <button type="button" onClick={() => submitSearch(true)} data-testid="find-all">
            Find All
          </button>
          {enterIntent === 'find' ? (
            <button type="button" onClick={() => submitReplace(false)} data-testid="find-replace">
              Replace
            </button>
          ) : null}
          <button type="button" onClick={() => submitReplace(true)} data-testid="find-replace-all">
            Replace All
          </button>
        </div>

        {displayedError === null ? null : (
          <p className="dialog-error" role="alert" data-testid="find-error">
            {displayedError}
          </p>
        )}

        {summary === null ? null : (
          <p className="dialog-message" role="status" data-testid="find-summary">
            {summary}
          </p>
        )}

        {outcome !== null && outcome.matches.length > 0 ? (
          <div
            className="find-results"
            data-testid="find-results"
            role="region"
            aria-label="Find results"
          >
            <p className="find-results-count">
              {outcome.total} match{outcome.total === 1 ? '' : 'es'}
            </p>
            <ul className="find-result-list">
              {visibleMatches.map((match, index) => (
                <li key={match.key}>
                  <button
                    type="button"
                    className="find-result"
                    data-testid={`find-result-${index}`}
                    onClick={() => onActivateMatch(match)}
                  >
                    <span className="find-result-address">
                      {match.sheetName}!{match.address}
                    </span>
                    {match.preview === '' ? null : (
                      <span className="find-result-preview">{match.preview}</span>
                    )}
                  </button>
                </li>
              ))}
            </ul>
            {outcome.matches.length > visibleMatches.length ? (
              <p className="find-results-count">
                Showing the first {visibleMatches.length} of {outcome.matches.length} matches.
              </p>
            ) : null}
          </div>
        ) : null}
      </div>
    </Dialog>
  );
}

function outcomeSummary(outcome: FindOutcome): string {
  if (outcome.message !== null) return outcome.message;
  if (outcome.replaced > 0) {
    return `Replaced ${outcome.replaced} cell${outcome.replaced === 1 ? '' : 's'}.`;
  }
  if (outcome.total > 0) {
    return `Found ${outcome.total} match${outcome.total === 1 ? '' : 'es'}.`;
  }
  return 'No matches found.';
}
