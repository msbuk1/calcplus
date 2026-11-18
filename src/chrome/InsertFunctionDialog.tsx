import { useEffect, useId, useMemo, useRef, useState } from 'react';
import type { JSX, KeyboardEvent as ReactKeyboardEvent } from 'react';
import {
  FUNCTION_CATEGORY_OPTIONS,
  functionsInCategory,
  searchFunctions,
} from '@engine/function-catalog';
import type { FunctionCatalogEntry, FunctionCategoryOption } from '@engine/function-catalog';
import type { PublishedFunction } from '@engine/function-registry';
import { Dialog } from './Dialog';
import type { DialogCloseReason } from './Dialog';

export interface InsertFunctionDialogProps {
  /** Functions already inserted this session, most recent first. */
  readonly recentlyUsed?: readonly PublishedFunction[];
  readonly onClose: (reason: DialogCloseReason) => void;
  /** Insert the chosen function's skeleton into the formula edit surface. */
  readonly onInsert: (name: PublishedFunction) => void;
}

/**
 * The v1 Insert Function dialog (ticket 30): search, category, function list,
 * signature and description. Confirming Insert places `=NAME(` in the formula
 * bar with the caret on the first argument; the argument wizard is V2 and has
 * no controls here. The Dialog shell owns the focus trap, Esc and the default
 * button, so this component supplies only the form and the list.
 */
export function InsertFunctionDialog({
  recentlyUsed = [],
  onClose,
  onInsert,
}: InsertFunctionDialogProps): JSX.Element {
  const [query, setQuery] = useState('');
  const [category, setCategory] = useState<FunctionCategoryOption>('Recently used');
  const [selected, setSelected] = useState<PublishedFunction | null>(null);

  const searchRef = useRef<HTMLInputElement | null>(null);
  const optionRefs = useRef<Array<HTMLDivElement | null>>([]);
  const queryId = useId();
  const categoryId = useId();
  const listId = useId();
  const detailsId = useId();

  const visible = useMemo(
    () =>
      query.trim() === '' ? functionsInCategory(category, recentlyUsed) : searchFunctions(query),
    [query, category, recentlyUsed],
  );

  // The highlighted row is the one Insert acts on. A category or query change
  // that hides the previous choice falls back to the first visible row.
  const activeEntry: FunctionCatalogEntry | null =
    visible.find((entry) => entry.name === selected) ?? visible[0] ?? null;
  const activeIndex =
    activeEntry === null ? -1 : visible.findIndex((entry) => entry.name === activeEntry.name);

  // The shell focuses its default action first; this dialog starts in search.
  useEffect(() => {
    const handle = requestAnimationFrame(() => {
      const input = searchRef.current;
      if (input === null) return;
      input.focus();
      input.select();
    });
    return () => cancelAnimationFrame(handle);
  }, []);

  const focusOption = (index: number): void => {
    const entry = visible[index];
    if (entry === undefined) return;
    setSelected(entry.name);
    optionRefs.current[index]?.focus();
  };

  const moveActive = (offset: number): void => {
    if (visible.length === 0) return;
    const from = activeIndex < 0 ? (offset > 0 ? -1 : 0) : activeIndex;
    focusOption((from + offset + visible.length) % visible.length);
  };

  const onListKeyDown = (event: ReactKeyboardEvent<HTMLDivElement>): void => {
    if (event.key === 'ArrowDown') {
      event.preventDefault();
      moveActive(1);
      return;
    }
    if (event.key === 'ArrowUp') {
      event.preventDefault();
      moveActive(-1);
      return;
    }
    if (event.key === 'Home') {
      event.preventDefault();
      focusOption(0);
      return;
    }
    if (event.key === 'End') {
      event.preventDefault();
      focusOption(visible.length - 1);
    }
  };

  const submit = (): void => {
    if (activeEntry === null) return;
    onInsert(activeEntry.name);
  };

  return (
    <Dialog
      title="Insert function"
      onClose={onClose}
      defaultAction={{ label: 'Insert', onSelect: submit }}
      secondaryAction={{ label: 'Cancel', onSelect: () => onClose('dismiss') }}
      className="insert-function-dialog"
    >
      <div className="insert-function-search">
        <span className="insert-function-search-icon" aria-hidden="true">
          <svg viewBox="0 0 16 16" focusable="false">
            <circle cx="7" cy="7" r="4.5" />
            <path d="M10.5 10.5 14 14" />
          </svg>
        </span>
        <input
          ref={searchRef}
          id={queryId}
          type="text"
          className="find-input insert-function-search-input"
          value={query}
          spellCheck={false}
          autoComplete="off"
          placeholder="Search for a function"
          aria-label="Search for a function"
          aria-controls={listId}
          aria-describedby={detailsId}
          data-testid="insert-function-search"
          onChange={(event) => {
            setQuery(event.target.value);
            setSelected(null);
          }}
          onKeyDown={(event) => {
            if (event.key === 'ArrowDown') {
              event.preventDefault();
              focusOption(Math.max(0, activeIndex));
            }
          }}
        />
      </div>

      <div className="insert-function-category">
        <label className="find-label" htmlFor={categoryId}>
          Select category
        </label>
        <select
          id={categoryId}
          className="find-select"
          value={category}
          data-testid="insert-function-category"
          onChange={(event) => {
            setCategory(event.target.value as FunctionCategoryOption);
            // A category switch shows that category; it ends an active search.
            setQuery('');
            setSelected(null);
          }}
        >
          {FUNCTION_CATEGORY_OPTIONS.map((option) => (
            <option key={option} value={option}>
              {option}
            </option>
          ))}
        </select>
      </div>

      {visible.length === 0 ? (
        <p className="insert-function-empty" role="status" data-testid="insert-function-empty">
          {query.trim() === ''
            ? 'No functions have been used yet. Choose All or another category.'
            : `No function matches "${query.trim()}".`}
        </p>
      ) : (
        <div
          id={listId}
          className="insert-function-list"
          role="listbox"
          aria-label="Functions"
          data-testid="insert-function-list"
          onKeyDown={onListKeyDown}
        >
          {visible.map((entry, index) => (
            <div
              key={entry.name}
              ref={(node) => {
                optionRefs.current[index] = node;
              }}
              className={
                entry.name === activeEntry?.name
                  ? 'insert-function-option insert-function-option-active'
                  : 'insert-function-option'
              }
              role="option"
              aria-selected={entry.name === activeEntry?.name}
              tabIndex={index === activeIndex ? 0 : -1}
              data-testid={`insert-function-option-${entry.name}`}
              onClick={() => setSelected(entry.name)}
            >
              {entry.name}
            </div>
          ))}
        </div>
      )}

      <div className="insert-function-details" id={detailsId}>
        {activeEntry === null ? null : (
          <>
            <p className="insert-function-signature" data-testid="insert-function-signature">
              {activeEntry.signature}
            </p>
            <p className="insert-function-description" data-testid="insert-function-description">
              {activeEntry.description}
            </p>
          </>
        )}
      </div>
    </Dialog>
  );
}
