import { useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import type { JSX, KeyboardEvent as ReactKeyboardEvent } from 'react';
import { searchChecklist } from '@model/autofilter';
import type { ChecklistEntry, ColumnFilter, ValueChecklist } from '@model/autofilter';
import type { SortDirection } from '@model/sort';
import { Dialog } from './Dialog';
import type { DialogCloseReason } from './Dialog';
import {
  buildConditions,
  checklistProblem,
  conditionsDraftFrom,
  criterionFromChecklist,
  defaultConditionsDraft,
  FILTER_ROW_HEIGHT,
  filterListWindow,
  filterMenuLabel,
  initialChecklistState,
  operatorAcceptsRange,
  operatorsFor,
  selectAllState,
  toggleChecklistBlanks,
  toggleChecklistEntry,
  toggleSelectAll,
} from './filter-dropdown-state';
import type {
  ConditionDraft,
  ConditionsDraft,
  FilterChecklistState,
  FilterColumnType,
} from './filter-dropdown-state';
import { placeAnchoredPanel } from './viewport-placement';

export interface FilterApplyOutcome {
  readonly ok: boolean;
  readonly message: string | null;
}

export interface FilterDropdownProps {
  /** The header label the AutoFilter uses for this column ("Column1" when it looks like data). */
  readonly columnLabel: string;
  /** The model's checklist: typed values, grouped by formatted display, in §2.4 order. */
  readonly checklist: ValueChecklist;
  /** The column's current criterion, or `null` when it has none. */
  readonly criterion: ColumnFilter | null;
  readonly filterType: FilterColumnType;
  readonly date1904: boolean;
  /** Where the overlay sits, in viewport pixels, under the column's header button. */
  readonly anchor: {
    readonly top: number;
    readonly left: number;
    /** Top edge of the header cell, used when the panel flips above it. */
    readonly headerTop: number;
  };
  readonly onSort: (direction: SortDirection) => void;
  readonly onCustomSort: () => void;
  readonly onClearColumn: () => void;
  readonly onApply: (criterion: ColumnFilter) => FilterApplyOutcome;
  readonly onClose: (reason: DialogCloseReason) => void;
}

const FOCUSABLE =
  'button:not([disabled]), [href], input:not([disabled]), select:not([disabled]), [tabindex]:not([tabindex="-1"])';

/**
 * The AutoFilter dropdown (design-refs m5-menus-dialogs §2; sort-autofilter
 * §3.8). It is a DOM overlay anchored under the column's header button, never
 * part of the canvas, and it holds the whole edit in view state: opening it,
 * sorting the checklist and typing in the search box create no transaction.
 *
 * The entries follow the capture's ruled deltas (ticket 28): Sort A to Z, Sort
 * Z to A, Customised Sort, then Clear Filter from '<column>', the
 * type-dependent filter menu, the search box, Select All, the values and
 * (Blanks), and the green Apply button. Select All is the first row of the
 * value list and scrolls with the values, exactly as Excel shows it (ticket
 * 36 item 5). Out-of-scope entries Excel greys
 * (Sort By Colour, Filter By Colour, Sheet View) are not rendered at all.
 *
 * Apply dispatches one `data.filterColumn` (or the clear command when every
 * value is ticked), so a filter change is one undo step (§3.9).
 */
export function FilterDropdown({
  columnLabel,
  checklist,
  criterion,
  filterType,
  date1904,
  anchor,
  onSort,
  onCustomSort,
  onClearColumn,
  onApply,
  onClose,
}: FilterDropdownProps): JSX.Element {
  const rootRef = useRef<HTMLDivElement | null>(null);
  const listRef = useRef<HTMLDivElement | null>(null);
  const selectAllRef = useRef<HTMLInputElement | null>(null);
  const [query, setQuery] = useState('');
  const [scrollTop, setScrollTop] = useState(0);
  const [state, setState] = useState<FilterChecklistState>(() =>
    initialChecklistState(checklist.entries, checklist.hasBlanks, criterion),
  );
  const [filtersOpen, setFiltersOpen] = useState(false);
  const [conditions, setConditions] = useState<ConditionsDraft | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [position, setPosition] = useState(() => ({
    top: anchor.top,
    left: anchor.left,
    flipped: false,
  }));
  const unsupported = criterion !== null && criterion.kind === 'unsupported';
  const selectAll = selectAllState(state, checklist.entries);
  const visibleEntries = useMemo(
    () => searchChecklist(checklist.entries, query),
    [checklist.entries, query],
  );
  // The state is indexed by the full checklist, while the search box only
  // narrows what is rendered. Mapping entries to their full-list index once
  // keeps a scroll in a long list from doing an indexOf per visible row.
  const entryIndex = useMemo(() => {
    const index = new Map<ChecklistEntry, number>();
    checklist.entries.forEach((entry, position) => index.set(entry, position));
    return index;
  }, [checklist.entries]);
  // Select All is the first slot of the list, so it counts as a row and
  // scrolls with the values (ticket 36 item 5).
  const rowCount = 1 + visibleEntries.length + (checklist.hasBlanks ? 1 : 0);
  const listWindow = filterListWindow(rowCount, scrollTop);

  // Measure before the first paint so an anchored panel near the bottom edge
  // is already flipped when it becomes visible.
  useLayoutEffect(() => {
    const root = rootRef.current;
    if (root === null) return;
    const rect = root.getBoundingClientRect();
    setPosition(
      placeAnchoredPanel(anchor.top, anchor.headerTop, anchor.left, rect.width, rect.height),
    );
  }, [anchor.headerTop, anchor.left, anchor.top]);

  useEffect(() => {
    const root = rootRef.current;
    root?.querySelector<HTMLElement>(FOCUSABLE)?.focus();
  }, []);

  useEffect(() => {
    if (selectAllRef.current !== null) selectAllRef.current.indeterminate = selectAll.some;
  }, [selectAll.some]);

  // A narrower search starts the list at the top again, so the first match is
  // visible rather than the list keeping a scroll offset from the old rows.
  useEffect(() => {
    setScrollTop(0);
    if (listRef.current !== null) listRef.current.scrollTop = 0;
  }, [query]);

  // A press outside closes without a transaction; the conditions window is a
  // modal of its own and owns the pointer while it is open.
  useEffect(() => {
    if (conditions !== null) return;
    const onPointerDown = (event: PointerEvent): void => {
      if (event.target instanceof Node && rootRef.current?.contains(event.target) === true) return;
      onClose('dismiss');
    };
    window.addEventListener('pointerdown', onPointerDown, true);
    return () => window.removeEventListener('pointerdown', onPointerDown, true);
  }, [conditions, onClose]);

  // The overlay is placed from the grid geometry, so a grid scroll or a window
  // resize dismisses it rather than leaving it detached from its column.
  // Scrolling inside the dropdown's own value list is not a grid scroll.
  useEffect(() => {
    const dismiss = (event?: Event): void => {
      if (event?.target instanceof Node && rootRef.current?.contains(event.target) === true) return;
      onClose('dismiss');
    };
    const onResize = (): void => dismiss();
    window.addEventListener('resize', onResize);
    window.addEventListener('scroll', dismiss, true);
    return () => {
      window.removeEventListener('resize', onResize);
      window.removeEventListener('scroll', dismiss, true);
    };
  }, [onClose]);

  const openConditions = (operator?: string): void => {
    const base =
      conditionsDraftFrom(criterion, filterType, date1904) ?? defaultConditionsDraft(filterType);
    // The quick items prefill the operator; an existing criterion keeps its
    // value when the operator matches, so reopening a filter shows it back.
    const draft =
      operator === undefined || base.first.operator === operator
        ? base
        : { ...base, first: { operator, value: '', to: '' } };
    setFiltersOpen(false);
    setError(null);
    setConditions(draft);
  };

  const applyChecklist = (): void => {
    if (unsupported) return;
    const problem = checklistProblem(state, checklist.entries);
    if (problem !== null) {
      setError(problem);
      return;
    }
    const next = criterionFromChecklist(state, checklist.entries);
    if (next === null) {
      // Every value is visible, so the column ends up unfiltered: one clear
      // command, which is still one undo step.
      onClearColumn();
      onClose('dismiss');
      return;
    }
    const outcome = onApply(next);
    if (!outcome.ok) {
      setError(outcome.message ?? 'The filter could not be applied.');
      return;
    }
    onClose('dismiss');
  };

  const applyConditions = (draft: ConditionsDraft): string | null => {
    const built = buildConditions(draft, filterType, date1904);
    if (built.ok === false) return built.message;
    const outcome = onApply({
      kind: 'conditions',
      join: built.join,
      conditions: built.conditions,
    });
    if (!outcome.ok) return outcome.message ?? 'The filter could not be applied.';
    setConditions(null);
    onClose('dismiss');
    return null;
  };

  const onKeyDown = (event: ReactKeyboardEvent<HTMLDivElement>): void => {
    if (conditions !== null) return;
    if (event.key === 'Escape') {
      event.preventDefault();
      event.stopPropagation();
      onClose('escape');
      return;
    }
    if (event.key === 'Enter') {
      const target = event.target;
      if (
        target instanceof HTMLButtonElement ||
        target instanceof HTMLSelectElement ||
        target instanceof HTMLAnchorElement
      ) {
        return;
      }
      // Enter in the search box or on a checkbox applies, as §3.8 specifies.
      event.preventDefault();
      event.stopPropagation();
      applyChecklist();
      return;
    }
    if (event.key !== 'ArrowDown' && event.key !== 'ArrowUp') return;
    const items = [...(rootRef.current?.querySelectorAll<HTMLElement>(FOCUSABLE) ?? [])];
    if (items.length === 0) return;
    event.preventDefault();
    event.stopPropagation();
    const current = items.indexOf(document.activeElement as HTMLElement);
    const offset = event.key === 'ArrowDown' ? 1 : -1;
    const next = current < 0 ? 0 : (current + offset + items.length) % items.length;
    items[next]?.focus();
  };

  // Only the rows inside the virtual window exist in the DOM. Select All is
  // the first slot and `(Blanks)` the last, so `(Blanks)` appears when the
  // list is scrolled to the bottom and Select All leaves with the values when
  // it scrolls away (ticket 36 item 5).
  const renderedRows: JSX.Element[] = [];
  for (let row = listWindow.start; row < listWindow.end; row += 1) {
    if (row === 0) {
      renderedRows.push(
        <label
          key="select-all"
          className="filter-check filter-select-all"
          style={{ height: FILTER_ROW_HEIGHT }}
        >
          <input
            ref={selectAllRef}
            type="checkbox"
            checked={selectAll.all}
            data-testid="filter-select-all"
            onChange={() => setState(toggleSelectAll(state, checklist.entries))}
          />
          <span className="filter-check-text">Select All</span>
        </label>,
      );
      continue;
    }
    if (row - 1 < visibleEntries.length) {
      const entry = visibleEntries[row - 1]!;
      const index = entryIndex.get(entry) ?? 0;
      renderedRows.push(
        <label
          key={`value-${index}`}
          className="filter-check"
          style={{ height: FILTER_ROW_HEIGHT }}
        >
          <input
            type="checkbox"
            checked={state.checked.has(index)}
            data-testid={`filter-value-${index}`}
            onChange={() => setState(toggleChecklistEntry(state, index))}
          />
          <span className="filter-check-text">{entry.display}</span>
        </label>,
      );
      continue;
    }
    renderedRows.push(
      <label key="blanks" className="filter-check" style={{ height: FILTER_ROW_HEIGHT }}>
        <input
          type="checkbox"
          checked={state.blanksChecked}
          data-testid="filter-blanks"
          onChange={() => setState(toggleChecklistBlanks(state))}
        />
        <span className="filter-check-text">(Blanks)</span>
      </label>,
    );
  }

  return (
    <div
      ref={rootRef}
      className="filter-dropdown"
      role="dialog"
      aria-label={`Filter ${columnLabel}`}
      style={{ top: position.top, left: position.left }}
      data-flipped={position.flipped ? 'true' : 'false'}
      onKeyDown={onKeyDown}
      data-testid="filter-dropdown"
    >
      <div className="filter-menu" role="menu" aria-label={`Sort and filter ${columnLabel}`}>
        <button
          type="button"
          role="menuitem"
          className="menu-item"
          data-testid="filter-sort-ascending"
          onClick={() => onSort('ascending')}
        >
          <span className="filter-icon" aria-hidden="true">
            A↓Z
          </span>
          <span className="menu-label">Sort A to Z</span>
        </button>
        <button
          type="button"
          role="menuitem"
          className="menu-item"
          data-testid="filter-sort-descending"
          onClick={() => onSort('descending')}
        >
          <span className="filter-icon" aria-hidden="true">
            Z↓A
          </span>
          <span className="menu-label">Sort Z to A</span>
        </button>
        <button
          type="button"
          role="menuitem"
          className="menu-item"
          data-testid="filter-custom-sort"
          onClick={onCustomSort}
        >
          <span className="filter-icon" aria-hidden="true" />
          <span className="menu-label">Customised Sort</span>
        </button>
        <div className="menu-separator" role="separator" />
        <button
          type="button"
          role="menuitem"
          className="menu-item"
          data-testid="filter-clear-column"
          disabled={criterion === null}
          onClick={() => {
            onClearColumn();
            onClose('dismiss');
          }}
        >
          <span className="filter-icon" aria-hidden="true" />
          <span className="menu-label">{`Clear Filter from '${columnLabel}'`}</span>
        </button>
      </div>

      {unsupported ? (
        <p className="filter-notice" role="status" data-testid="filter-unsupported">
          This filter uses a feature calcplus cannot edit.
        </p>
      ) : (
        <>
          <div className="filter-type-menu">
            <button
              type="button"
              className="menu-item filter-type-button"
              aria-haspopup="menu"
              aria-expanded={filtersOpen}
              data-testid="filter-type-menu"
              onClick={() => setFiltersOpen((open) => !open)}
            >
              <span className="filter-icon" aria-hidden="true" />
              <span className="menu-label">{filterMenuLabel(filterType)}</span>
              <span className="filter-submenu-caret" aria-hidden="true">
                ›
              </span>
            </button>
            {filtersOpen ? (
              <div
                className="filter-submenu"
                role="menu"
                aria-label={`${filterMenuLabel(filterType)} options`}
                data-testid="filter-type-options"
              >
                {operatorsFor(filterType).map((option) => (
                  <button
                    key={option.value}
                    type="button"
                    role="menuitem"
                    className="menu-item"
                    data-testid={`filter-type-${option.value}`}
                    onClick={() => openConditions(option.value)}
                  >
                    <span className="menu-label">{option.label}…</span>
                  </button>
                ))}
                <div className="menu-separator" role="separator" />
                <button
                  type="button"
                  role="menuitem"
                  className="menu-item"
                  data-testid="filter-custom-filter"
                  onClick={() => openConditions()}
                >
                  <span className="menu-label">Custom Filter…</span>
                </button>
              </div>
            ) : null}
          </div>

          <div className="filter-search">
            <span className="filter-search-icon" aria-hidden="true">
              ⌕
            </span>
            <input
              type="search"
              className="filter-search-input"
              aria-label="Search values"
              placeholder="Search"
              value={query}
              spellCheck={false}
              autoComplete="off"
              data-testid="filter-search"
              onChange={(event) => setQuery(event.target.value)}
            />
          </div>

          {visibleEntries.length === 0 ? <p className="filter-empty">No matching values.</p> : null}

          <div
            ref={listRef}
            className="filter-checklist"
            data-testid="filter-checklist"
            onScroll={(event) => setScrollTop(event.currentTarget.scrollTop)}
          >
            <div
              className="filter-spacer"
              style={{ height: listWindow.leading }}
              aria-hidden="true"
            />
            {renderedRows}
            <div
              className="filter-spacer"
              style={{ height: listWindow.trailing }}
              aria-hidden="true"
            />
          </div>

          <div className="filter-actions">
            <button
              type="button"
              className="filter-apply"
              data-testid="filter-apply"
              onClick={applyChecklist}
            >
              Apply
            </button>
          </div>
        </>
      )}

      {error === null ? null : (
        <p className="dialog-error filter-error" role="alert" data-testid="filter-error">
          {error}
        </p>
      )}

      {conditions === null ? null : (
        <CustomAutoFilterDialog
          type={filterType}
          initial={conditions}
          onClose={() => setConditions(null)}
          onApply={applyConditions}
        />
      )}
    </div>
  );
}

interface CustomAutoFilterDialogProps {
  readonly type: FilterColumnType;
  readonly initial: ConditionsDraft;
  readonly onClose: () => void;
  /** Returns an error message to keep the window open, or `null` once applied. */
  readonly onApply: (draft: ConditionsDraft) => string | null;
}

/**
 * The Custom AutoFilter conditions window (ui-chrome §4.1's "AutoFilter
 * conditions" dialog): two conditions joined by And or Or, with the operators
 * sort-autofilter §3.2 allows for the column's type. It dispatches nothing
 * itself; OK hands the draft to the dropdown, which submits one
 * `data.filterColumn`.
 */
function CustomAutoFilterDialog({
  type,
  initial,
  onClose,
  onApply,
}: CustomAutoFilterDialogProps): JSX.Element {
  const [draft, setDraft] = useState<ConditionsDraft>(initial);
  const [error, setError] = useState<string | null>(null);
  const options = operatorsFor(type);

  const submit = (): void => {
    const message = onApply(draft);
    if (message !== null) setError(message);
  };

  return (
    <Dialog
      title="Custom AutoFilter"
      onClose={() => onClose()}
      defaultAction={{ label: 'OK', onSelect: submit }}
      secondaryAction={{ label: 'Cancel', onSelect: onClose }}
      className="custom-filter-dialog"
    >
      <p className="dialog-hint">Show rows where:</p>
      <div className="custom-filter-row">
        <select
          className="find-select custom-filter-operator"
          aria-label="Operator 1"
          data-testid="custom-filter-operator-1"
          value={draft.first.operator}
          onChange={(event) =>
            setDraft({
              ...draft,
              first: { ...draft.first, operator: event.target.value, value: '', to: '' },
            })
          }
        >
          {options.map((option) => (
            <option key={option.value} value={option.value}>
              {option.label}
            </option>
          ))}
        </select>
        <ConditionValues
          idPrefix="custom-filter-1"
          label="Value 1"
          type={type}
          draft={draft.first}
          onChange={(first) => setDraft({ ...draft, first })}
        />
      </div>

      <div className="custom-filter-row custom-filter-join">
        <select
          className="find-select"
          aria-label="Join conditions"
          data-testid="custom-filter-join"
          value={draft.join}
          onChange={(event) =>
            setDraft({ ...draft, join: event.target.value === 'or' ? 'or' : 'and' })
          }
        >
          <option value="and">And</option>
          <option value="or">Or</option>
        </select>
      </div>

      <div className="custom-filter-row">
        <select
          className="find-select custom-filter-operator"
          aria-label="Operator 2"
          data-testid="custom-filter-operator-2"
          value={draft.second.operator}
          onChange={(event) =>
            setDraft({
              ...draft,
              second: { ...draft.second, operator: event.target.value, value: '', to: '' },
            })
          }
        >
          {options.map((option) => (
            <option key={option.value} value={option.value}>
              {option.label}
            </option>
          ))}
        </select>
        <ConditionValues
          idPrefix="custom-filter-2"
          label="Value 2"
          type={type}
          draft={draft.second}
          onChange={(second) => setDraft({ ...draft, second })}
        />
      </div>

      {error === null ? null : (
        <p className="dialog-error" role="alert" data-testid="custom-filter-error">
          {error}
        </p>
      )}
    </Dialog>
  );
}

interface ConditionValuesProps {
  readonly idPrefix: string;
  readonly label: string;
  readonly type: FilterColumnType;
  readonly draft: ConditionDraft;
  readonly onChange: (draft: ConditionDraft) => void;
}

/** The text, number or date value fields; a range operator takes two bounds. */
function ConditionValues({
  idPrefix,
  label,
  type,
  draft,
  onChange,
}: ConditionValuesProps): JSX.Element {
  const range = operatorAcceptsRange(draft.operator);
  const inputType = type === 'date' ? 'date' : type === 'number' ? 'number' : 'text';
  return (
    <span className="custom-filter-values">
      <input
        type={inputType}
        className="find-input custom-filter-value"
        aria-label={range ? `${label} from` : label}
        data-testid={`${idPrefix}-value`}
        value={draft.value}
        spellCheck={type === 'text' ? false : undefined}
        autoComplete="off"
        onChange={(event) => onChange({ ...draft, value: event.target.value })}
      />
      {range ? (
        <>
          <span className="custom-filter-and" aria-hidden="true">
            and
          </span>
          <input
            type={inputType}
            className="find-input custom-filter-value"
            aria-label={`${label} to`}
            data-testid={`${idPrefix}-to`}
            value={draft.to}
            spellCheck={type === 'text' ? false : undefined}
            autoComplete="off"
            onChange={(event) => onChange({ ...draft, to: event.target.value })}
          />
        </>
      ) : null}
    </span>
  );
}
