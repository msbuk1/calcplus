import { useEffect, useLayoutEffect, useRef, useState } from 'react';
import type { JSX, MouseEvent as ReactMouseEvent } from 'react';

export interface SheetTabInfo {
  readonly id: string;
  readonly name: string;
}

export interface SheetTabsProps {
  /** The workbook's ordered visible sheets, as plain chrome records. */
  readonly sheets: readonly SheetTabInfo[];
  readonly activeSheetId: string | null;
  readonly onSelectSheet: (sheetId: string) => void;
  readonly onInsertSheet: () => void;
  readonly onRenameSheet: (sheetId: string, name: string) => void;
  readonly onDeleteSheet: (sheetId: string) => void;
}

interface TabMenu {
  readonly sheetId: string;
  readonly x: number;
  readonly y: number;
}

/**
 * The sheet tab strip.
 *
 * It reads only plain ordered sheet records and command callbacks; it never
 * imports the workbook model. Switching tabs is view state, while insert,
 * rename and delete go through the command surface supplied by the app.
 * Deleting a sheet is permanent, so the app asks before the command runs and
 * the menu labels the action as one that cannot be undone.
 */
export function SheetTabs({
  sheets,
  activeSheetId,
  onSelectSheet,
  onInsertSheet,
  onRenameSheet,
  onDeleteSheet,
}: SheetTabsProps): JSX.Element {
  const stripRef = useRef<HTMLDivElement | null>(null);
  const inputRef = useRef<HTMLInputElement | null>(null);
  const menuRef = useRef<HTMLDivElement | null>(null);
  const editingRef = useRef<string | null>(null);
  const [editingId, setEditingId] = useState<string | null>(null);
  const [draft, setDraft] = useState('');
  const [menu, setMenu] = useState<TabMenu | null>(null);

  useLayoutEffect(() => {
    if (editingId === null) return;
    const input = inputRef.current;
    if (input === null) return;
    input.focus();
    input.select();
  }, [editingId]);

  // An inserted or selected tab must not stay out of view behind the arrows.
  useEffect(() => {
    const strip = stripRef.current;
    if (strip === null || activeSheetId === null) return;
    const active = strip.querySelector<HTMLElement>('[aria-selected="true"]');
    if (active === null) return;
    const stripRect = strip.getBoundingClientRect();
    const tabRect = active.getBoundingClientRect();
    if (tabRect.left < stripRect.left) {
      strip.scrollLeft -= stripRect.left - tabRect.left;
    } else if (tabRect.right > stripRect.right) {
      strip.scrollLeft += tabRect.right - stripRect.right;
    }
  }, [activeSheetId, sheets.length]);

  useEffect(() => {
    if (menu === null) return;
    menuRef.current?.querySelector<HTMLButtonElement>('[role="menuitem"]')?.focus();
    const closeOnPointer = (event: PointerEvent): void => {
      const element = menuRef.current;
      if (element !== null && event.target instanceof Node && element.contains(event.target)) {
        return;
      }
      setMenu(null);
    };
    const closeOnKey = (event: KeyboardEvent): void => {
      const element = menuRef.current;
      const items =
        element === null
          ? []
          : [...element.querySelectorAll<HTMLButtonElement>('[role="menuitem"]')];
      if (event.key === 'Escape') {
        event.preventDefault();
        event.stopPropagation();
        setMenu(null);
        const strip = stripRef.current;
        const tab = [...(strip?.querySelectorAll<HTMLButtonElement>('[role="tab"]') ?? [])].find(
          (candidate) => candidate.dataset.sheetId === menu.sheetId,
        );
        tab?.focus();
        return;
      }
      if (event.key === 'Tab') {
        setMenu(null);
        return;
      }
      if (event.key !== 'ArrowDown' && event.key !== 'ArrowUp') return;
      if (items.length === 0) return;
      event.preventDefault();
      const current = items.indexOf(document.activeElement as HTMLButtonElement);
      const offset = event.key === 'ArrowDown' ? 1 : -1;
      const next = (current + offset + items.length) % items.length;
      items[next]?.focus();
    };
    window.addEventListener('pointerdown', closeOnPointer);
    window.addEventListener('keydown', closeOnKey);
    return () => {
      window.removeEventListener('pointerdown', closeOnPointer);
      window.removeEventListener('keydown', closeOnKey);
    };
  }, [menu]);

  const scrollTabs = (direction: -1 | 1): void => {
    stripRef.current?.scrollBy({ left: direction * 160, behavior: 'smooth' });
  };

  const beginRename = (sheet: SheetTabInfo): void => {
    editingRef.current = sheet.id;
    setEditingId(sheet.id);
    setDraft(sheet.name);
    setMenu(null);
  };

  const finishRename = (commit: boolean): void => {
    const sheetId = editingRef.current;
    if (sheetId === null) return;
    editingRef.current = null;
    setEditingId(null);
    if (!commit) return;
    const name = draft.trim();
    const sheet = sheets.find((candidate) => candidate.id === sheetId);
    if (sheet === undefined || name === '' || name === sheet.name) return;
    onRenameSheet(sheetId, name);
  };

  const openMenu = (event: ReactMouseEvent<HTMLElement>, sheetId: string): void => {
    event.preventDefault();
    event.stopPropagation();
    setMenu({ sheetId, x: event.clientX, y: event.clientY });
  };

  const menuSheet = menu === null ? undefined : sheets.find((sheet) => sheet.id === menu.sheetId);

  return (
    <div className="sheet-tabs" data-testid="sheet-tabs">
      <button
        type="button"
        className="sheet-tabs-scroll"
        title="Scroll sheet tabs left"
        aria-label="Scroll sheet tabs left"
        onClick={() => scrollTabs(-1)}
        data-testid="sheet-tabs-scroll-left"
      >
        ‹
      </button>
      <div className="sheet-tab-strip" ref={stripRef} role="tablist" aria-label="Sheets">
        {sheets.map((sheet) =>
          editingId === sheet.id ? (
            <input
              key={sheet.id}
              ref={inputRef}
              className="sheet-tab-rename"
              value={draft}
              aria-label={`Rename sheet ${sheet.name}`}
              spellCheck={false}
              data-testid="sheet-rename-input"
              onChange={(event) => setDraft(event.target.value)}
              onBlur={() => finishRename(true)}
              onKeyDown={(event) => {
                if (event.key === 'Enter') {
                  event.preventDefault();
                  finishRename(true);
                  return;
                }
                if (event.key === 'Escape') {
                  event.preventDefault();
                  finishRename(false);
                }
              }}
            />
          ) : (
            <button
              key={sheet.id}
              type="button"
              role="tab"
              aria-selected={sheet.id === activeSheetId}
              className={sheet.id === activeSheetId ? 'sheet-tab sheet-tab-active' : 'sheet-tab'}
              title={sheet.name}
              data-sheet-id={sheet.id}
              data-testid={`sheet-tab-${sheet.name}`}
              onClick={() => onSelectSheet(sheet.id)}
              onDoubleClick={() => beginRename(sheet)}
              onKeyDown={(event) => {
                if (event.key !== 'F2') return;
                event.preventDefault();
                beginRename(sheet);
              }}
              onContextMenu={(event) => openMenu(event, sheet.id)}
            >
              {sheet.name}
            </button>
          ),
        )}
      </div>
      <button
        type="button"
        className="sheet-tabs-scroll"
        title="Scroll sheet tabs right"
        aria-label="Scroll sheet tabs right"
        onClick={() => scrollTabs(1)}
        data-testid="sheet-tabs-scroll-right"
      >
        ›
      </button>
      <button
        type="button"
        className="sheet-tab-add"
        title="Insert sheet"
        aria-label="Insert sheet"
        onClick={onInsertSheet}
        data-command="sheet.insert"
        data-testid="sheet-add-button"
      >
        +
      </button>

      {menu !== null && menuSheet !== undefined ? (
        <div
          ref={menuRef}
          className="sheet-tab-menu"
          role="menu"
          style={{ left: `${menu.x}px`, top: `${menu.y}px` }}
          data-testid="sheet-tab-menu"
        >
          <button
            type="button"
            role="menuitem"
            data-command="sheet.rename"
            data-testid="sheet-rename-action"
            onClick={() => {
              if (menuSheet !== undefined) beginRename(menuSheet);
            }}
          >
            Rename
          </button>
          <button
            type="button"
            role="menuitem"
            data-command="sheet.delete"
            data-testid="sheet-delete-action"
            title="Delete this sheet permanently; this cannot be undone"
            aria-label={`Delete sheet ${menuSheet.name}. This cannot be undone.`}
            onClick={() => {
              if (menuSheet === undefined) return;
              setMenu(null);
              onDeleteSheet(menuSheet.id);
            }}
          >
            Delete
          </button>
        </div>
      ) : null}
    </div>
  );
}
