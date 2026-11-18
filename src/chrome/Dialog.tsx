import { useEffect, useId, useRef } from 'react';
import type { JSX, KeyboardEvent as ReactKeyboardEvent, ReactNode } from 'react';

export interface DialogAction {
  readonly label: string;
  readonly onSelect: () => void;
}

/** `escape` returns keyboard focus to the grid; `dismiss` returns it to the invoker. */
export type DialogCloseReason = 'dismiss' | 'escape';

export interface DialogProps {
  readonly title: string;
  readonly children: ReactNode;
  readonly onClose: (reason: DialogCloseReason) => void;
  /** Enter activates this action except in a multiline field, select, button or link. */
  readonly defaultAction?: DialogAction;
  /** The cancel action; Esc always calls `onClose`. */
  readonly secondaryAction?: DialogAction;
  readonly className?: string;
  /** Feature dialogs return focus to the invoker; the shell can return to the grid instead. */
  readonly restoreFocus?: boolean;
}

const FOCUSABLE =
  'button:not([disabled]), [href], input:not([disabled]), select:not([disabled]), textarea:not([disabled]), [tabindex]:not([tabindex="-1"])';

/**
 * The dialog shell for v1. It traps focus, closes on Esc and restores focus to
 * the invoking control on dismissal; Esc hands focus back to the grid.
 * Individual dialogs supply only their content and the commands they submit;
 * the shell owns layout and focus semantics.
 */
export function Dialog({
  title,
  children,
  onClose,
  defaultAction,
  secondaryAction,
  className,
  restoreFocus = true,
}: DialogProps): JSX.Element {
  const dialogRef = useRef<HTMLDivElement | null>(null);
  const defaultRef = useRef<HTMLButtonElement | null>(null);
  const returnFocusRef = useRef<HTMLElement | null>(null);
  const closedByEscapeRef = useRef(false);
  const titleId = useId();
  const bodyId = useId();

  useEffect(() => {
    returnFocusRef.current =
      document.activeElement instanceof HTMLElement ? document.activeElement : null;
    const dialog = dialogRef.current;
    if (dialog === null) return;
    focusInitial(dialog, defaultRef.current);
    return () => {
      if (!restoreFocus || closedByEscapeRef.current) return;
      const target = returnFocusRef.current;
      if (target !== null && target.isConnected) target.focus();
    };
  }, [restoreFocus]);

  // The trap is owned by the shell, not by the dialog's own keydown handler:
  // focus can leave the dialog through the backdrop or a programmatic change,
  // and Esc must close even then.
  useEffect(() => {
    const dialog = dialogRef.current;
    if (dialog === null) return;
    const onDocumentKeyDown = (event: KeyboardEvent): void => {
      if (event.key === 'Escape') {
        // A dialog can open its own popup over the form (for example the font
        // family list). While that popup is open its owner handles Escape to
        // close only the popup; the modal stays open.
        if (
          event.target instanceof Element &&
          event.target.closest('[data-dialog-popup-open="true"]') !== null
        ) {
          return;
        }
        event.preventDefault();
        event.stopPropagation();
        closedByEscapeRef.current = true;
        onClose('escape');
        return;
      }
      if (event.key === 'Tab') trapTab(event, dialog);
    };
    const onFocusIn = (event: FocusEvent): void => {
      if (event.target instanceof Node && dialog.contains(event.target)) return;
      // A backdrop press that slips through the pointer guard lands on the
      // body. Only that fallback is recaptured; a close action may legitimately
      // move focus to the grid while the dialog is unmounting.
      if (event.target !== document.body) return;
      focusInitial(dialog, defaultRef.current);
    };
    document.addEventListener('keydown', onDocumentKeyDown, true);
    document.addEventListener('focusin', onFocusIn);
    return () => {
      document.removeEventListener('keydown', onDocumentKeyDown, true);
      document.removeEventListener('focusin', onFocusIn);
    };
  }, [onClose]);

  const onKeyDown = (event: ReactKeyboardEvent<HTMLDivElement>): void => {
    if (event.key === 'Enter' && defaultAction !== undefined) {
      const target = event.target;
      if (
        target instanceof HTMLTextAreaElement ||
        target instanceof HTMLButtonElement ||
        target instanceof HTMLAnchorElement ||
        target instanceof HTMLSelectElement
      ) {
        return;
      }
      event.preventDefault();
      defaultAction.onSelect();
    }
  };

  return (
    <div
      className="dialog-backdrop"
      onPointerDown={(event) => {
        if (event.target !== event.currentTarget) return;
        // A backdrop press keeps the keyboard inside the modal. It never
        // dismisses: the dialog's own actions and Esc own closing.
        event.preventDefault();
        const dialog = dialogRef.current;
        if (dialog !== null) focusInitial(dialog, defaultRef.current);
      }}
      onClick={(event) => {
        if (event.target !== event.currentTarget) return;
        const dialog = dialogRef.current;
        if (dialog !== null) focusInitial(dialog, defaultRef.current);
      }}
    >
      <div
        ref={dialogRef}
        className={className === undefined ? 'dialog' : `dialog ${className}`}
        role="dialog"
        aria-modal="true"
        aria-labelledby={titleId}
        aria-describedby={bodyId}
        tabIndex={-1}
        onKeyDown={onKeyDown}
      >
        <header className="dialog-header">
          <h2 id={titleId} className="dialog-title">
            {title}
          </h2>
          <button
            type="button"
            className="dialog-close"
            aria-label="Close dialog"
            onClick={() => onClose('dismiss')}
          >
            ×
          </button>
        </header>
        <div id={bodyId} className="dialog-body">
          {children}
        </div>
        {defaultAction === undefined && secondaryAction === undefined ? null : (
          <footer className="dialog-actions">
            {secondaryAction === undefined ? null : (
              <button type="button" className="dialog-button" onClick={secondaryAction.onSelect}>
                {secondaryAction.label}
              </button>
            )}
            {defaultAction === undefined ? null : (
              <button
                ref={defaultRef}
                type="button"
                className="dialog-button dialog-button-primary"
                onClick={defaultAction.onSelect}
              >
                {defaultAction.label}
              </button>
            )}
          </footer>
        )}
      </div>
    </div>
  );
}

export interface DialogTab {
  readonly id: string;
  readonly label: string;
}

export interface DialogTabsProps {
  readonly tabs: readonly DialogTab[];
  readonly activeId: string;
  readonly onChange: (id: string) => void;
}

/** The tab strip shared by tabbed dialogs such as Format Cells. */
export function DialogTabs({ tabs, activeId, onChange }: DialogTabsProps): JSX.Element {
  return (
    <div className="dialog-tabs" role="tablist">
      {tabs.map((tab) => (
        <button
          key={tab.id}
          type="button"
          role="tab"
          aria-selected={tab.id === activeId}
          className={tab.id === activeId ? 'dialog-tab dialog-tab-active' : 'dialog-tab'}
          onClick={() => onChange(tab.id)}
          onKeyDown={(event) => {
            if (event.key !== 'ArrowLeft' && event.key !== 'ArrowRight') return;
            event.preventDefault();
            const index = tabs.findIndex((candidate) => candidate.id === activeId);
            const offset = event.key === 'ArrowRight' ? 1 : -1;
            const next = tabs[(index + offset + tabs.length) % tabs.length];
            if (next !== undefined) onChange(next.id);
          }}
        >
          {tab.label}
        </button>
      ))}
    </div>
  );
}

function trapTab(
  event: Pick<KeyboardEvent, 'preventDefault' | 'shiftKey'>,
  dialog: HTMLDivElement | null,
): void {
  if (dialog === null) return;
  const elements = focusable(dialog);
  if (elements.length === 0) {
    event.preventDefault();
    return;
  }
  const first = elements[0]!;
  const last = elements[elements.length - 1]!;
  const active = document.activeElement;
  if (event.shiftKey && (active === first || active === dialog || !dialog.contains(active))) {
    event.preventDefault();
    last.focus();
    return;
  }
  if (!event.shiftKey && (active === last || active === dialog || !dialog.contains(active))) {
    event.preventDefault();
    first.focus();
  }
}

function focusInitial(dialog: HTMLDivElement, preferred: HTMLButtonElement | null): void {
  const target = preferred ?? focusable(dialog)[0] ?? dialog;
  target.focus();
}

function focusable(root: HTMLElement): HTMLElement[] {
  return [...root.querySelectorAll<HTMLElement>(FOCUSABLE)];
}
