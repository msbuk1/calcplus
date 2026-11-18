import { useEffect, useRef } from 'react';
import type { JSX, KeyboardEvent as ReactKeyboardEvent, RefObject } from 'react';
import type { CSSProperties } from 'react';

export interface MenuItemModel {
  readonly id: string;
  readonly label: string;
  readonly commandId?: string;
  readonly args?: unknown;
  readonly shortcut?: string;
  readonly disabled?: boolean;
  readonly checked?: boolean;
  readonly dividerBefore?: boolean;
  readonly testId?: string;
}

export interface MenuProps {
  readonly items: readonly MenuItemModel[];
  readonly onSelect: (item: MenuItemModel) => void;
  readonly onClose: () => void;
  readonly label?: string;
  /** A bold, non-interactive section heading above the items (Excel style). */
  readonly heading?: string;
  readonly className?: string;
  readonly style?: CSSProperties;
  /** A click on this element (the trigger) must not count as an outside click. */
  readonly anchorRef?: RefObject<HTMLElement | null>;
  /** The rendered menu element, for owners that must measure it before positioning. */
  readonly rootRef?: RefObject<HTMLDivElement | null>;
}

/**
 * One keyboard-navigable menu for the File menu, ribbon split buttons and the
 * grid context menu. Arrow keys move, Enter or Space activates, Esc closes.
 * The menu never dispatches anything itself; it hands the item to its owner.
 */
export function Menu({
  items,
  onSelect,
  onClose,
  label = 'Menu',
  heading,
  className,
  style,
  anchorRef,
  rootRef,
}: MenuProps): JSX.Element {
  const menuRef = useRef<HTMLDivElement | null>(null);
  const itemRefs = useRef<(HTMLButtonElement | null)[]>([]);
  const setRoot = (node: HTMLDivElement | null): void => {
    menuRef.current = node;
    if (rootRef !== undefined) rootRef.current = node;
  };

  useEffect(() => {
    const first = items.findIndex((item) => item.disabled !== true);
    if (first >= 0) itemRefs.current[first]?.focus();
    // Focus the first item once, when the menu opens.
  }, []);

  useEffect(() => {
    const closeOnPointer = (event: PointerEvent): void => {
      const target = event.target;
      if (!(target instanceof Node)) return;
      if (menuRef.current?.contains(target) === true) return;
      if (anchorRef?.current?.contains(target) === true) return;
      onClose();
    };
    const closeOnResize = (): void => onClose();
    /**
     * A click can scroll the anchor into view, and the browser delivers that
     * scroll event after the menu has mounted. Ignore scrolls until the next
     * frame so the opening interaction cannot close the menu it just opened;
     * a real scroll afterwards still closes it.
     */
    let armed = false;
    const frame = requestAnimationFrame(() => {
      armed = true;
    });
    const closeOnScroll = (event: Event): void => {
      if (!armed) return;
      if (event.target instanceof Node && menuRef.current?.contains(event.target) === true) return;
      onClose();
    };
    window.addEventListener('pointerdown', closeOnPointer);
    window.addEventListener('resize', closeOnResize);
    window.addEventListener('scroll', closeOnScroll, true);
    return () => {
      cancelAnimationFrame(frame);
      window.removeEventListener('pointerdown', closeOnPointer);
      window.removeEventListener('resize', closeOnResize);
      window.removeEventListener('scroll', closeOnScroll, true);
    };
  }, [anchorRef, onClose]);

  const focusAt = (start: number, direction: 1 | -1): void => {
    if (items.length === 0) return;
    let index = start;
    for (let step = 0; step < items.length; step += 1) {
      index = (index + direction + items.length) % items.length;
      const item = items[index];
      if (item !== undefined && item.disabled !== true) {
        itemRefs.current[index]?.focus();
        return;
      }
    }
  };

  const onKeyDown = (event: ReactKeyboardEvent<HTMLDivElement>): void => {
    if (event.key === 'Escape') {
      event.preventDefault();
      event.stopPropagation();
      onClose();
      return;
    }
    if (event.key === 'Tab') {
      event.preventDefault();
      event.stopPropagation();
      onClose();
      return;
    }
    if (event.key === 'F6') {
      onClose();
      return;
    }
    if (event.key === 'ArrowDown') {
      event.preventDefault();
      event.stopPropagation();
      focusAt(currentIndex(itemRefs.current), 1);
      return;
    }
    if (event.key === 'ArrowUp') {
      event.preventDefault();
      event.stopPropagation();
      const current = currentIndex(itemRefs.current);
      focusAt(current < 0 ? 0 : current, -1);
      return;
    }
    if (event.key === 'Home') {
      event.preventDefault();
      event.stopPropagation();
      focusAt(-1, 1);
      return;
    }
    if (event.key === 'End') {
      event.preventDefault();
      event.stopPropagation();
      focusAt(0, -1);
      return;
    }
    if (event.key === 'Enter' || event.key === ' ') {
      const index = currentIndex(itemRefs.current);
      const item = index < 0 ? undefined : items[index];
      if (item === undefined || item.disabled === true) return;
      event.preventDefault();
      event.stopPropagation();
      onSelect(item);
    }
  };

  return (
    <div
      ref={setRoot}
      className={className === undefined ? 'menu' : `menu ${className}`}
      role="menu"
      aria-label={label}
      style={style}
      onKeyDown={onKeyDown}
    >
      {heading === undefined ? null : (
        <div className="menu-heading" role="presentation">
          {heading}
        </div>
      )}
      {items.map((item, index) => (
        <MenuRow
          key={item.id}
          item={item}
          itemRef={(element) => {
            itemRefs.current[index] = element;
          }}
          onSelect={onSelect}
        />
      ))}
    </div>
  );
}

interface MenuRowProps {
  readonly item: MenuItemModel;
  readonly itemRef: (element: HTMLButtonElement | null) => void;
  readonly onSelect: (item: MenuItemModel) => void;
}

function MenuRow({ item, itemRef, onSelect }: MenuRowProps): JSX.Element {
  return (
    <>
      {item.dividerBefore === true ? <div className="menu-separator" role="separator" /> : null}
      <button
        ref={itemRef}
        type="button"
        role={item.checked === undefined ? 'menuitem' : 'menuitemcheckbox'}
        aria-checked={item.checked === undefined ? undefined : item.checked}
        className="menu-item"
        disabled={item.disabled === true}
        data-command={item.commandId}
        data-testid={item.testId}
        onClick={() => {
          if (item.disabled === true) return;
          onSelect(item);
        }}
      >
        <span className="menu-check" aria-hidden="true">
          {item.checked === true ? '✓' : ''}
        </span>
        <span className="menu-label">{item.label}</span>
        {item.shortcut === undefined ? null : (
          <span className="menu-shortcut" aria-hidden="true">
            {item.shortcut}
          </span>
        )}
      </button>
    </>
  );
}

function currentIndex(items: readonly (HTMLButtonElement | null)[]): number {
  const active = document.activeElement;
  return items.findIndex((item) => item === active);
}
