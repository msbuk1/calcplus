import { useEffect, useLayoutEffect, useRef, useState } from 'react';
import type { JSX, KeyboardEvent as ReactKeyboardEvent, RefObject } from 'react';
import type { CSSProperties } from 'react';
import { RibbonIcon } from './RibbonIcon';
import type { RibbonIconName } from './ribbon-icons';

export interface MenuItemModel {
  readonly id: string;
  readonly label: string;
  readonly commandId?: string;
  readonly args?: unknown;
  readonly shortcut?: string;
  /** A name from `src/chrome/ribbon-icons.ts`, drawn ahead of the label. */
  readonly icon?: RibbonIconName;
  readonly disabled?: boolean;
  readonly checked?: boolean;
  readonly dividerBefore?: boolean;
  readonly testId?: string;
  /**
   * A one-level flyout, as Excel's Insert and Delete context-menu entries show.
   * Child items are leaves: they never carry items of their own.
   */
  readonly items?: readonly MenuItemModel[];
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

const EMPTY_ITEMS: readonly MenuItemModel[] = [];
const VIEWPORT_MARGIN = 8;

/**
 * One keyboard-navigable menu for the File menu, ribbon split buttons and the
 * grid context menu. Arrow keys move, Enter or Space activates, Esc closes.
 * An item with `items` opens a one-level flyout: Right (or Enter) opens it,
 * Left or Esc closes it back to its parent item. The menu never dispatches
 * anything itself; it hands the item to its owner.
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
  const flyoutRef = useRef<HTMLDivElement | null>(null);
  const flyoutItemRefs = useRef<(HTMLButtonElement | null)[]>([]);
  const focusFlyoutOnOpen = useRef(false);
  const [openIndex, setOpenIndex] = useState<number | null>(null);
  const [flyoutPosition, setFlyoutPosition] = useState<{
    readonly left: number;
    readonly top: number;
  } | null>(null);
  const setRoot = (node: HTMLDivElement | null): void => {
    menuRef.current = node;
    if (rootRef !== undefined) rootRef.current = node;
  };

  const openItem = openIndex === null ? undefined : items[openIndex];
  const flyoutItems = openItem?.items ?? EMPTY_ITEMS;

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

  /**
   * Place the flyout beside its parent row and clamp it to the viewport. The
   * panel is fixed so the menu's own scroll container cannot clip it. The first
   * paint is hidden fully off-screen; the layout effect corrects it before the
   * browser paints.
   */
  useLayoutEffect(() => {
    if (openIndex === null) {
      setFlyoutPosition(null);
      return;
    }
    const anchor = itemRefs.current[openIndex];
    if (anchor === null) return;
    const anchorRect = anchor.getBoundingClientRect();
    const panelRect = flyoutRef.current?.getBoundingClientRect();
    const width = panelRect?.width ?? 220;
    const height = panelRect?.height ?? 0;
    const viewportWidth = window.innerWidth;
    const viewportHeight = window.innerHeight;
    let left = anchorRect.right + 2;
    if (left + width > viewportWidth - VIEWPORT_MARGIN) {
      left = Math.max(VIEWPORT_MARGIN, anchorRect.left - width - 2);
    }
    let top = anchorRect.top - 4;
    if (height > 0 && top + height > viewportHeight - VIEWPORT_MARGIN) {
      top = Math.max(VIEWPORT_MARGIN, viewportHeight - VIEWPORT_MARGIN - height);
    }
    setFlyoutPosition((current) =>
      current !== null && current.left === left && current.top === top ? current : { left, top },
    );
    if (focusFlyoutOnOpen.current) {
      focusFlyoutOnOpen.current = false;
      const first = flyoutItems.findIndex((item) => item.disabled !== true);
      if (first >= 0) flyoutItemRefs.current[first]?.focus();
    }
  }, [openIndex, flyoutItems]);

  const focusAt = (
    refs: readonly (HTMLButtonElement | null)[],
    levelItems: readonly MenuItemModel[],
    start: number,
    direction: 1 | -1,
  ): void => {
    if (levelItems.length === 0) return;
    let index = start;
    for (let step = 0; step < levelItems.length; step += 1) {
      index = (index + direction + levelItems.length) % levelItems.length;
      const item = levelItems[index];
      if (item !== undefined && item.disabled !== true) {
        refs[index]?.focus();
        return;
      }
    }
  };

  const focusLast = (
    refs: readonly (HTMLButtonElement | null)[],
    levelItems: readonly MenuItemModel[],
  ): void => {
    focusAt(refs, levelItems, 0, -1);
  };

  const openSubmenu = (index: number, focusFirst: boolean): void => {
    focusFlyoutOnOpen.current = focusFirst;
    setOpenIndex(index);
  };

  const closeSubmenu = (refocusParent?: number): void => {
    setOpenIndex(null);
    setFlyoutPosition(null);
    if (refocusParent !== undefined) itemRefs.current[refocusParent]?.focus();
  };

  const onKeyDown = (event: ReactKeyboardEvent<HTMLDivElement>): void => {
    const parentIndex = currentIndex(itemRefs.current);
    const childIndex = currentIndex(flyoutItemRefs.current);
    const onChild = openIndex !== null && childIndex >= 0;

    if (event.key === 'Escape') {
      event.preventDefault();
      event.stopPropagation();
      if (openIndex !== null) {
        closeSubmenu(openIndex);
        return;
      }
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
    if (event.key === 'ArrowLeft') {
      // One level only: Left backs out of a flyout and never closes the menu.
      if (openIndex !== null) {
        event.preventDefault();
        event.stopPropagation();
        closeSubmenu(openIndex);
      }
      return;
    }
    if (event.key === 'ArrowRight') {
      const item = parentIndex < 0 ? undefined : items[parentIndex];
      if (
        !onChild &&
        item !== undefined &&
        item.disabled !== true &&
        item.items !== undefined &&
        item.items.length > 0
      ) {
        event.preventDefault();
        event.stopPropagation();
        openSubmenu(parentIndex, true);
      }
      return;
    }
    if (event.key === 'ArrowDown') {
      event.preventDefault();
      event.stopPropagation();
      if (onChild) {
        focusAt(flyoutItemRefs.current, flyoutItems, childIndex, 1);
      } else {
        if (openIndex !== null) closeSubmenu();
        focusAt(itemRefs.current, items, parentIndex, 1);
      }
      return;
    }
    if (event.key === 'ArrowUp') {
      event.preventDefault();
      event.stopPropagation();
      if (onChild) {
        focusAt(flyoutItemRefs.current, flyoutItems, childIndex, -1);
      } else {
        if (openIndex !== null) closeSubmenu();
        focusAt(itemRefs.current, items, parentIndex < 0 ? 0 : parentIndex, -1);
      }
      return;
    }
    if (event.key === 'Home') {
      event.preventDefault();
      event.stopPropagation();
      if (onChild) {
        focusAt(flyoutItemRefs.current, flyoutItems, -1, 1);
      } else {
        if (openIndex !== null) closeSubmenu();
        focusAt(itemRefs.current, items, -1, 1);
      }
      return;
    }
    if (event.key === 'End') {
      event.preventDefault();
      event.stopPropagation();
      if (onChild) {
        focusLast(flyoutItemRefs.current, flyoutItems);
      } else {
        if (openIndex !== null) closeSubmenu();
        focusLast(itemRefs.current, items);
      }
      return;
    }
    if (event.key === 'Enter' || event.key === ' ') {
      if (onChild) {
        const child = flyoutItems[childIndex];
        if (child === undefined || child.disabled === true) return;
        event.preventDefault();
        event.stopPropagation();
        onSelect(child);
        return;
      }
      const item = parentIndex < 0 ? undefined : items[parentIndex];
      if (item === undefined || item.disabled === true) return;
      if (item.items !== undefined && item.items.length > 0) {
        event.preventDefault();
        event.stopPropagation();
        if (openIndex === parentIndex) {
          const first = flyoutItems.findIndex((child) => child.disabled !== true);
          if (first >= 0) flyoutItemRefs.current[first]?.focus();
        } else {
          openSubmenu(parentIndex, true);
        }
        return;
      }
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
          expanded={openIndex === index}
          itemRef={(element) => {
            itemRefs.current[index] = element;
          }}
          onActivate={() => {
            if (item.items !== undefined && item.items.length > 0) {
              // A click always offers the flyout; Esc, Left or another row
              // closes it. Hover already opened it in the common path.
              openSubmenu(index, false);
              return;
            }
            onSelect(item);
          }}
          onHover={() => {
            if (item.disabled === true) return;
            if (item.items !== undefined && item.items.length > 0) {
              if (openIndex !== index) openSubmenu(index, false);
              return;
            }
            if (openIndex !== null) closeSubmenu();
          }}
        />
      ))}
      {openIndex === null || flyoutItems.length === 0 ? null : (
        <div
          ref={flyoutRef}
          className="menu context-menu"
          role="menu"
          aria-label={openItem?.label}
          style={{
            left: `${flyoutPosition?.left ?? -9999}px`,
            top: `${flyoutPosition?.top ?? -9999}px`,
          }}
        >
          {flyoutItems.map((item, index) => (
            <MenuRow
              key={item.id}
              item={item}
              itemRef={(element) => {
                flyoutItemRefs.current[index] = element;
              }}
              onActivate={() => onSelect(item)}
              onHover={() => undefined}
            />
          ))}
        </div>
      )}
    </div>
  );
}

interface MenuRowProps {
  readonly item: MenuItemModel;
  readonly itemRef: (element: HTMLButtonElement | null) => void;
  readonly onActivate: () => void;
  readonly onHover: () => void;
  readonly expanded?: boolean;
}

function MenuRow({ item, itemRef, onActivate, onHover, expanded }: MenuRowProps): JSX.Element {
  const hasItems = item.items !== undefined && item.items.length > 0;
  return (
    <>
      {item.dividerBefore === true ? <div className="menu-separator" role="separator" /> : null}
      <button
        ref={itemRef}
        type="button"
        role={item.checked === undefined ? 'menuitem' : 'menuitemcheckbox'}
        aria-checked={item.checked === undefined ? undefined : item.checked}
        aria-haspopup={hasItems ? 'menu' : undefined}
        aria-expanded={hasItems ? expanded === true : undefined}
        className="menu-item"
        disabled={item.disabled === true}
        data-command={item.commandId}
        data-testid={item.testId}
        onClick={() => {
          if (item.disabled === true) return;
          onActivate();
        }}
        onMouseEnter={onHover}
      >
        <span className="menu-check" aria-hidden="true">
          {item.checked === true ? '✓' : ''}
        </span>
        {item.icon === undefined ? null : (
          <span className="menu-icon" aria-hidden="true">
            <RibbonIcon name={item.icon} />
          </span>
        )}
        <span className="menu-label">{item.label}</span>
        {hasItems ? (
          <span className="menu-shortcut" aria-hidden="true">
            ›
          </span>
        ) : item.shortcut === undefined ? null : (
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
