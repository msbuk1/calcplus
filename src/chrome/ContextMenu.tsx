import { useLayoutEffect, useRef, useState } from 'react';
import type { JSX } from 'react';
import { Menu } from './Menu';
import type { MenuItemModel } from './Menu';

export interface ContextMenuProps {
  readonly x: number;
  readonly y: number;
  readonly items: readonly MenuItemModel[];
  readonly onSelect: (item: MenuItemModel) => void;
  readonly onClose: () => void;
}

const VIEWPORT_MARGIN = 8;

function clampToViewport(
  x: number,
  y: number,
  width: number,
  height: number,
): { left: number; top: number } {
  const viewportWidth = typeof window === 'undefined' ? 1024 : window.innerWidth;
  const viewportHeight = typeof window === 'undefined' ? 768 : window.innerHeight;
  return {
    left: Math.max(VIEWPORT_MARGIN, Math.min(x, viewportWidth - VIEWPORT_MARGIN - width)),
    top: Math.max(VIEWPORT_MARGIN, Math.min(y, viewportHeight - VIEWPORT_MARGIN - height)),
  };
}

/**
 * A context menu is a Menu at a pointer position. It keeps the same keyboard
 * contract as every other menu, so a pointer-opened menu stays usable without
 * a mouse. The menu is clamped to the viewport: it opens at the pointer and
 * shifts up or left when its measured size would run past the bottom or right
 * edge, so the last items stay clickable on short windows.
 */
export function ContextMenu({ x, y, items, onSelect, onClose }: ContextMenuProps): JSX.Element {
  const menuRef = useRef<HTMLDivElement | null>(null);
  // The first paint uses a rough size; the layout effect corrects it before
  // the browser paints, so the estimate is never visible.
  const [position, setPosition] = useState(() => clampToViewport(x, y, 240, 48));

  useLayoutEffect(() => {
    const menu = menuRef.current;
    if (menu === null) return;
    const rect = menu.getBoundingClientRect();
    setPosition(clampToViewport(x, y, rect.width, rect.height));
  }, [x, y]);

  return (
    <Menu
      items={items}
      onSelect={onSelect}
      onClose={onClose}
      label="Context menu"
      className="context-menu"
      rootRef={menuRef}
      style={{ left: `${position.left}px`, top: `${position.top}px` }}
    />
  );
}
