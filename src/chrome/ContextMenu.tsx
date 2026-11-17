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

/**
 * A context menu is a Menu at a pointer position. It keeps the same keyboard
 * contract as every other menu, so a pointer-opened menu stays usable without
 * a mouse.
 */
export function ContextMenu({ x, y, items, onSelect, onClose }: ContextMenuProps): JSX.Element {
  const viewportWidth = typeof window === 'undefined' ? 1024 : window.innerWidth;
  const viewportHeight = typeof window === 'undefined' ? 768 : window.innerHeight;
  const left = Math.max(8, Math.min(x, viewportWidth - 240));
  const top = Math.max(8, Math.min(y, viewportHeight - 48));
  return (
    <Menu
      items={items}
      onSelect={onSelect}
      onClose={onClose}
      label="Context menu"
      className="context-menu"
      style={{ left: `${left}px`, top: `${top}px` }}
    />
  );
}
