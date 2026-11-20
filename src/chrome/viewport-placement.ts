/**
 * Viewport placement for floating chrome surfaces. Menus and the AutoFilter
 * dropdown both measure themselves and shift inside the window; keeping the
 * arithmetic here means the two overlays cannot drift apart.
 */

/** Space kept between an overlay and the window edge, in CSS pixels. */
export const VIEWPORT_MARGIN = 8;

function viewportSize(): { readonly width: number; readonly height: number } {
  return typeof window === 'undefined'
    ? { width: 1024, height: 768 }
    : { width: window.innerWidth, height: window.innerHeight };
}

/** Shift a box so every edge stays inside the window. */
export function clampToViewport(
  x: number,
  y: number,
  width: number,
  height: number,
): { readonly left: number; readonly top: number } {
  const viewport = viewportSize();
  return {
    left: Math.max(VIEWPORT_MARGIN, Math.min(x, viewport.width - VIEWPORT_MARGIN - width)),
    top: Math.max(VIEWPORT_MARGIN, Math.min(y, viewport.height - VIEWPORT_MARGIN - height)),
  };
}

export interface AnchoredPanelPlacement {
  readonly left: number;
  readonly top: number;
  /** `true` when the panel was flipped above its anchor to stay on screen. */
  readonly flipped: boolean;
}

/**
 * Place a panel whose top-left would sit below an anchor. When the panel would
 * run past the bottom edge and there is room above, it flips to sit above the
 * anchor. If neither direction has room, the panel is clamped to the window;
 * the caller lets its scrolling body shrink so a footer stays visible.
 */
export function placeAnchoredPanel(
  anchorTop: number,
  anchorBottom: number,
  anchorLeft: number,
  width: number,
  height: number,
): AnchoredPanelPlacement {
  const viewport = viewportSize();
  const left = Math.max(
    VIEWPORT_MARGIN,
    Math.min(anchorLeft, viewport.width - VIEWPORT_MARGIN - width),
  );
  if (anchorTop + height <= viewport.height - VIEWPORT_MARGIN) {
    return { left, top: anchorTop, flipped: false };
  }
  const aboveTop = anchorBottom - height;
  if (aboveTop >= VIEWPORT_MARGIN) {
    return { left, top: aboveTop, flipped: true };
  }
  const top = Math.max(
    VIEWPORT_MARGIN,
    Math.min(anchorTop, viewport.height - VIEWPORT_MARGIN - height),
  );
  return { left, top, flipped: false };
}
