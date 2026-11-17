/**
 * Focus zones (command-system §7). The grid, the formula bar, the name box and
 * the ribbon own keyboard input; the focus manager routes events by zone.
 */
export type FocusZone = 'name-box' | 'formula-bar' | 'grid' | 'ribbon';

export const FOCUS_ZONES: readonly FocusZone[] = ['name-box', 'formula-bar', 'grid', 'ribbon'];

/** F6 moves to the next zone; Shift+F6 moves back. */
export function nextFocusZone(current: FocusZone, backwards: boolean): FocusZone {
  const index = FOCUS_ZONES.indexOf(current);
  if (index < 0) return FOCUS_ZONES[0] ?? 'grid';
  const offset = backwards ? -1 : 1;
  const next = (index + offset + FOCUS_ZONES.length) % FOCUS_ZONES.length;
  return FOCUS_ZONES[next] ?? 'grid';
}

/** Esc always returns to the grid. */
export function focusAfterEscape(): FocusZone {
  return 'grid';
}
