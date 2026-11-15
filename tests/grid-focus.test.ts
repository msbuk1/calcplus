import { describe, expect, it } from 'vitest';
import { FOCUS_ZONES, focusAfterEscape, nextFocusZone } from '@chrome/focus';

describe('focus zones', () => {
  it('cycles the zones with F6 and returns to the grid with Esc', () => {
    expect(FOCUS_ZONES).toEqual(['name-box', 'formula-bar', 'grid']);
    expect(nextFocusZone('grid', false)).toBe('name-box');
    expect(nextFocusZone('name-box', false)).toBe('formula-bar');
    expect(nextFocusZone('formula-bar', false)).toBe('grid');
    expect(nextFocusZone('grid', true)).toBe('formula-bar');
    expect(nextFocusZone('name-box', true)).toBe('grid');
    expect(focusAfterEscape()).toBe('grid');
  });
});
