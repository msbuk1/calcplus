import { describe, expect, it } from 'vitest';
import { clampToViewport, placeAnchoredPanel, VIEWPORT_MARGIN } from '@chrome/viewport-placement';

/**
 * Floating chrome measures itself and then shifts inside the window. The
 * fallback viewport used outside a browser is 1024 by 768, so these cases are
 * deterministic in the node test environment.
 */

describe('viewport placement', () => {
  it('clamps a measured box inside every edge', () => {
    expect(clampToViewport(1000, 700, 240, 200)).toEqual({ left: 776, top: 560 });
    expect(clampToViewport(-20, -20, 240, 200)).toEqual({
      left: VIEWPORT_MARGIN,
      top: VIEWPORT_MARGIN,
    });
  });

  it('keeps an anchored panel below when that fits', () => {
    expect(placeAnchoredPanel(100, 120, 40, 200, 300)).toEqual({
      left: 40,
      top: 100,
      flipped: false,
    });
  });

  it('flips a panel above when below overflows and above has room', () => {
    expect(placeAnchoredPanel(500, 480, 40, 200, 300)).toEqual({
      left: 40,
      top: 180,
      flipped: true,
    });
  });

  it('clamps when the panel fits on neither side of the anchor', () => {
    expect(placeAnchoredPanel(400, 420, 1000, 200, 800)).toEqual({
      left: 816,
      top: VIEWPORT_MARGIN,
      flipped: false,
    });
  });
});
