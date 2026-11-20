import { describe, expect, it } from 'vitest';
import {
  FILL_PATTERNS,
  PATTERN_PREVIEW_BACKGROUND,
  fillPatternPreview,
  patternTileSvg,
  patternTileUrl,
} from '@chrome/fill-patterns';

/**
 * Ticket 39: the Fill tab's Pattern control previews every pattern in the
 * dialog's colours. The previews are pure data (an inline SVG tile plus two CSS
 * colours), so the whole contract is checked here without a browser; the
 * Playwright format-dialog spec covers the rendered control.
 */

const TEXTURED = FILL_PATTERNS.filter(
  (pattern) => pattern.value !== 'none' && pattern.value !== 'solid',
);

function decodeTile(url: string | null): string {
  if (url === null) throw new Error('expected a pattern tile URL');
  const match = /^url\("data:image\/svg\+xml,(.*)"\)$/.exec(url);
  if (match === null) throw new Error(`not an inline SVG data URL: ${url}`);
  return decodeURIComponent(match[1]!);
}

describe('fill pattern previews', () => {
  it('offers every pattern the grid paints, once, with a distinct drawn tile', () => {
    expect(FILL_PATTERNS.map((pattern) => pattern.value)).toEqual([
      'none',
      'solid',
      'gray125',
      'gray0625',
      'lightGray',
      'mediumGray',
      'darkGray',
      'lightHorizontal',
      'darkHorizontal',
      'lightVertical',
      'darkVertical',
      'lightDown',
      'darkDown',
      'lightUp',
      'darkUp',
      'lightGrid',
      'darkGrid',
      'lightTrellis',
      'darkTrellis',
    ]);
    expect(new Set(FILL_PATTERNS.map((pattern) => pattern.label)).size).toBe(FILL_PATTERNS.length);

    const tiles = TEXTURED.map((pattern) => patternTileSvg(pattern.value, '#C00000'));
    for (const tile of tiles) {
      expect(tile).not.toBeNull();
      expect(tile).toContain("width='8' height='8' viewBox='0 0 8 8'");
      expect(tile).toContain('#C00000');
    }
    // No two patterns may fall back to the same drawing.
    expect(new Set(tiles).size).toBe(TEXTURED.length);
  });

  it('draws solid as the fill colour and none as a blank swatch', () => {
    expect(fillPatternPreview('solid', '#C00000', '#EAF1F8')).toEqual({
      backgroundColor: '#C00000',
      backgroundImage: null,
    });
    // A solid fill with no colour paints nothing, so the preview stays white.
    expect(fillPatternPreview('solid', null, '#EAF1F8')).toEqual({
      backgroundColor: PATTERN_PREVIEW_BACKGROUND,
      backgroundImage: null,
    });
    expect(fillPatternPreview('none', '#C00000', '#EAF1F8')).toEqual({
      backgroundColor: PATTERN_PREVIEW_BACKGROUND,
      backgroundImage: null,
    });
  });

  it('puts the texture in the fill colour over the pattern background', () => {
    const preview = fillPatternPreview('lightGrid', '#C00000', '#EAF1F8');
    expect(preview.backgroundColor).toBe('#EAF1F8');
    const tile = decodeTile(preview.backgroundImage);
    expect(tile).toContain('#C00000');
    expect(tile).not.toContain('#EAF1F8');

    // No pattern background: the grid paints white behind the texture.
    expect(fillPatternPreview('lightGrid', '#C00000', null).backgroundColor).toBe(
      PATTERN_PREVIEW_BACKGROUND,
    );
    // No fill colour: only the background shows, exactly as the grid paints it.
    expect(fillPatternPreview('lightGrid', null, '#EAF1F8')).toEqual({
      backgroundColor: '#EAF1F8',
      backgroundImage: null,
    });
  });

  it('tiles each preview as an inline SVG with no external asset', () => {
    const url = patternTileUrl('darkTrellis', '#123456');
    expect(url).not.toBeNull();
    expect(url).toContain('url("data:image/svg+xml,');
    // The hash is percent-encoded so the data URL survives a CSS declaration.
    expect(url).not.toContain('#');
    expect(decodeTile(url)).toContain('#123456');
  });

  it('ignores a colour that is not a plain six-digit hex', () => {
    for (const hostile of ['red', '#12345', '#GGGGGG', "#'/><script>alert(1)</script>"]) {
      expect(patternTileSvg('lightGrid', hostile)).toBeNull();
      expect(fillPatternPreview('lightGrid', hostile, null)).toEqual({
        backgroundColor: PATTERN_PREVIEW_BACKGROUND,
        backgroundImage: null,
      });
    }
  });
});
