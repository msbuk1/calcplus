import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { COMMAND_INVENTORY } from '@commands/inventory';
import {
  RIBBON_ICONS,
  RIBBON_ICON_NAMES,
  RIBBON_ICON_SIZE,
  isRibbonIconName,
} from '@chrome/ribbon-icons';
import { RIBBON_TABS, compactControls, ribbonGroupsFor } from '@chrome/ribbon-model';

/**
 * Ticket 33: every ribbon command carries original artwork, every rendered
 * control renders it, and the state tokens keep icons at 3:1.
 */
describe('the original ribbon icon family', () => {
  it('draws every icon on the same 16 x 16 canvas with vector-only content', () => {
    expect(RIBBON_ICON_SIZE).toBe(16);
    expect(new Set(RIBBON_ICON_NAMES).size).toBe(RIBBON_ICON_NAMES.length);
    for (const name of RIBBON_ICON_NAMES) {
      const icon = RIBBON_ICONS[name];
      const shapes =
        icon.paths.length +
        (icon.filled?.length ?? 0) +
        (icon.dashed?.length ?? 0) +
        (icon.muted?.length ?? 0) +
        (icon.circles?.length ?? 0);
      expect(shapes, name).toBeGreaterThan(0);
      for (const path of [
        ...icon.paths,
        ...(icon.filled ?? []),
        ...(icon.dashed ?? []),
        ...(icon.muted ?? []),
      ]) {
        expect(path.length, name).toBeGreaterThan(0);
        // Artwork is line/shape data, never a raster or CSS colour.
        expect(path, name).not.toMatch(/#[0-9a-f]{3,8}|rgb\(|hsl\(|url\(/i);
      }
      for (const circle of icon.circles ?? []) {
        expect(circle.radius, name).toBeGreaterThan(0);
        expect(circle.cx, name).toBeGreaterThanOrEqual(0);
        expect(circle.cx, name).toBeLessThanOrEqual(RIBBON_ICON_SIZE);
        expect(circle.cy, name).toBeGreaterThanOrEqual(0);
        expect(circle.cy, name).toBeLessThanOrEqual(RIBBON_ICON_SIZE);
      }
    }
  });

  it('covers every command with a ribbon placement', () => {
    for (const entry of COMMAND_INVENTORY) {
      if ((entry.ribbon ?? []).length === 0) continue;
      expect(entry.icon, entry.id).toBeDefined();
      expect(isRibbonIconName(entry.icon ?? ''), entry.id).toBe(true);
      for (const [index, placement] of (entry.ribbon ?? []).entries()) {
        if (placement.kind === 'launcher' || placement.menuOnly === true) continue;
        expect(placement.icon, `${entry.id} placement ${index}`).toBeDefined();
        expect(isRibbonIconName(placement.icon ?? ''), `${entry.id} placement ${index}`).toBe(true);
      }
    }
  });

  it('uses every artwork in the family', () => {
    const used = new Set<string>();
    for (const entry of COMMAND_INVENTORY) {
      if (entry.icon !== undefined) used.add(entry.icon);
      for (const placement of entry.ribbon ?? []) {
        if (placement.icon !== undefined) used.add(placement.icon);
        for (const item of placement.menu ?? []) {
          if (item.icon !== undefined) used.add(item.icon);
        }
      }
      if (entry.compact?.icon !== undefined) used.add(entry.compact.icon);
      if (entry.quickAccess?.icon !== undefined) used.add(entry.quickAccess.icon);
    }
    expect([...used].sort()).toEqual([...RIBBON_ICON_NAMES].sort());
  });

  it('keeps the six alignment commands a distinct icon family', () => {
    const alignment = COMMAND_INVENTORY.find((entry) => entry.id === 'format.alignment');
    expect(alignment?.ribbon?.map((placement) => placement.icon)).toEqual([
      'align-left',
      'align-centre',
      'align-right',
      'align-top',
      'align-middle',
      'align-bottom',
    ]);
    expect(alignment?.ribbon?.every((placement) => placement.iconOnly === true)).toBe(true);
  });

  it('renders an icon on every full-ribbon control', () => {
    for (const tab of RIBBON_TABS) {
      for (const group of ribbonGroupsFor(tab.id)) {
        for (const control of group.controls) {
          expect(control.icon, `${tab.id} ${group.id} ${control.commandId}`).toBeDefined();
          expect(isRibbonIconName(control.icon ?? '')).toBe(true);
          // The control owns the accessible name; the SVG stays decorative.
          expect(
            control.title.length,
            `${tab.id} ${group.id} ${control.commandId}`,
          ).toBeGreaterThan(0);
        }
      }
    }
  });

  it('renders an icon on every compact control', () => {
    for (const control of compactControls()) {
      expect(control.icon, control.commandId).toBeDefined();
      expect(isRibbonIconName(control.icon ?? '')).toBe(true);
      expect(control.title.length, control.commandId).toBeGreaterThan(0);
    }
  });
});

describe('ribbon icon state contrast', () => {
  const css = readFileSync(new URL('../src/styles.css', import.meta.url), 'utf8');

  function token(name: string): string {
    const match = new RegExp(`--${name}:\\s*(#[0-9a-fA-F]{6})`).exec(css);
    if (match === null) throw new Error(`styles.css is missing --${name}`);
    return match[1]!;
  }

  function channel(hex: string, offset: number): number {
    const value = Number.parseInt(hex.slice(offset, offset + 2), 16) / 255;
    return value <= 0.04045 ? value / 12.92 : ((value + 0.055) / 1.055) ** 2.4;
  }

  function luminance(hex: string): number {
    return 0.2126 * channel(hex, 1) + 0.7152 * channel(hex, 3) + 0.0722 * channel(hex, 5);
  }

  function contrast(first: string, second: string): number {
    const values = [luminance(first), luminance(second)].sort((left, right) => right - left);
    return (values[0]! + 0.05) / (values[1]! + 0.05);
  }

  it('holds at least 3:1 in default, hover, pressed, disabled and checked states', () => {
    const pairs: readonly (readonly [string, string])[] = [
      ['text', 'chrome-surface'],
      ['text', 'command-surface'],
      ['text', 'button-hover'],
      ['text', 'button-pressed'],
      ['icon-disabled', 'chrome-surface'],
      ['icon-disabled', 'command-surface'],
      ['icon-disabled', 'accent-soft'],
      ['accent-hover', 'accent-soft'],
    ];
    for (const [foreground, background] of pairs) {
      expect(
        contrast(token(foreground), token(background)),
        `${foreground} on ${background}`,
      ).toBeGreaterThanOrEqual(3);
    }
  });
});
