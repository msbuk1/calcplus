import { describe, expect, it } from 'vitest';
import { SparseAxisGeometry, columnWidthCharsToPixels, pointsToPixels } from '@view/geometry';

describe('sparse axis geometry', () => {
  it('lays out a uniform axis by index', () => {
    const geometry = new SparseAxisGeometry(1_048_576, 20, new Map());
    expect(geometry.sizeAt(0)).toBe(20);
    expect(geometry.offsetOf(0)).toBe(0);
    expect(geometry.offsetOf(10)).toBe(200);
    expect(geometry.totalSize()).toBe(1_048_576 * 20);
    expect(geometry.indexAt(0)).toBe(0);
    expect(geometry.indexAt(20)).toBe(1);
    expect(geometry.indexAt(219)).toBe(10);
    expect(geometry.indexAt(220)).toBe(11);
  });

  it('honours sparse size overrides without allocating the axis', () => {
    const overrides = new Map([
      [5, 30],
      [9, 10],
    ]);
    const geometry = new SparseAxisGeometry(100, 20, overrides);
    expect(geometry.sizeAt(5)).toBe(30);
    expect(geometry.sizeAt(9)).toBe(10);
    expect(geometry.sizeAt(6)).toBe(20);
    // The overrides shift every later offset by their difference.
    expect(geometry.offsetOf(6)).toBe(5 * 20 + 30);
    expect(geometry.offsetOf(10)).toBe(10 * 20);
    expect(geometry.totalSize()).toBe(100 * 20 + 10 - 10);
  });

  it('finds the index that owns an offset across overrides', () => {
    const overrides = new Map([
      [5, 30],
      [9, 10],
    ]);
    const geometry = new SparseAxisGeometry(100, 20, overrides);
    expect(geometry.indexAt(99)).toBe(4);
    expect(geometry.indexAt(100)).toBe(5);
    expect(geometry.indexAt(129)).toBe(5);
    expect(geometry.indexAt(130)).toBe(6);
    expect(geometry.indexAt(185)).toBe(8);
    expect(geometry.indexAt(195)).toBe(9);
    expect(geometry.indexAt(205)).toBe(10);
  });

  it('keeps a million-row axis addressable with constant memory', () => {
    const overrides = new Map([[999_999, 40]]);
    const geometry = new SparseAxisGeometry(1_048_576, 20, overrides);
    expect(geometry.offsetOf(1_048_575)).toBe(1_048_575 * 20 + 20);
    expect(geometry.indexAt(1_048_575 * 20)).toBe(1_048_574);
    expect(geometry.indexAt(geometry.totalSize() - 1)).toBe(1_048_575);
    expect(geometry.sizeAt(999_999)).toBe(40);
  });

  it('converts Excel row heights and column widths to pixels', () => {
    expect(pointsToPixels(15)).toBe(20);
    expect(pointsToPixels(22)).toBe(29);
    expect(columnWidthCharsToPixels(8.43)).toBe(64);
    expect(columnWidthCharsToPixels(18.5)).toBe(135);
  });
});
