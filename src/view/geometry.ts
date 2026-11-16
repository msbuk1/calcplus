/**
 * Sparse row and column geometry.
 *
 * The grid is always 1,048,576 rows by 16,384 columns, but only a handful of
 * rows or columns carry an imported size. The geometry therefore stores the
 * default size plus the sparse overrides and answers every question by index,
 * never by iteration (workbook-model §11).
 */

/** Excel row heights are points; the canvas works in CSS pixels at 96 dpi. */
export function pointsToPixels(points: number): number {
  return Math.round((points * 96) / 72);
}

/** Excel column widths are character units; the canvas works in CSS pixels. */
export function columnWidthCharsToPixels(chars: number): number {
  return Math.round(chars * 7 + 5);
}

/** The inverse of `columnWidthCharsToPixels`, for autofit and resize. */
export function columnWidthPixelsToChars(pixels: number): number {
  return Math.max(0, (pixels - 5) / 7);
}

/** The inverse of `pointsToPixels`, for row autofit and resize. */
export function pixelsToPoints(pixels: number): number {
  return (pixels * 72) / 96;
}

export class SparseAxisGeometry {
  readonly count: number;
  readonly defaultSize: number;
  private readonly sizes: number[] = [];
  private readonly indices: number[] = [];
  /** `prefixDeltas[k]` is the summed size difference of the first k overrides. */
  private readonly prefixDeltas: number[] = [0];

  constructor(count: number, defaultSize: number, overrides: ReadonlyMap<number, number>) {
    this.count = count;
    this.defaultSize = defaultSize;
    // A zero-size override is a hidden row or column: its offset collapses to
    // the next visible index without allocating anything.
    const entries = [...overrides.entries()]
      .filter(([index, size]) => index >= 0 && index < count && size >= 0)
      .sort((left, right) => left[0] - right[0]);
    let delta = 0;
    for (const [index, size] of entries) {
      this.indices.push(index);
      this.sizes.push(size);
      delta += size - defaultSize;
      this.prefixDeltas.push(delta);
    }
  }

  /** The size of one row or column. */
  sizeAt(index: number): number {
    if (index < 0 || index >= this.count) return this.defaultSize;
    const position = this.indexOfOverride(index);
    if (position < 0) return this.defaultSize;
    return this.sizes[position] ?? this.defaultSize;
  }

  /** The offset of the start of a row or column, measured from the sheet edge. */
  offsetOf(index: number): number {
    if (index <= 0) return 0;
    if (index >= this.count) return this.totalSize();
    return index * this.defaultSize + this.deltaBefore(index);
  }

  /** The total size of the axis. */
  totalSize(): number {
    return this.count * this.defaultSize + (this.prefixDeltas[this.prefixDeltas.length - 1] ?? 0);
  }

  /** The row or column that contains an offset. */
  indexAt(offset: number): number {
    if (this.count <= 0) return 0;
    // A hidden row or column has zero size, so the offset at its start belongs
    // to the next visible index. Never hand a click or an arrow key a hidden
    // boundary index.
    if (offset <= 0) return this.firstVisibleIndex();
    const total = this.totalSize();
    if (offset >= total) return this.lastVisibleIndex();
    // Binary search for the last override that starts at or before the offset.
    let low = 0;
    let high = this.indices.length - 1;
    let found = -1;
    while (low <= high) {
      const middle = (low + high) >> 1;
      if (this.startOffsetOfOverride(middle) <= offset) {
        found = middle;
        low = middle + 1;
      } else {
        high = middle - 1;
      }
    }
    const candidate = (() => {
      if (found < 0) return Math.min(this.count - 1, Math.floor(offset / this.defaultSize));
      const start = this.startOffsetOfOverride(found);
      const size = this.sizes[found] ?? this.defaultSize;
      const relative = offset - start;
      if (relative < size) return this.indices[found] ?? 0;
      const rest = relative - size;
      const index = (this.indices[found] ?? 0) + 1 + Math.floor(rest / this.defaultSize);
      return Math.min(this.count - 1, Math.max(0, index));
    })();
    return this.visibleAtOrAfter(candidate);
  }

  private visibleAtOrAfter(index: number): number {
    let target = index;
    while (target < this.count && this.sizeAt(target) <= 0) target += 1;
    return target < this.count ? target : this.lastVisibleIndex();
  }

  private firstVisibleIndex(): number {
    let index = 0;
    while (index < this.count && this.sizeAt(index) <= 0) index += 1;
    return index < this.count ? index : Math.max(0, this.count - 1);
  }

  private lastVisibleIndex(): number {
    let index = this.count - 1;
    while (index >= 0 && this.sizeAt(index) <= 0) index -= 1;
    return index >= 0 ? index : 0;
  }

  private deltaBefore(index: number): number {
    const position = this.lowerBound(index);
    return this.prefixDeltas[position] ?? 0;
  }

  /** The number of overrides with an index strictly below `index`. */
  private lowerBound(index: number): number {
    let low = 0;
    let high = this.indices.length;
    while (low < high) {
      const middle = (low + high) >> 1;
      if ((this.indices[middle] ?? 0) < index) low = middle + 1;
      else high = middle;
    }
    return low;
  }

  private indexOfOverride(index: number): number {
    const position = this.lowerBound(index + 1);
    if (position === 0) return -1;
    if (this.indices[position - 1] !== index) return -1;
    return position - 1;
  }

  private startOffsetOfOverride(position: number): number {
    const index = this.indices[position] ?? 0;
    return index * this.defaultSize + (this.prefixDeltas[position] ?? 0);
  }
}
