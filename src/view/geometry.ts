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
    const entries = [...overrides.entries()]
      .filter(([index, size]) => index >= 0 && index < count && size > 0)
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
    if (offset <= 0) return 0;
    const total = this.totalSize();
    if (offset >= total) return this.count - 1;
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
    if (found < 0) return Math.min(this.count - 1, Math.floor(offset / this.defaultSize));
    const start = this.startOffsetOfOverride(found);
    const size = this.sizes[found] ?? this.defaultSize;
    const relative = offset - start;
    if (relative < size) return this.indices[found] ?? 0;
    const rest = relative - size;
    const index = (this.indices[found] ?? 0) + 1 + Math.floor(rest / this.defaultSize);
    return Math.min(this.count - 1, Math.max(0, index));
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
