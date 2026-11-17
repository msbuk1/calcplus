import type { JSX } from 'react';
import type { SelectionStatistics } from '@view/types';

export interface StatusBarProps {
  /** The transient status message; the ready state sits beside it. */
  readonly status: string;
  readonly activeCell: string;
  readonly activeValue: string;
  readonly stale: boolean;
  readonly staleCount: number;
  /** `null` when the selection holds no numeric cell. */
  readonly statistics: SelectionStatistics | null;
  readonly showGridLines: boolean;
  readonly showHeadings: boolean;
  /** Display zoom as an Excel percentage, 10-400. */
  readonly zoom: number;
  readonly onToggleGridLines: () => void;
  readonly onToggleHeadings: () => void;
  readonly onZoomChange: (zoomPercent: number) => void;
}

const ZOOM_PRESETS: readonly number[] = [10, 25, 50, 75, 100, 125, 150, 175, 200, 250, 300, 400];

/**
 * The status bar: ready state, live selection statistics and the view
 * controls. Every view control here changes transient view state only; none
 * of them dispatches a document command or marks the workbook dirty.
 */
export function StatusBar({
  status,
  activeCell,
  activeValue,
  stale,
  staleCount,
  statistics,
  showGridLines,
  showHeadings,
  zoom,
  onToggleGridLines,
  onToggleHeadings,
  onZoomChange,
}: StatusBarProps): JSX.Element {
  const zoomOptions = [...new Set([...ZOOM_PRESETS, Math.round(zoom)])].sort(
    (left, right) => left - right,
  );
  return (
    <footer className="status-bar">
      <span className="status-ready" data-testid="status-ready">
        Ready
      </span>
      <span className="status-state" data-testid="status-message">
        {status}
      </span>
      <span data-testid="active-cell">{activeCell}</span>
      <span data-testid="active-value">{activeValue}</span>
      {stale ? (
        <span className="stale-marker" data-testid="active-stale">
          not recalculated
        </span>
      ) : null}
      {statistics !== null ? (
        <span className="selection-statistics" data-testid="selection-statistics">
          <span data-testid="selection-sum">Sum: {formatStatistic(statistics.sum)}</span>
          <span data-testid="selection-average">
            Average: {formatStatistic(statistics.average)}
          </span>
          <span data-testid="selection-count">Count: {statistics.count}</span>
        </span>
      ) : null}
      <span className="status-spacer" />
      {staleCount > 0 ? (
        <span
          className="stale-marker"
          data-testid="stale-count"
          title="Cells whose cached result was kept because it could not be recalculated"
        >
          {staleCount} not recalculated
        </span>
      ) : null}
      <label className="status-toggle">
        <input
          type="checkbox"
          checked={showGridLines}
          onChange={onToggleGridLines}
          data-testid="gridlines-toggle"
        />
        Gridlines
      </label>
      <label className="status-toggle">
        <input
          type="checkbox"
          checked={showHeadings}
          onChange={onToggleHeadings}
          data-testid="headings-toggle"
        />
        Headings
      </label>
      <div className="status-zoom" role="group" aria-label="Zoom">
        <select
          className="zoom-menu"
          aria-label="Zoom level"
          value={String(Math.round(zoom))}
          onChange={(event) => onZoomChange(Number(event.target.value))}
          data-testid="zoom-menu"
        >
          {zoomOptions.map((option) => (
            <option key={option} value={String(option)}>
              {option}%
            </option>
          ))}
        </select>
        <input
          className="zoom-slider"
          type="range"
          min={10}
          max={400}
          step={1}
          value={zoom}
          aria-label="Zoom slider"
          onChange={(event) => onZoomChange(Number(event.target.value))}
          data-testid="zoom-slider"
        />
        <span className="zoom-value" data-testid="zoom-value">
          {Math.round(zoom)}%
        </span>
      </div>
    </footer>
  );
}

const numberFormatter = new Intl.NumberFormat(undefined, {
  maximumFractionDigits: 10,
});

function formatStatistic(value: number): string {
  if (!Number.isFinite(value)) return String(value);
  return numberFormatter.format(value);
}
