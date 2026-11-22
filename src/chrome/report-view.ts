import type {
  CompatibilityReport,
  ExportReport,
  ExportReportCategory,
  ReportItem,
  ReportItemKind,
} from '@ooxml/report';
import { EXPORT_REPORT_CATEGORIES, reportKindLabel } from '@ooxml/report';

/**
 * The dialog views for the two reports (file-compat §7). Pure mapping from the
 * report model to headings and rows, so the React component stays a renderer
 * and the mapping can be tested without a DOM.
 */

export interface ReportEntryView {
  readonly id: string;
  readonly title: string;
  /** The sheet and anchor, when the item names a position. */
  readonly where: string | null;
  readonly detail: string;
}

export interface ReportSectionView {
  readonly id: string;
  readonly title: string;
  readonly entries: readonly ReportEntryView[];
}

const KIND_ORDER: readonly ReportItemKind[] = [
  'chart',
  'chartSheet',
  'drawing',
  'image',
  'pivotTable',
  'pivotCache',
  'externalLink',
  'table',
  'slicer',
  'timeline',
  'sparkline',
  'macro',
  'signature',
  'protection',
  'unsupportedFunction',
  'refusedFormula',
  'unsupportedConditionalFormat',
  'customXml',
  'customProperty',
  'comment',
  'other',
];

/** Group the open-time items by kind, in a stable, readable order. */
export function compatibilitySections(report: CompatibilityReport): ReportSectionView[] {
  const sections: ReportSectionView[] = [];
  for (const kind of KIND_ORDER) {
    const items = report.items.filter((item) => item.kind === kind);
    if (items.length === 0) continue;
    sections.push({
      id: kind,
      title: reportKindLabel(kind),
      entries: items.map((item, index) => entryView(item, `${kind}:${index}`)),
    });
  }
  return sections;
}

const CATEGORY_TITLES: Readonly<Record<ExportReportCategory, string>> = {
  unchanged: 'Unchanged',
  transformed: 'Transformed',
  stale: 'Preserved but stale',
  removed: 'Removed',
  warning: 'Warning',
};

/** Group the save-time entries by the five §7 categories, in report order. */
export function exportSections(report: ExportReport): ReportSectionView[] {
  const sections: ReportSectionView[] = [];
  for (const category of EXPORT_REPORT_CATEGORIES) {
    const entries = report.entries.filter((entry) => entry.category === category);
    if (entries.length === 0) continue;
    sections.push({
      id: category,
      title: `${CATEGORY_TITLES[category]} (${entries.length})`,
      entries: entries.map((entry, index) => entryView(entry, `${category}:${index}`)),
    });
  }
  return sections;
}

function entryView(item: ReportItem, id: string): ReportEntryView {
  const where =
    item.sheetName === null
      ? item.anchor
      : item.anchor === null
        ? item.sheetName
        : `${item.sheetName}!${item.anchor}`;
  return {
    id,
    title: item.part === null ? item.title : `${item.title} — ${item.part}`,
    where,
    detail: item.detail,
  };
}
