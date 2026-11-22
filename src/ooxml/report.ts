import type { GridRange, Workbook } from '@model/types';
import { columnToLetters, formatCellRef, parseCellKey } from '@model/types';
import type { UnsupportedConditionalFormat } from './conditional-format-xml';
import { findRefusedFormulaCells } from './patch/cell-xml';
import type { CompatibilityInfo } from './open';
import type { PackagePatchReport, WorkbookPackage } from './types';

/**
 * Compatibility and export reports (file-compat §7).
 *
 * Both reports are read-only views assembled from facts the open and save
 * paths already hold: the `CompatibilityInfo` from `open.ts`, the banked
 * unsupported-conditional-format listing, the save's `PackagePatchReport`,
 * and the before/after packages. Nothing here mutates a workbook, touches a
 * part or contacts a network; the text form is what the dialogs copy.
 *
 * The export report uses exactly the five §7 categories. "Unchanged",
 * "Transformed" and "Preserved but stale" describe preserved parts;
 * "Removed" is limited to an explicit format change such as `.xlsm` →
 * `.xlsx`; "Warning" is a signature, unenforced protection or flattened
 * content. The report is surfaced only when something is stale, removed or
 * warned about (`needsAttention`).
 */

// ---------------------------------------------------------------------------
// Kinds and labels
// ---------------------------------------------------------------------------

/** A Tier B construct, classified by its part path. */
export type PreservedPartKind =
  | 'chart'
  | 'chartSheet'
  | 'drawing'
  | 'image'
  | 'pivotTable'
  | 'pivotCache'
  | 'externalLink'
  | 'macro'
  | 'signature'
  | 'table'
  | 'slicer'
  | 'timeline'
  | 'sparkline'
  | 'customXml'
  | 'customProperty'
  | 'comment'
  | 'other';

/** Everything a report item can describe: a part, a setting, or a cell. */
export type ReportItemKind =
  | PreservedPartKind
  | 'protection'
  | 'unsupportedFunction'
  | 'unsupportedConditionalFormat'
  | 'refusedFormula';

const PRESERVED_LABELS: Readonly<Record<PreservedPartKind, string>> = {
  chart: 'Chart',
  chartSheet: 'Chart sheet',
  drawing: 'Drawing',
  image: 'Image',
  pivotTable: 'PivotTable',
  pivotCache: 'Pivot cache',
  externalLink: 'External link',
  macro: 'Macros (VBA project)',
  signature: 'Digital signature',
  table: 'Table',
  slicer: 'Slicer',
  timeline: 'Timeline',
  sparkline: 'Sparkline',
  customXml: 'Custom XML',
  customProperty: 'Custom property',
  comment: 'Comment part',
  other: 'Preserved part',
};

/**
 * What calcplus will and will not do with a construct (file-compat §4, §6).
 * The reports quote this text so the promise and the limitation sit together.
 */
const PRESERVED_ACTIONS: Readonly<Record<PreservedPartKind, string>> = {
  chart:
    'Preserved byte-for-byte but not rendered. Plain A1 series ranges are rewritten on structural edits; a chart fed by a defined name blocks the edit.',
  chartSheet: 'Preserved byte-for-byte but not rendered.',
  drawing:
    'Preserved byte-for-byte but not rendered. Anchors are transformed on row or column insert and delete where they are modelled.',
  image: 'Preserved byte-for-byte but not rendered.',
  pivotTable:
    'Definition and cache preserved and never refreshed. Pivot output shows the stored values; an edit to the source range is blocked.',
  pivotCache: 'Preserved and never refreshed.',
  externalLink: 'Preserved and never resolved or refreshed. Cached values are shown.',
  macro: 'Preserved byte-for-byte and never executed, inspected or edited.',
  signature: 'Preserved, but any save invalidates the signature.',
  table:
    'Preserved. Structured references are not evaluated, so a table reference can go stale and is reported.',
  slicer: 'Preserved but not displayed.',
  timeline: 'Preserved but not displayed.',
  sparkline: 'Preserved but not displayed.',
  customXml: 'Preserved byte-for-byte.',
  customProperty: 'Preserved but not editable.',
  comment: 'Preserved byte-for-byte.',
  other: 'Preserved byte-for-byte.',
};

const PROTECTION_ACTION =
  'Preserved but not enforced: calcplus can make edits Excel would have blocked.';

const CF_REASON_TEXT: Readonly<Record<string, string>> = {
  formula: 'a formula-based rule',
  aboveAverage: 'an above/below-average rule',
  blankOrError: 'a blank-or-error rule',
  unknownRuleType: 'a rule type calcplus does not model',
  unknownOperator: 'an unknown operator',
  formulaThreshold: 'a formula threshold',
};

/**
 * Classify a preserved part by path. The listing exists so the report can
 * name a construct without reading the part; lane 4's part policy supplies
 * the same classification for save-time facts.
 */
export function classifyPreservedPart(path: string): PreservedPartKind {
  const lower = path.toLowerCase();
  if (lower.startsWith('xl/chartsheets/')) return 'chartSheet';
  if (lower.startsWith('xl/charts/')) return 'chart';
  if (lower.startsWith('xl/media/')) return 'image';
  if (lower.startsWith('xl/drawings/')) return 'drawing';
  if (lower.startsWith('xl/pivottables/')) return 'pivotTable';
  if (lower.startsWith('xl/pivotcache/')) return 'pivotCache';
  if (lower.startsWith('xl/externallinks/')) return 'externalLink';
  if (lower.endsWith('vbaproject.bin') || lower.endsWith('vbadata.xml')) return 'macro';
  if (lower.startsWith('_xmlsignatures/') || lower.endsWith('.sigs')) return 'signature';
  if (lower.startsWith('xl/tables/')) return 'table';
  if (lower.startsWith('xl/slicers/') || lower.startsWith('xl/slicercaches/')) return 'slicer';
  if (lower.startsWith('xl/timelines/') || lower.startsWith('xl/timelinecaches/'))
    return 'timeline';
  if (lower.startsWith('customxml/')) return 'customXml';
  if (lower === 'docprops/custom.xml') return 'customProperty';
  if (lower.startsWith('xl/comments') || lower.startsWith('xl/threadedcomments')) return 'comment';
  return 'other';
}

/**
 * Package plumbing is not reportable content. `calcChain.xml` is deliberately
 * save-managed (§4) and document properties are not user constructs, so the
 * reports stay about workbook content.
 */
export function isReportablePart(path: string): boolean {
  const lower = path.toLowerCase();
  if (lower.endsWith('.rels')) return false;
  if (lower === '[content_types].xml') return false;
  if (lower === 'xl/calcchain.xml') return false;
  if (lower === 'docprops/core.xml' || lower === 'docprops/app.xml') return false;
  return true;
}

// ---------------------------------------------------------------------------
// Shared item shape
// ---------------------------------------------------------------------------

/** One listed construct, setting or cell. */
export interface ReportItem {
  readonly kind: ReportItemKind;
  readonly title: string;
  /** The package part, when the item describes one. */
  readonly part: string | null;
  readonly sheetName: string | null;
  /** An A1 cell or range, when the item names a position. */
  readonly anchor: string | null;
  /** What calcplus will and will not do. */
  readonly detail: string;
}

/**
 * A preserved-part status from lane 4's `part-policy.ts` listing
 * (`listPreservedParts`). The fields are optional so the report also accepts a
 * minimal synthetic listing; when lane 4's entry is passed, its `category`,
 * `description` and `note` win over the path-based classification.
 */
export type PartReportStatus = 'unchanged' | 'transformed' | 'stale' | 'removed' | 'warning';

export interface PreservedPartStatus {
  readonly path: string;
  readonly status: PartReportStatus;
  /** Lane 4's `PartCategory`, when the listing supplies it. */
  readonly category?: string;
  /** An explicit kind, taking precedence over the category and the path. */
  readonly kind?: PreservedPartKind;
  /** A human title such as "Chart", when the listing supplies one. */
  readonly description?: string;
  /** What calcplus will and will not do, when the listing supplies it. */
  readonly note?: string;
  /** An explicit detail, taking precedence over `note`. */
  readonly detail?: string;
  readonly sheetName?: string | null;
  readonly anchor?: string | null;
}

/** Lane 4's categories that differ from this module's kind names. */
const CATEGORY_KINDS: Readonly<Record<string, PreservedPartKind>> = {
  chart: 'chart',
  chartSheet: 'chartSheet',
  drawing: 'drawing',
  vmlDrawing: 'drawing',
  media: 'image',
  oleObject: 'drawing',
  pivotTable: 'pivotTable',
  pivotCache: 'pivotCache',
  externalLink: 'externalLink',
  vbaProject: 'macro',
  signature: 'signature',
  table: 'table',
  slicer: 'slicer',
  customXml: 'customXml',
  customProperties: 'customProperty',
};

function kindForPreserved(status: PreservedPartStatus): PreservedPartKind {
  if (status.kind !== undefined) return status.kind;
  if (status.category !== undefined && CATEGORY_KINDS[status.category] !== undefined) {
    return CATEGORY_KINDS[status.category]!;
  }
  return classifyPreservedPart(status.path);
}

function partItem(
  kind: PreservedPartKind,
  path: string,
  extra: Partial<Pick<ReportItem, 'title' | 'sheetName' | 'anchor' | 'detail'>> = {},
): ReportItem {
  return {
    kind,
    title: extra.title ?? `${PRESERVED_LABELS[kind]} ${partFileName(path)}`,
    part: path,
    sheetName: extra.sheetName ?? null,
    anchor: extra.anchor ?? null,
    detail: extra.detail ?? PRESERVED_ACTIONS[kind],
  };
}

/** Map a lane-4 listing entry onto a report item, preferring its authored text. */
function itemFromPreserved(status: PreservedPartStatus): ReportItem {
  const kind = kindForPreserved(status);
  return partItem(kind, status.path, {
    ...(status.description === undefined ? {} : { title: status.description }),
    ...(status.detail === undefined ? {} : { detail: status.detail }),
    ...(status.detail === undefined && status.note !== undefined ? { detail: status.note } : {}),
    ...(status.sheetName === undefined ? {} : { sheetName: status.sheetName }),
    ...(status.anchor === undefined ? {} : { anchor: status.anchor }),
  });
}

function partFileName(path: string): string {
  const slash = path.lastIndexOf('/');
  return slash < 0 ? path : path.slice(slash + 1);
}

// ---------------------------------------------------------------------------
// Compatibility report (on open)
// ---------------------------------------------------------------------------

export interface CompatibilityReportInput {
  readonly workbook: Workbook;
  /** The open path's compatibility facts. */
  readonly compatibility: CompatibilityInfo;
  /** The banked unsupported-rule listing (`listUnsupportedConditionalFormats`). */
  readonly unsupportedConditionalFormats?: readonly UnsupportedConditionalFormat[];
  /** Lane 4's preserved-part classification, when it is available. */
  readonly preserved?: readonly PreservedPartStatus[] | null;
}

export interface CompatibilityReport {
  readonly items: readonly ReportItem[];
  /** Items per kind, in first-seen order. */
  readonly counts: Readonly<Partial<Record<ReportItemKind, number>>>;
  /**
   * `true` only when the report has something to report. Notes alone (such as
   * the locale number-format note) do not count, matching §7's banner rule.
   */
  readonly hasContent: boolean;
  /** Copy-as-text form, including notes. */
  readonly text: string;
  /** Open-path notes that are not constructs, kept for the panel's text. */
  readonly notes: readonly string[];
}

export function buildCompatibilityReport(input: CompatibilityReportInput): CompatibilityReport {
  const items: ReportItem[] = [];
  const preserved = input.preserved ?? null;

  if (preserved !== null) {
    for (const status of [...preserved].sort((left, right) =>
      left.path.localeCompare(right.path),
    )) {
      if (!isReportablePart(status.path)) continue;
      items.push(itemFromPreserved(status));
    }
  } else {
    for (const path of [...input.compatibility.opaquePartPaths].sort()) {
      if (!isReportablePart(path)) continue;
      items.push(partItem(classifyPreservedPart(path), path));
    }
  }
  // An external-link relationship may point at a part the listing already
  // holds; only add a link the listing missed (a non-part external relationship).
  for (const path of [...input.compatibility.externalLinkParts].sort()) {
    if (items.some((item) => item.part === path)) continue;
    items.push(partItem('externalLink', path));
  }
  if (preserved === null) {
    if (input.compatibility.hasMacros && !items.some((item) => item.kind === 'macro')) {
      items.push({
        kind: 'macro',
        title: PRESERVED_LABELS.macro,
        part: findPart(input.workbook, 'vbaproject.bin'),
        sheetName: null,
        anchor: null,
        detail: PRESERVED_ACTIONS.macro,
      });
    }
    if (
      input.compatibility.hasDigitalSignatures &&
      !items.some((item) => item.kind === 'signature')
    ) {
      items.push({
        kind: 'signature',
        title: PRESERVED_LABELS.signature,
        part: null,
        sheetName: null,
        anchor: null,
        detail: PRESERVED_ACTIONS.signature,
      });
    }
  }

  for (const sheet of input.workbook.sheets) {
    if (sheet.protection === null) continue;
    items.push({
      kind: 'protection',
      title: 'Sheet protection',
      part: null,
      sheetName: sheet.name,
      anchor: null,
      detail: PROTECTION_ACTION,
    });
  }
  if (input.workbook.protection !== null) {
    items.push({
      kind: 'protection',
      title: 'Workbook protection',
      part: null,
      sheetName: null,
      anchor: null,
      detail: PROTECTION_ACTION,
    });
  }

  for (const rule of input.unsupportedConditionalFormats ?? []) {
    items.push({
      kind: 'unsupportedConditionalFormat',
      title: 'Conditional formatting rule',
      part: null,
      sheetName: rule.sheetName,
      anchor: rule.ranges.map(rangeText).join(' '),
      detail: `Preserved byte-for-byte but not painted: ${CF_REASON_TEXT[rule.reason] ?? rule.reason}${rule.ruleType === '' ? '' : ` (${rule.ruleType})`}.`,
    });
  }

  for (const stale of staleCells(input.workbook)) {
    items.push(stale);
  }

  for (const refused of refusedFormulaItems(input.workbook)) {
    items.push(refused);
  }

  const counts: Partial<Record<ReportItemKind, number>> = {};
  for (const item of items) counts[item.kind] = (counts[item.kind] ?? 0) + 1;

  const report: CompatibilityReport = {
    items,
    counts,
    hasContent: items.length > 0,
    text: '',
    notes: input.compatibility.notes,
  };
  return { ...report, text: formatCompatibilityReportText({ ...report, text: '' }) };
}

/**
 * Cells the engine could not recalculate. The report names each one so the
 * user can tell a stale cached result from a live one (§6.1).
 */
function staleCells(workbook: Workbook): ReportItem[] {
  const items: ReportItem[] = [];
  for (const sheet of workbook.sheets) {
    const cells: { key: string; row: number; col: number }[] = [];
    for (const [key, cell] of sheet.cells) {
      if (cell.stale !== true) continue;
      const address = parseCellKey(key);
      cells.push({ key, row: address.row, col: address.col });
    }
    cells.sort((left, right) => left.row - right.row || left.col - right.col);
    for (const address of cells) {
      const cell = sheet.cells.get(address.key);
      const ref = formatCellRef(address.row, address.col);
      const cached =
        cell === undefined || cell.savedResult === null
          ? 'No cached result was saved, so #NAME? is shown.'
          : 'The saved cached result is shown and marked stale.';
      items.push({
        kind: 'unsupportedFunction',
        title:
          cell?.formula === null || cell?.formula === undefined ? `Cell ${ref}` : `Formula ${ref}`,
        part: null,
        sheetName: sheet.name,
        anchor: ref,
        detail: `${cached} calcplus did not evaluate the construct, so Excel recalculates on load.`,
      });
    }
  }
  return items;
}

function findPart(workbook: Workbook, suffix: string): string | null {
  for (const path of workbook.opaqueParts.keys()) {
    if (path.toLowerCase().endsWith(suffix)) return path;
  }
  return null;
}

/**
 * Cells whose formula Excel refuses (a wrong argument count). Excel strips such
 * a formula when it opens the file, so calcplus keeps it visible in the model
 * but never writes it: the open report names it, and the save writes the
 * cached value alone (file-compat §1, §7).
 */
function refusedFormulaItems(workbook: Workbook): ReportItem[] {
  return findRefusedFormulaCells(workbook).map((found) => ({
    kind: 'refusedFormula',
    title: `Formula ${found.ref}`,
    part: null,
    sheetName: found.sheetName,
    anchor: found.ref,
    detail: `Excel will not accept \`${found.formula}\` — it has the wrong number of arguments — so calcplus does not write it to the file. The formula stays visible here so it can be corrected.`,
  }));
}

export function compatibilityReportSummary(report: CompatibilityReport): string {
  if (!report.hasContent) return 'Nothing to report.';
  const parts: string[] = [];
  for (const [kind, count] of Object.entries(report.counts)) {
    if (count === undefined || count === 0) continue;
    parts.push(`${count} ${labelForKind(kind as ReportItemKind, count)}`);
  }
  return parts.join(', ');
}

// ---------------------------------------------------------------------------
// Export report (on save)
// ---------------------------------------------------------------------------

/** The five categories of file-compat §7, in report order. */
export const EXPORT_REPORT_CATEGORIES = [
  'unchanged',
  'transformed',
  'stale',
  'removed',
  'warning',
] as const;

export type ExportReportCategory = (typeof EXPORT_REPORT_CATEGORIES)[number];

export interface ExportReportCounts {
  readonly unchanged: number;
  readonly transformed: number;
  readonly stale: number;
  readonly removed: number;
  readonly warning: number;
}

export interface ExportReportEntry extends ReportItem {
  readonly category: ExportReportCategory;
}

export interface ExportReport {
  readonly entries: readonly ExportReportEntry[];
  readonly counts: ExportReportCounts;
  /**
   * §10 criterion 10: the report is shown whenever anything is stale, removed
   * or warned about. An unchanged or transformed part alone does not open it.
   */
  readonly needsAttention: boolean;
  /** Copy-as-text form. */
  readonly text: string;
}

export interface ExportReportInput {
  readonly workbook: Workbook;
  /** The package as it stood before this save. */
  readonly before: WorkbookPackage;
  /** The package this save produced. */
  readonly after: WorkbookPackage;
  readonly report: PackagePatchReport;
  readonly compatibility: CompatibilityInfo;
  /** A no-op save writes nothing, so it has nothing to report. */
  readonly noOp: boolean;
  /**
   * Whether a structural edit (insert/delete) happened since the file was
   * opened. Without lane 4's status listing it is the only signal that an
   * external link or pivot was left unrefreshed (§6.2, §6.4).
   */
  readonly structuralChange?: boolean;
  /** A confirmed `.xlsm` → `.xlsx` Save As, the only route to a removal (§7). */
  readonly formatChange?: { readonly from: string; readonly to: string } | null;
  /** Lane 4's preserved-part listing, when it is available. */
  readonly preserved?: readonly PreservedPartStatus[] | null;
}

const EMPTY_COUNTS: ExportReportCounts = {
  unchanged: 0,
  transformed: 0,
  stale: 0,
  removed: 0,
  warning: 0,
};

export function buildExportReport(input: ExportReportInput): ExportReport {
  if (input.noOp) {
    const empty: ExportReport = {
      entries: [],
      counts: EMPTY_COUNTS,
      needsAttention: false,
      text: '',
    };
    return { ...empty, text: formatExportReportText(empty) };
  }

  const entries: ExportReportEntry[] = [];
  const preservedByPath = new Map(
    (input.preserved ?? []).map((status) => [status.path, status] as const),
  );
  const structural = input.structuralChange === true;
  const paths = [...input.workbook.opaqueParts.keys()]
    .filter((path) => isReportablePart(path))
    .sort();

  for (const path of paths) {
    const before = input.before.entries.get(path);
    if (before === undefined) continue;
    const after = input.after.entries.get(path);
    if (after === undefined) continue; // A removal is handled below, and only for a format change.
    const status = preservedByPath.get(path);
    const kind = status === undefined ? classifyPreservedPart(path) : kindForPreserved(status);
    const item = status === undefined ? partItem(kind, path) : itemFromPreserved(status);
    if (status?.status === 'removed') {
      entries.push({ ...item, category: 'removed' });
      continue;
    }
    if (status?.status === 'warning') {
      entries.push({ ...item, category: 'warning' });
      continue;
    }
    const unchanged = bytesEqual(before, after);
    const category: ExportReportCategory =
      status?.status === 'stale'
        ? 'stale'
        : status?.status === 'transformed'
          ? 'transformed'
          : status?.status === 'unchanged'
            ? 'unchanged'
            : !unchanged
              ? 'transformed'
              : structural &&
                  (kind === 'externalLink' || kind === 'pivotTable' || kind === 'pivotCache')
                ? 'stale'
                : 'unchanged';
    entries.push({
      ...item,
      ...(category === 'stale' && status?.detail === undefined && status?.note === undefined
        ? {
            detail:
              kind === 'externalLink'
                ? `${PRESERVED_ACTIONS.externalLink} The structural edit was not propagated to the external reference.`
                : kind === 'pivotCache'
                  ? `${PRESERVED_ACTIONS.pivotCache} The structural edit was not propagated to the cache.`
                  : `${PRESERVED_ACTIONS.pivotTable} It is preserved but not refreshed after the structural edit.`,
          }
        : {}),
      category,
    });
  }

  for (const status of input.preserved ?? []) {
    if (paths.includes(status.path)) continue;
    if (!isReportablePart(status.path)) continue;
    const item = itemFromPreserved(status);
    if (status.status === 'removed') {
      if (input.after.entries.has(status.path)) continue;
      entries.push({ ...item, category: 'removed' });
      continue;
    }
    if (status.status === 'warning') {
      entries.push({ ...item, category: 'warning' });
      continue;
    }
    entries.push({
      ...item,
      category:
        status.status === 'stale'
          ? 'stale'
          : status.status === 'transformed'
            ? 'transformed'
            : 'unchanged',
    });
  }

  if (input.formatChange !== null && input.formatChange !== undefined) {
    const removedCandidates = new Set<string>([
      ...input.report.removedParts,
      ...paths.filter((path) => !input.after.entries.has(path)),
    ]);
    for (const path of [...removedCandidates].sort()) {
      if (input.after.entries.has(path)) continue;
      if (entries.some((entry) => entry.part === path && entry.category === 'removed')) continue;
      const kind = classifyPreservedPart(path);
      if (kind !== 'macro' && kind !== 'signature') continue;
      entries.push({
        ...partItem(kind, path, {
          detail: `Removed because the file was saved as .${input.formatChange.to}.`,
        }),
        category: 'removed',
      });
    }
  }

  const signatureAlreadyReported =
    input.compatibility.hasDigitalSignatures &&
    entries.some(
      (entry) =>
        (entry.kind === 'signature' || entry.part?.toLowerCase().endsWith('.sigs') === true) &&
        (entry.category === 'warning' || entry.category === 'removed'),
    );
  if (input.compatibility.hasDigitalSignatures && !signatureAlreadyReported) {
    entries.push({
      kind: 'signature',
      title: 'Digital signature invalidated',
      part: null,
      sheetName: null,
      anchor: null,
      detail: 'Saving rewrites the package, so the saved file carries an invalid signature.',
      category: 'warning',
    });
  }
  const protectedSheets = input.workbook.sheets.filter((sheet) => sheet.protection !== null);
  if (protectedSheets.length > 0 || input.workbook.protection !== null) {
    entries.push({
      kind: 'protection',
      title: 'Protection not enforced',
      part: null,
      sheetName: protectedSheets.map((sheet) => sheet.name).join(', ') || null,
      anchor: null,
      detail:
        'Protection flags are preserved, but calcplus does not enforce them; an edit Excel would have blocked may reach this save.',
      category: 'warning',
    });
  }

  // File-compat §1: dropping a formula is a change to the user's data, so it is
  // reported rather than silent. The cell is written with its cached value; the
  // formula text stays visible in the app, where it can be corrected.
  for (const found of findRefusedFormulaCells(input.workbook)) {
    entries.push({
      kind: 'refusedFormula',
      title: `Formula ${found.ref} not written`,
      part: null,
      sheetName: found.sheetName,
      anchor: found.ref,
      detail: `\`${found.formula}\` has the wrong number of arguments, so Excel would remove it when it opened the file. It was not written; the saved cell holds its value only.`,
      category: 'warning',
    });
  }

  const ordered = EXPORT_REPORT_CATEGORIES.flatMap((category) =>
    entries.filter((entry) => entry.category === category),
  );
  const counts = countExportEntries(ordered);
  const report: ExportReport = {
    entries: ordered,
    counts,
    needsAttention: counts.stale + counts.removed + counts.warning > 0,
    text: '',
  };
  return { ...report, text: formatExportReportText(report) };
}

function countExportEntries(entries: readonly ExportReportEntry[]): ExportReportCounts {
  const counts = { ...EMPTY_COUNTS } as Record<ExportReportCategory, number>;
  for (const entry of entries) counts[entry.category] += 1;
  return counts;
}

export function exportReportSummary(report: ExportReport): string {
  if (report.entries.length === 0) return 'Nothing to report.';
  const parts: string[] = [];
  for (const category of EXPORT_REPORT_CATEGORIES) {
    const count = report.counts[category];
    if (count === 0) continue;
    parts.push(
      `${count} ${count === 1 ? categoryLabel(category) : `${categoryLabel(category)} items`}`,
    );
  }
  return parts.join(', ');
}

function categoryLabel(category: ExportReportCategory): string {
  switch (category) {
    case 'unchanged':
      return 'unchanged';
    case 'transformed':
      return 'transformed';
    case 'stale':
      return 'preserved but stale';
    case 'removed':
      return 'removed';
    case 'warning':
      return 'warning';
  }
}

// ---------------------------------------------------------------------------
// Text forms (the dialogs copy these)
// ---------------------------------------------------------------------------

export function formatCompatibilityReportText(report: CompatibilityReport): string {
  const lines = ['Compatibility report'];
  if (!report.hasContent) {
    lines.push('', 'Nothing to report.');
  } else {
    lines.push(`  ${compatibilityReportSummary(report)}`, '');
    for (const item of report.items) {
      lines.push(`[${kindLabel(item.kind)}] ${itemLabel(item)}`);
      lines.push(`  ${item.detail}`);
    }
  }
  if (report.notes.length > 0) {
    lines.push('', 'Notes');
    for (const note of report.notes) lines.push(`  - ${note}`);
  }
  return lines.join('\n');
}

export function formatExportReportText(report: ExportReport): string {
  const lines = ['Export report'];
  if (report.entries.length === 0) {
    lines.push('', 'Nothing to report.');
    return lines.join('\n');
  }
  lines.push(`  ${exportReportSummary(report)}`);
  for (const category of EXPORT_REPORT_CATEGORIES) {
    const entries = report.entries.filter((entry) => entry.category === category);
    if (entries.length === 0) continue;
    lines.push('', `${categoryLabel(category)} (${entries.length})`);
    for (const entry of entries) {
      lines.push(`[${kindLabel(entry.kind)}] ${itemLabel(entry)}`);
      lines.push(`  ${entry.detail}`);
    }
  }
  return lines.join('\n');
}

function itemLabel(item: ReportItem): string {
  const where: string[] = [];
  if (item.sheetName !== null) where.push(item.sheetName);
  if (item.anchor !== null) where.push(item.anchor);
  const position = where.length === 0 ? '' : ` — ${where.join('!')}`;
  const part = item.part === null ? '' : ` (${item.part})`;
  return `${item.title}${part}${position}`;
}

function kindLabel(kind: ReportItemKind): string {
  switch (kind) {
    case 'protection':
      return 'Protection';
    case 'unsupportedFunction':
      return 'Not recalculated';
    case 'unsupportedConditionalFormat':
      return 'Unsupported rule';
    case 'refusedFormula':
      return 'Refused formula';
    default:
      return PRESERVED_LABELS[kind];
  }
}

/** The display label for one item kind (the dialogs' section headings). */
export function reportKindLabel(kind: ReportItemKind): string {
  return kindLabel(kind);
}

function labelForKind(kind: ReportItemKind, count: number): string {
  const label = kindLabel(kind);
  if (count === 1) return label.toLowerCase();
  switch (kind) {
    case 'externalLink':
      return 'external links';
    case 'unsupportedFunction':
      return 'not recalculated';
    case 'unsupportedConditionalFormat':
      return 'unsupported rules';
    case 'protection':
      return 'protection settings';
    case 'customXml':
      return 'custom XML parts';
    default:
      return `${label.toLowerCase()}s`;
  }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function rangeText(range: GridRange): string {
  const start = `${columnToLetters(range.startCol)}${range.startRow + 1}`;
  const end = `${columnToLetters(range.endCol)}${range.endRow + 1}`;
  return start === end ? start : `${start}:${end}`;
}

function bytesEqual(left: Uint8Array, right: Uint8Array): boolean {
  if (left.byteLength !== right.byteLength) return false;
  for (let index = 0; index < left.byteLength; index += 1) {
    if (left[index] !== right[index]) return false;
  }
  return true;
}
