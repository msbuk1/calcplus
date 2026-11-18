import type { CellValue } from '@model/value';
import type { Sheet } from '@model/types';
import type {
  CfCalendarDate,
  CfCellSource,
  CfCellSubject,
  CfDxfResolver,
  ConditionalFormatRule,
} from '@model/conditional-format';
import { createConditionalFormatIndex } from '@model/conditional-format';
import {
  conditionalFormatRulesFor,
  createDxfResolver,
  parseDxfTable,
  type DxfTable,
} from '@ooxml/conditional-format-xml';
import type { ConditionalFormattingInputs, ProjectedConditionalFormatting } from './types';

/**
 * Conditional formatting inside the sheet projection (conditional formatting
 * §5, §6 and ADR 0007).
 *
 * The projection is the only place conditional formatting meets the model: it
 * builds the engine's per-sheet index, hands it a cell source backed by the
 * sheet's values and display text, and exposes paint-ready overlays to the
 * canvas. The grid never imports the engine and never evaluates a rule.
 *
 * Rule source: the model keeps the authored `<conditionalFormatting>` elements
 * verbatim in `sheet.conditionalFormats`, so the rules are rebuilt from those
 * records. A caller may pass explicit rules instead — that is the seam a test
 * fixture or a future rule-authoring surface uses, and the one surviving
 * integration point is the `dxfs` table (see `ConditionalFormattingInputs`).
 */

export interface ConditionalFormattingSources {
  readonly sheet: Sheet;
  /** The document version; a rule re-parse only happens when it changes. */
  readonly version: number;
  readonly date1904: boolean;
  readonly indexedPalette: readonly string[] | null;
  /** The sheet's worksheet-part bytes, for the x14 extension payloads. */
  readonly sourceBytes: Uint8Array | null;
  /** The styles-part bytes, for the `dxfs` table a `dxfId` rule formats with. */
  readonly stylesBytes: Uint8Array | null;
  /** Typed values and display text, for rule evaluation. */
  readonly cells: CfCellSource;
  /** Number formatting under an arbitrary code, for a CF number format. */
  readonly format: (value: CellValue, code: string) => string;
  readonly inputs?: ConditionalFormattingInputs | null;
}

/**
 * Parsed rules per sheet and document version. The rule records are small, but
 * parsing them for every paint would be waste: a zoom or other view-only change
 * rebuilds the projection without changing the version, so it reuses this
 * result. A document edit bumps the version and re-parses.
 */
const rulesCache = new WeakMap<
  Sheet,
  { readonly version: number; readonly rules: readonly ConditionalFormatRule[] }
>();

export function createSheetConditionalFormatting(
  sources: ConditionalFormattingSources,
): ProjectedConditionalFormatting | null {
  const inputs = sources.inputs;
  if (inputs === null) return null;

  const rules = inputs?.rules ?? rulesFromRecords(sources);
  if (rules.length === 0) return null;

  const index = createConditionalFormatIndex({
    sheetId: sources.sheet.id,
    rules,
    date1904: sources.date1904,
    today: inputs?.today ?? localToday(),
    cells: sources.cells,
    dxfFor: inputs?.dxfResolver ?? dxfResolverFor(sources.stylesBytes),
  });

  return {
    ruleCount: rules.length,
    evaluate(row, col, subject?: CfCellSubject) {
      return index.evaluateCell(row, col, subject);
    },
    textFor(value, code) {
      return sources.format(value, code);
    },
    invalidateCell(row, col) {
      index.invalidateCell(row, col);
    },
  };
}

function rulesFromRecords(sources: ConditionalFormattingSources): readonly ConditionalFormatRule[] {
  const { sheet, version } = sources;
  const cached = rulesCache.get(sheet);
  if (cached !== undefined && cached.version === version) return cached.rules;
  const rules = conditionalFormatRulesFor(sheet, sources.sourceBytes, {
    indexedPalette: sources.indexedPalette,
  });
  rulesCache.set(sheet, { version, rules });
  return rules;
}

/**
 * The `dxfs` table a `dxfId` rule formats with. The raw styles bytes are
 * stable per document, so the parse is cached on the bytes themselves; a
 * projection rebuild (a zoom) never re-parses the styles part.
 */
const dxfTableCache = new WeakMap<Uint8Array, DxfTable>();

function dxfResolverFor(stylesBytes: Uint8Array | null): CfDxfResolver | undefined {
  if (stylesBytes === null) return undefined;
  let table = dxfTableCache.get(stylesBytes);
  if (table === undefined) {
    table = parseDxfTable(stylesBytes, 'xl/styles.xml');
    dxfTableCache.set(stylesBytes, table);
  }
  return createDxfResolver(table);
}

/**
 * The local calendar date, read once per projection. The engine never reads a
 * clock; the relative date presets get "today" from here (§3.3).
 */
function localToday(): CfCalendarDate {
  const now = new Date();
  return { year: now.getFullYear(), month: now.getMonth() + 1, day: now.getDate() };
}
