import type { GridRange, HyperlinkRecord, OpaquePart, SheetId, Workbook } from '@model/types';
import { formatCellRef } from '@model/types';
import {
  planOpaquePartTransforms,
  transformCellAddress,
  transformHyperlinkLocation,
  transformMergeRange,
  transformRange,
  transformRangeBearingXml,
  type CrossSheetHyperlinkRenumber,
  type OpaquePartKind,
  type OpaquePartRewrite,
  type RangeTransformPlan,
  type StructuralEdit,
  type TransformBlock,
} from '@model/range-transform';
import type { ParsedWorksheet } from './worksheet-xml';
import { parseWorksheetXml } from './worksheet-xml';
import { WorkbookSaveError } from './errors';

/**
 * Structural renumbering of a worksheet part's source bytes.
 *
 * After a structural edit the model's cells sit at new addresses but the last
 * saved part still carries the old ones. A later save patches by address, so
 * the part must be renumbered first: `<row r>`, every cell `r`, `<col min/max>`
 * spans, `<dimension ref>`, merged ranges, hyperlink `ref` and internal
 * `location`, CF/validation `sqref` and rule formulas and the AutoFilter
 * rectangle.
 *
 * The transform is textual: only modelled attributes are rewritten, so every
 * unmodelled attribute, child element and byte the edit does not touch stays
 * exactly as authored (file-compat identity rules). Formulas inside `<f>` are
 * deliberately left alone: the command layer marks every rewritten formula
 * cell dirty, and the normal patch writer writes the model's text.
 *
 * Tier B parts do not live in the worksheet part, so `planOpaquePartPatches`
 * (and the combined `planStructuralPartPatches`) renumber the drawing, chart,
 * comment and pivot-cache parts with the same edit. `renumberWorksheetPart`
 * stays the worksheet-only primitive.
 */

export interface StructurePartInput {
  readonly bytes: Uint8Array;
  readonly partName: string;
  readonly sheetName: string;
  readonly edit: StructuralEdit;
  readonly workbook: Workbook;
}

export interface StructurePartResult {
  readonly bytes: Uint8Array;
  readonly changed: boolean;
}

export function renumberWorksheetPart(input: StructurePartInput): StructurePartResult {
  let parsed: ParsedWorksheet;
  try {
    parsed = parseWorksheetXml(input.bytes, {
      sharedStrings: null,
      partPath: input.partName,
      name: input.sheetName,
    });
  } catch (error) {
    throw new WorkbookSaveError(
      'unsafe-part',
      `The worksheet ${input.partName} could not be read for a structural edit: ${
        error instanceof Error ? error.message : String(error)
      }`,
    );
  }
  // Implied row and cell references are positionally recoverable (the parse
  // computes them), so only genuinely ambiguous structure blocks the edit.
  const blockers = parsed.unsafeReasons.filter(
    (reason) =>
      !reason.startsWith('row without a numeric r') && !reason.startsWith('cell without a valid r'),
  );
  if (blockers.length > 0) {
    throw new WorkbookSaveError(
      'unsafe-part',
      `The worksheet ${input.partName} has a structure calcplus cannot renumber safely (${blockers.join(
        ', ',
      )}). The edit was blocked.`,
    );
  }

  const xml = parsed.xml;
  const edits: TextEdit[] = [];
  appendSheetDataEdits(xml, parsed, input.edit, edits);
  appendColumnEdits(xml, parsed, input.edit, edits);
  appendDimensionEdit(xml, parsed, input.edit, edits);
  appendMergeEdits(xml, parsed, input.edit, edits);
  appendHyperlinkEdits(xml, parsed, input, edits);
  appendConditionalFormatEdits(xml, parsed, input, edits);
  appendValidationEdits(xml, parsed, input, edits);
  appendAutoFilterEdits(xml, parsed, input.edit, edits);

  const next = applyEdits(xml, edits);
  // The transform only rewrites attributes and removes elements whose spans
  // came from a validated parse, so the output cannot be malformed. The save
  // path parses and validates the part again before it writes it.
  return { bytes: new TextEncoder().encode(next), changed: next !== xml };
}

export interface HyperlinkPartInput {
  readonly bytes: Uint8Array;
  readonly partName: string;
  readonly sheetName: string;
  readonly sheetId: SheetId;
  readonly edit: StructuralEdit;
  readonly workbook: Workbook;
}

/**
 * Rewrite only the internal `location` attributes of a worksheet part whose
 * hyperlinks point into the sheet being edited. This is the sibling-sheet half
 * of the structural save: the anchor cells do not move, so `<row>`/`<c>` stay
 * byte-identical and only the target reference changes (ticket 32).
 *
 * A location the transformer cannot parse is left byte-identical. It cannot be
 * attributed to the edited sheet, and an unrelated free-text `location` on
 * another sheet must not block the edit from here.
 */
export function renumberHyperlinkLocationsPart(input: HyperlinkPartInput): StructurePartResult {
  let parsed: ParsedWorksheet;
  try {
    parsed = parseWorksheetXml(input.bytes, {
      sharedStrings: null,
      partPath: input.partName,
      name: input.sheetName,
    });
  } catch (error) {
    throw new WorkbookSaveError(
      'unsafe-part',
      `The worksheet ${input.partName} could not be read for a structural edit: ${
        error instanceof Error ? error.message : String(error)
      }`,
    );
  }
  const container = parsed.rootChildren.find((child) => child.local === 'hyperlinks');
  if (container === undefined) return { bytes: input.bytes, changed: false };
  const xml = parsed.xml;
  const slice = xml.slice(container.start, container.end);
  const patched = patchElements(slice, scanElements(slice, 'hyperlink'), (text) => {
    const match = /(\blocation\s*=\s*)(["'])([^"']*)\2/.exec(text);
    if (match === null) return text;
    const outcome = transformHyperlinkLocation(
      decodeXmlAttributeValue(match[3] ?? ''),
      input.edit,
      { workbook: input.workbook },
      input.sheetId,
    );
    if (!outcome.ok || !outcome.changed) return text;
    return replaceAttribute(text, match, outcome.location);
  });
  if (patched === slice) return { bytes: input.bytes, changed: false };
  const next = `${xml.slice(0, container.start)}${patched}${xml.slice(container.end)}`;
  return { bytes: new TextEncoder().encode(next), changed: true };
}

// ---------------------------------------------------------------------------
// Tier B parts (drawings, charts, comments, pivots)
// ---------------------------------------------------------------------------

export type StructuralPartKind = 'worksheet' | OpaquePartKind;

/**
 * Hyperlink maps that ride a worksheet part patch. A link on another sheet has
 * no cell in the edited sheet to move, but its internal target is still a
 * reference: the same transaction records its before/after map so undo restores
 * the exact link (ticket 32).
 */
export interface StructuralHyperlinkChange {
  readonly sheetId: SheetId;
  readonly before: readonly (readonly [string, HyperlinkRecord])[];
  readonly after: readonly (readonly [string, HyperlinkRecord])[];
}

export interface StructuralPartPatch {
  readonly path: string;
  readonly kind: StructuralPartKind;
  /** The part exactly as it was saved, for the transaction's inverse. */
  readonly before: Uint8Array;
  readonly after: Uint8Array;
  /** Present when the patch also moves a sheet's internal hyperlink targets. */
  readonly hyperlinks?: StructuralHyperlinkChange;
}

export interface StructuralPartPlan {
  readonly patches: readonly StructuralPartPatch[];
  readonly block: TransformBlock | null;
}

/**
 * Plan the renumbering of every opaque part the edit touches, plus the sibling
 * worksheets whose hyperlinks point into the edited sheet. This is the minimal
 * hook for the structural command: pass the plan the command already computed,
 * store the patches in the same transaction, and undo writes `before` back.
 */
export function planOpaquePartPatches(
  workbook: Workbook,
  edit: StructuralEdit,
  plan: RangeTransformPlan | null = null,
): StructuralPartPlan {
  const opaque =
    plan === null
      ? planOpaquePartTransforms(workbook, edit, { workbook })
      : { rewrites: plan.opaqueParts, block: plan.block };
  const patches = opaque.rewrites.map(partPatch);
  if (plan !== null) patches.push(...planCrossSheetHyperlinkPatches(workbook, edit, plan));
  return {
    patches,
    block: opaque.block,
  };
}

/**
 * Rewrite each sibling worksheet part whose hyperlinks point into the edited
 * sheet. The modified bytes and the matching model maps ride one
 * `StructuralPartPatch`, so a save writes the part and undo restores both.
 */
export function planCrossSheetHyperlinkPatches(
  workbook: Workbook,
  edit: StructuralEdit,
  plan: RangeTransformPlan,
): StructuralPartPatch[] {
  const patches: StructuralPartPatch[] = [];
  for (const entry of plan.crossSheetHyperlinks) {
    const sheet = workbook.sheets.find((candidate) => candidate.id === entry.sheetId);
    if (sheet === undefined) continue;
    const before = workbook.sourceParts.get(sheet.partPath);
    if (before === undefined) continue;
    const result = renumberHyperlinkLocationsPart({
      bytes: before,
      partName: sheet.partPath,
      sheetName: sheet.name,
      sheetId: sheet.id,
      edit,
      workbook,
    });
    if (!result.changed) continue;
    patches.push({
      path: sheet.partPath,
      kind: 'worksheet',
      before,
      after: result.bytes,
      hyperlinks: hyperlinkChange(entry),
    });
  }
  return patches;
}

function hyperlinkChange(entry: CrossSheetHyperlinkRenumber): StructuralHyperlinkChange {
  return { sheetId: entry.sheetId, before: entry.before, after: entry.after };
}

/**
 * Plan every part a structural edit renumbers, including the worksheet source
 * part. A caller that uses this must not also call `renumberWorksheetPart`:
 * renumbering twice would apply the edit twice. Throws `WorkbookSaveError`
 * (`unsafe-part`) when the worksheet cannot be cut safely — the structural
 * command already turns that into a block.
 */
export function planStructuralPartPatches(
  workbook: Workbook,
  edit: StructuralEdit,
  plan: RangeTransformPlan | null = null,
): StructuralPartPlan {
  const sheet = workbook.sheets.find((candidate) => candidate.id === edit.sheetId);
  if (sheet === undefined) return { patches: [], block: null };
  const patches: StructuralPartPatch[] = [];
  const sourceBefore = workbook.sourceParts.get(sheet.partPath);
  if (sourceBefore !== undefined) {
    const worksheet = renumberWorksheetPart({
      bytes: sourceBefore,
      partName: sheet.partPath,
      sheetName: sheet.name,
      edit,
      workbook,
    });
    if (worksheet.changed) {
      patches.push({
        path: sheet.partPath,
        kind: 'worksheet',
        before: sourceBefore,
        after: worksheet.bytes,
      });
    }
  }
  const opaque = planOpaquePartPatches(workbook, edit, plan);
  patches.push(...opaque.patches);
  return { patches, block: opaque.block };
}

/**
 * Write one side of a set of part patches into the model's part maps. The
 * inverse is the same call with `'before'`, so a transaction restores the exact
 * prior bytes without recomputing a transform.
 */
export function applyStructuralPartPatches(
  workbook: Workbook,
  patches: readonly StructuralPartPatch[],
  direction: 'before' | 'after',
): void {
  for (const patch of patches) {
    const bytes = direction === 'after' ? patch.after : patch.before;
    if (patch.hyperlinks !== undefined) {
      const sheet = workbook.sheets.find((candidate) => candidate.id === patch.hyperlinks!.sheetId);
      if (sheet !== undefined) {
        replaceHyperlinkMap(
          sheet.hyperlinks,
          direction === 'after' ? patch.hyperlinks.after : patch.hyperlinks.before,
        );
      }
    }
    if (patch.kind === 'worksheet') {
      workbook.sourceParts.set(patch.path, bytes);
      continue;
    }
    const previous: OpaquePart | undefined = workbook.opaqueParts.get(patch.path);
    workbook.opaqueParts.set(patch.path, {
      path: patch.path,
      bytes,
      anchor: previous?.anchor ?? null,
    });
    for (const sheet of workbook.sheets) {
      const existing = sheet.opaqueParts.get(patch.path);
      if (existing !== undefined) {
        sheet.opaqueParts.set(patch.path, { path: patch.path, bytes, anchor: existing.anchor });
      }
    }
  }
}

function partPatch(rewrite: OpaquePartRewrite): StructuralPartPatch {
  return {
    path: rewrite.path,
    kind: rewrite.kind,
    before: rewrite.before,
    after: rewrite.after,
  };
}

function replaceHyperlinkMap(
  map: Map<string, HyperlinkRecord>,
  entries: readonly (readonly [string, HyperlinkRecord])[],
): void {
  map.clear();
  for (const [key, record] of entries) map.set(key, { ...record });
}

// ---------------------------------------------------------------------------
// Rows and cells
// ---------------------------------------------------------------------------

function appendSheetDataEdits(
  xml: string,
  parsed: ParsedWorksheet,
  edit: StructuralEdit,
  edits: TextEdit[],
): void {
  const wholeRowAxis = edit.axis === 'row' && edit.band === undefined;
  for (const row of parsed.rows.values()) {
    const rowAddress = transformCellAddress(row.number, 0, edit);
    if (wholeRowAxis && rowAddress === null) {
      edits.push({ start: row.start, end: row.end, text: '' });
      continue;
    }
    const newRow = wholeRowAxis ? (rowAddress?.row ?? row.number) : row.number;
    if (newRow !== row.number) {
      edits.push({
        start: row.start,
        end: row.openEnd,
        text: setAttribute(xml.slice(row.start, row.openEnd), 'r', String(newRow + 1)),
      });
    }
    for (const cell of row.cells.values()) {
      const next = transformCellAddress(row.number, cell.col, edit);
      if (next === null) {
        edits.push({ start: cell.start, end: cell.end, text: '' });
        continue;
      }
      if (next.row === row.number && next.col === cell.col) continue;
      edits.push({
        start: cell.start,
        end: cell.openEnd,
        text: setAttribute(
          xml.slice(cell.start, cell.openEnd),
          'r',
          formatCellRef(next.row, next.col),
        ),
      });
    }
  }
}

// ---------------------------------------------------------------------------
// Column records
// ---------------------------------------------------------------------------

function appendColumnEdits(
  xml: string,
  parsed: ParsedWorksheet,
  edit: StructuralEdit,
  edits: TextEdit[],
): void {
  // A banded cell edit never moves a column record: the columns themselves
  // stay where they are and only cells inside the band shift.
  if (edit.axis !== 'column' || edit.band !== undefined) return;
  for (const column of parsed.columns) {
    const kept: number[] = [];
    for (let index = column.min; index <= column.max; index += 1) {
      const next = transformCellAddress(0, index, edit);
      if (next !== null) kept.push(next.col);
    }
    if (kept.length === 0) {
      edits.push({ start: column.start, end: column.end, text: '' });
      continue;
    }
    const openTag = xml.slice(column.start, column.openEnd);
    const selfClosing = /\/\s*>$/.test(openTag);
    const replacement = coalesce(kept)
      .map(([min, max]) => setColumnSpan(openTag, selfClosing, min, max))
      .join('');
    if (replacement === xml.slice(column.start, column.end)) continue;
    edits.push({ start: column.start, end: column.end, text: replacement });
  }
}

function setColumnSpan(openTag: string, selfClosing: boolean, min: number, max: number): string {
  const updated = setAttribute(
    setAttribute(openTag, 'min', String(min + 1)),
    'max',
    String(max + 1),
  );
  if (selfClosing) return updated;
  const qualified = /^<([^\s/>]+)/.exec(updated)?.[1] ?? 'col';
  return `${updated.replace(/\/\s*>$/, '>')}</${qualified}>`;
}

function coalesce(values: readonly number[]): [number, number][] {
  const sorted = [...values].sort((left, right) => left - right);
  const ranges: [number, number][] = [];
  for (const value of sorted) {
    const last = ranges.at(-1);
    if (last !== undefined && value === last[1] + 1) last[1] = value;
    else ranges.push([value, value]);
  }
  return ranges;
}

// ---------------------------------------------------------------------------
// Used range
// ---------------------------------------------------------------------------

function appendDimensionEdit(
  xml: string,
  parsed: ParsedWorksheet,
  edit: StructuralEdit,
  edits: TextEdit[],
): void {
  if (edit.band !== undefined) return;
  const element = parsed.rootChildren.find((child) => child.local === 'dimension');
  if (element === undefined) return;
  const slice = xml.slice(element.start, element.openEnd);
  const match = /(\bref\s*=\s*)(["'])([^"']*)\2/.exec(slice);
  if (match === null) return;
  const range = parseRangeText(match[3] ?? '');
  const transformed = range === null ? null : transformRange(range, edit);
  const text = transformed === null ? 'A1' : formatRangeText(transformed);
  if (text === match[3]) return;
  edits.push({
    start: element.start,
    end: element.openEnd,
    text: `${slice.slice(0, match.index)}${match[1]}${match[2]}${text}${match[2]}${slice.slice(
      match.index + match[0].length,
    )}`,
  });
}

// ---------------------------------------------------------------------------
// Merged ranges
// ---------------------------------------------------------------------------

function appendMergeEdits(
  xml: string,
  parsed: ParsedWorksheet,
  edit: StructuralEdit,
  edits: TextEdit[],
): void {
  const container = parsed.rootChildren.find((child) => child.local === 'mergeCells');
  if (container === undefined) return;
  const slice = xml.slice(container.start, container.end);
  const elements = scanElements(slice, 'mergeCell');
  let removed = 0;
  const patched = patchElements(slice, elements, (text) => {
    const match = /(\bref\s*=\s*)(["'])([^"']*)\2/.exec(text);
    const range = match === null ? null : parseRangeText(match[3] ?? '');
    if (match === null || range === null) return text;
    const transformed = transformMergeRange(range, edit);
    if (transformed === null) {
      removed += 1;
      return null;
    }
    if (sameRange(range, transformed)) return text;
    return replaceAttribute(text, match, formatRangeText(transformed));
  });
  const withCount =
    removed > 0 ? updateCount(patched, Math.max(0, elements.length - removed)) : patched;
  if (withCount === slice) return;
  edits.push({ start: container.start, end: container.end, text: withCount });
}

// ---------------------------------------------------------------------------
// Hyperlinks
// ---------------------------------------------------------------------------

function appendHyperlinkEdits(
  xml: string,
  parsed: ParsedWorksheet,
  input: StructurePartInput,
  edits: TextEdit[],
): void {
  const container = parsed.rootChildren.find((child) => child.local === 'hyperlinks');
  if (container === undefined) return;
  const edit = input.edit;
  const slice = xml.slice(container.start, container.end);
  const patched = patchElements(slice, scanElements(slice, 'hyperlink'), (text) => {
    const match = /(\bref\s*=\s*)(["'])([^"']*)\2/.exec(text);
    const range = match === null ? null : parseRangeText(match[3] ?? '');
    if (match === null || range === null) return text;
    const transformed = transformRange(range, edit);
    if (transformed === null) return null;
    let next = text;
    if (!sameRange(range, transformed)) {
      next = replaceAttribute(next, match, formatRangeText(transformed));
    }
    // The internal `location` is a reference, not a label (file-compat §4): the
    // model has already transformed it, and this rewrite keeps the saved bytes
    // equal to the model. A location the transform refuses throws the same
    // unsafe-part block the plan raises, so the edit is never saved stale.
    const locationMatch = /(\blocation\s*=\s*)(["'])([^"']*)\2/.exec(next);
    if (locationMatch !== null) {
      const outcome = transformHyperlinkLocation(
        decodeXmlAttributeValue(locationMatch[3] ?? ''),
        edit,
        { workbook: input.workbook },
        edit.sheetId,
      );
      if (!outcome.ok) {
        throw new WorkbookSaveError(
          'unsafe-part',
          `${outcome.block.message} The edit was blocked.`,
        );
      }
      if (outcome.changed) next = replaceAttribute(next, locationMatch, outcome.location);
    }
    return next;
  });
  if (patched === slice) return;
  edits.push({ start: container.start, end: container.end, text: patched });
}

// ---------------------------------------------------------------------------
// Conditional formatting and data validation
// ---------------------------------------------------------------------------

function appendConditionalFormatEdits(
  xml: string,
  parsed: ParsedWorksheet,
  input: StructurePartInput,
  edits: TextEdit[],
): void {
  for (const container of parsed.rootChildren.filter(
    (child) => child.local === 'conditionalFormatting',
  )) {
    const slice = xml.slice(container.start, container.end);
    const hadSqref = /\bsqref\s*=/.test(slice);
    const updated = transformRangeBearingXml(
      slice,
      input.edit,
      { workbook: input.workbook },
      input.edit.sheetId,
    );
    if (hadSqref && !/\bsqref\s*=/.test(updated)) {
      edits.push({ start: container.start, end: container.end, text: '' });
      continue;
    }
    if (updated === slice) continue;
    edits.push({ start: container.start, end: container.end, text: updated });
  }
}

function appendValidationEdits(
  xml: string,
  parsed: ParsedWorksheet,
  input: StructurePartInput,
  edits: TextEdit[],
): void {
  for (const container of parsed.rootChildren.filter(
    (child) => child.local === 'dataValidations',
  )) {
    const slice = xml.slice(container.start, container.end);
    const elements = scanElements(slice, 'dataValidation');
    let removed = 0;
    const patched = patchElements(slice, elements, (text) => {
      const hadSqref = /\bsqref\s*=/.test(text);
      const next = transformRangeBearingXml(
        text,
        input.edit,
        { workbook: input.workbook },
        input.edit.sheetId,
      );
      if (hadSqref && !/\bsqref\s*=/.test(next)) {
        removed += 1;
        return null;
      }
      return next;
    });
    const withCount =
      removed > 0 ? updateCount(patched, Math.max(0, elements.length - removed)) : patched;
    if (withCount === slice) continue;
    edits.push({ start: container.start, end: container.end, text: withCount });
  }
}

// ---------------------------------------------------------------------------
// AutoFilter
// ---------------------------------------------------------------------------

function appendAutoFilterEdits(
  xml: string,
  parsed: ParsedWorksheet,
  edit: StructuralEdit,
  edits: TextEdit[],
): void {
  const container = parsed.rootChildren.find((child) => child.local === 'autoFilter');
  if (container === undefined) return;
  const slice = xml.slice(container.start, container.end);
  const refMatch = /(\bref\s*=\s*)(["'])([^"']*)\2/.exec(slice);
  const range = refMatch === null ? null : parseRangeText(refMatch[3] ?? '');
  const transformed = range === null ? null : transformRange(range, edit);
  if (range !== null && transformed === null) {
    edits.push({ start: container.start, end: container.end, text: '' });
    return;
  }

  let patched = slice;
  if (range !== null && transformed !== null && refMatch !== null) {
    if (!sameRange(range, transformed)) {
      patched = replaceAttribute(patched, refMatch, formatRangeText(transformed));
    }
    if (edit.axis === 'column') {
      const nextStartCol = transformed.startCol;
      patched = patchElements(patched, scanElements(patched, 'filterColumn'), (text) => {
        const match = /(\bcolId\s*=\s*)(["'])(\d*)\2/.exec(text);
        if (match === null) return text;
        const absolute = range.startCol + Number(match[3]);
        const next = transformCellAddress(0, absolute, edit);
        if (next === null) return null;
        const offset = next.col - nextStartCol;
        if (offset < 0) return null;
        if (offset === Number(match[3])) return text;
        return replaceAttribute(text, match, String(offset));
      });
    }
  }
  if (patched === slice) return;
  edits.push({ start: container.start, end: container.end, text: patched });
}

// ---------------------------------------------------------------------------
// Text helpers
// ---------------------------------------------------------------------------

interface TextEdit {
  readonly start: number;
  readonly end: number;
  readonly text: string;
}

interface ScannedElement {
  readonly start: number;
  readonly end: number;
  readonly openEnd: number;
  readonly text: string;
}

/** Apply `transform` to each element of a container; `null` removes it. */
function patchElements(
  slice: string,
  elements: readonly ScannedElement[],
  transform: (text: string) => string | null,
): string {
  let output = slice;
  for (const element of [...elements].reverse()) {
    const next = transform(element.text);
    if (next === null) {
      output = output.slice(0, element.start) + output.slice(element.end);
      continue;
    }
    if (next === element.text) continue;
    output = output.slice(0, element.start) + next + output.slice(element.end);
  }
  return output;
}

/**
 * Find every element with one local name in a fragment. The elements this
 * module scans never nest in themselves, so the first matching close tag ends
 * the element.
 */
function scanElements(xml: string, local: string): ScannedElement[] {
  const elements: ScannedElement[] = [];
  const pattern = new RegExp(`<(?:[\\w.-]+:)?${local}\\b`, 'g');
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(xml)) !== null) {
    const start = match.index;
    const openEnd = findTagEnd(xml, start);
    if (openEnd < 0) break;
    const openTag = xml.slice(start, openEnd);
    const qualified = /^<([^\s/>]+)/.exec(openTag)?.[1] ?? local;
    if (/\/\s*>$/.test(openTag)) {
      elements.push({ start, end: openEnd, openEnd, text: openTag });
      pattern.lastIndex = openEnd;
      continue;
    }
    const close = `</${qualified}>`;
    const closeIndex = xml.indexOf(close, openEnd);
    if (closeIndex < 0) break;
    const end = closeIndex + close.length;
    elements.push({ start, end, openEnd, text: xml.slice(start, end) });
    pattern.lastIndex = end;
  }
  return elements;
}

function findTagEnd(xml: string, start: number): number {
  let quoted: string | null = null;
  for (let index = start + 1; index < xml.length; index += 1) {
    const character = xml[index] ?? '';
    if (quoted !== null) {
      if (character === quoted) quoted = null;
      continue;
    }
    if (character === '"' || character === "'") {
      quoted = character;
      continue;
    }
    if (character === '>') return index + 1;
  }
  return -1;
}

function replaceAttribute(text: string, match: RegExpExecArray, value: string): string {
  const before = text.slice(0, match.index);
  const after = text.slice(match.index + match[0].length);
  return `${before}${match[1]}${match[2]}${escapeXmlAttributeValue(value)}${match[2]}${after}`;
}

function updateCount(text: string, count: number): string {
  const match = /(\bcount\s*=\s*)(["'])(\d*)\2/.exec(text);
  if (match === null) return text;
  return replaceAttribute(text, match, String(count));
}

/** Precompiled patterns for the attributes a structural renumber rewrites. */
const ATTRIBUTE_PATTERNS: ReadonlyMap<string, RegExp> = new Map([
  ['r', /(\sr\s*=\s*)(["'])([^"']*)\2/],
  ['ref', /(\bref\s*=\s*)(["'])([^"']*)\2/],
  ['min', /(\bmin\s*=\s*)(["'])([^"']*)\2/],
  ['max', /(\bmax\s*=\s*)(["'])([^"']*)\2/],
  ['count', /(\bcount\s*=\s*)(["'])([^"']*)\2/],
  ['sqref', /(\bsqref\s*=\s*)(["'])([^"']*)\2/],
  ['colId', /(\bcolId\s*=\s*)(["'])([^"']*)\2/],
]);

function setAttribute(openTag: string, name: string, value: string): string {
  const pattern = ATTRIBUTE_PATTERNS.get(name);
  if (pattern !== undefined && pattern.test(openTag)) {
    return openTag.replace(pattern, (_whole, prefix: string, quote: string) => {
      return `${prefix}${quote}${escapeXmlAttributeValue(value)}${quote}`;
    });
  }
  const close = openTag.lastIndexOf('>');
  if (close < 0) return openTag;
  const insertAt = openTag[close - 1] === '/' ? close - 1 : close;
  return `${openTag.slice(0, insertAt)} ${name}="${escapeXmlAttributeValue(value)}"${openTag.slice(
    insertAt,
  )}`;
}

function escapeXmlAttributeValue(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;');
}

/** The inverse of `escapeXmlAttributeValue`, in the order that cannot double-decode. */
function decodeXmlAttributeValue(value: string): string {
  return value
    .replaceAll('&quot;', '"')
    .replaceAll('&apos;', "'")
    .replaceAll('&lt;', '<')
    .replaceAll('&gt;', '>')
    .replaceAll('&amp;', '&');
}

function parseRangeText(text: string): GridRange | null {
  const trimmed = text.trim();
  if (trimmed === '') return null;
  const [startText, endText = startText] = trimmed.split(':');
  const start = parseCellAddress(startText ?? '');
  const end = parseCellAddress(endText);
  if (start === null || end === null) return null;
  return {
    startRow: Math.min(start.row, end.row),
    startCol: Math.min(start.col, end.col),
    endRow: Math.max(start.row, end.row),
    endCol: Math.max(start.col, end.col),
  };
}

function parseCellAddress(text: string): { row: number; col: number } | null {
  const match = /^([A-Z]{1,3})(\d{1,7})$/i.exec(text.replaceAll('$', ''));
  if (match === null) return null;
  const row = Number(match[2]) - 1;
  if (!Number.isInteger(row) || row < 0) return null;
  return { row, col: columnIndex(match[1]!) };
}

function columnIndex(letters: string): number {
  let value = 0;
  for (const character of letters.toUpperCase()) {
    value = value * 26 + (character.charCodeAt(0) - 64);
  }
  return value - 1;
}

function formatRangeText(range: GridRange): string {
  const start = formatCellRef(range.startRow, range.startCol);
  const end = formatCellRef(range.endRow, range.endCol);
  return start === end ? start : `${start}:${end}`;
}

function sameRange(left: GridRange, right: GridRange): boolean {
  return (
    left.startRow === right.startRow &&
    left.startCol === right.startCol &&
    left.endRow === right.endRow &&
    left.endCol === right.endCol
  );
}

function applyEdits(xml: string, edits: readonly TextEdit[]): string {
  // Build the output left to right so a large structural renumber is one pass
  // over the part, not one full copy per rewritten cell.
  const sorted = [...edits].sort((left, right) => left.start - right.start);
  let output = '';
  let cursor = 0;
  for (const edit of sorted) {
    if (edit.start < cursor) continue;
    output += xml.slice(cursor, edit.start);
    output += edit.text;
    cursor = edit.end;
  }
  return output + xml.slice(cursor);
}
