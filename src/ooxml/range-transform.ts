import type { GridRange, SheetId, Workbook } from '@model/types';
import {
  planRangeTransform,
  type MergeRewrite,
  type RangeTransformPlan,
  type StructuralEdit,
} from '@model/range-transform';
import type { WorkbookPackage } from './types';
import { parseWorksheetXml } from './worksheet-xml';
import { decodeXml, encodeXml, escapeXmlText, isWellFormedXml, validatePatchedPart } from './xml';

/**
 * M3C prototype of the OOXML range transform.
 *
 * It rewrites the range-bearing XML a structural edit touches — formula text,
 * merged ranges and defined names — against the original part bytes. Cell
 * movement and the full Tier B anchor pass remain M5/M6 work. The prototype is
 * deliberately read-only for the model: it returns a patched part map for
 * measurement and review, not a save path.
 */

export interface RangeTransformPrototypeInput {
  readonly workbook: Workbook;
  readonly package: WorkbookPackage;
  readonly edit: StructuralEdit;
}

export interface RangeTransformPrototypeReport {
  readonly patchedParts: readonly string[];
  readonly formulaRewrites: number;
  readonly mergesRewritten: number;
  readonly mergesRemoved: number;
  readonly namesRewritten: number;
  readonly groupedFormulaCount: number;
  readonly unchanged: boolean;
  /** Wall-clock time for the planning and part patching, for the M3 gate. */
  readonly elapsedMs: number;
}

export interface RangeTransformPrototypeResult {
  readonly entries: Map<string, Uint8Array>;
  readonly plan: RangeTransformPlan;
  readonly report: RangeTransformPrototypeReport;
}

export function applyRangeTransformPrototype(
  input: RangeTransformPrototypeInput,
): RangeTransformPrototypeResult {
  const started = now();
  const plan = planRangeTransform(input.workbook, input.edit);
  const entries = new Map(input.package.entries);
  const patchedParts: string[] = [];
  let formulaRewrites = 0;
  let mergesRewritten = 0;
  let mergesRemoved = 0;
  let namesRewritten = 0;

  for (const sheet of input.workbook.sheets) {
    const bytes = entries.get(sheet.partPath);
    if (bytes === undefined) continue;
    const result = patchWorksheetPart(bytes, sheet.partPath, sheet.id, sheet.name, plan, input);
    if (!result.changed) continue;
    entries.set(sheet.partPath, result.bytes);
    patchedParts.push(sheet.partPath);
    formulaRewrites += result.formulaRewrites;
    mergesRewritten += result.mergesRewritten;
    mergesRemoved += result.mergesRemoved;
  }

  if (plan.nameRewrites.length > 0) {
    const bytes = entries.get(input.package.workbookPartPath);
    if (bytes !== undefined) {
      const patched = patchDefinedNames(
        bytes,
        input.package.workbookPartPath,
        plan,
        input.workbook,
      );
      if (patched.changed) {
        entries.set(input.package.workbookPartPath, patched.bytes);
        patchedParts.push(input.package.workbookPartPath);
        namesRewritten = patched.namesRewritten;
      }
    }
  }

  for (const partName of patchedParts) {
    const patchedBytes = entries.get(partName);
    if (patchedBytes !== undefined) validatePatchedPart(patchedBytes, partName);
  }

  return {
    entries,
    plan,
    report: {
      patchedParts,
      formulaRewrites,
      mergesRewritten,
      mergesRemoved,
      namesRewritten,
      groupedFormulaCount: plan.groupedFormulaCount,
      unchanged: patchedParts.length === 0,
      elapsedMs: Math.max(0, now() - started),
    },
  };
}

function now(): number {
  return typeof performance !== 'undefined' ? performance.now() : Date.now();
}

interface WorksheetPartResult {
  readonly bytes: Uint8Array;
  readonly changed: boolean;
  readonly formulaRewrites: number;
  readonly mergesRewritten: number;
  readonly mergesRemoved: number;
}

function patchWorksheetPart(
  sourceBytes: Uint8Array,
  partName: string,
  sheetId: SheetId,
  sheetName: string,
  plan: RangeTransformPlan,
  input: RangeTransformPrototypeInput,
): WorksheetPartResult {
  const parsed = parseWorksheetXml(sourceBytes, {
    sharedStrings: input.package.sharedStrings,
    partPath: partName,
    name: sheetName,
  });
  const rewritesByKey = new Map(
    plan.formulaRewrites.map((rewrite) => [
      `${rewrite.sheetId}\u0000${rewrite.row}:${rewrite.col}`,
      rewrite,
    ]),
  );
  const formulaEdits: { start: number; end: number; text: string }[] = [];

  for (const row of parsed.rows.values()) {
    for (const cell of row.cells.values()) {
      if (cell.formulaText === null || cell.formulaKind !== null) continue;
      const rewrite = rewritesByKey.get(`${sheetId}\u0000${cell.row}:${cell.col}`);
      if (rewrite === undefined) continue;
      const cellXml = parsed.xml.slice(cell.start, cell.end);
      const fTag = qualified(cell.prefix, 'f');
      const formula = new RegExp(`<${fTag}(?:\\s[^>]*)?>([\\s\\S]*?)</${fTag}>`).exec(cellXml);
      if (formula === null) continue;
      const openTag = formula[0].slice(0, formula[0].indexOf('>') + 1);
      formulaEdits.push({
        start: cell.start + formula.index,
        end: cell.start + formula.index + formula[0].length,
        text: `${openTag}${escapeXmlText(rewrite.after)}</${fTag}>`,
      });
    }
  }

  const formulaXml = applyEdits(parsed.xml, formulaEdits);
  const mergeResult = patchMerges(formulaXml, plan.mergeRewrites, plan.edit.sheetId === sheetId);
  if (!isWellFormedXml(mergeResult.xml)) {
    throw new Error(`The range-transform prototype produced invalid XML in ${partName}`);
  }
  return {
    bytes: encodeXml(mergeResult.xml),
    changed: formulaEdits.length > 0 || mergeResult.changed,
    formulaRewrites: formulaEdits.length,
    mergesRewritten: mergeResult.rewritten,
    mergesRemoved: mergeResult.removed,
  };
}

interface MergePatchResult {
  readonly xml: string;
  readonly changed: boolean;
  readonly rewritten: number;
  readonly removed: number;
}

function patchMerges(
  xml: string,
  rewrites: readonly MergeRewrite[],
  isTargetSheet: boolean,
): MergePatchResult {
  if (!isTargetSheet || rewrites.length === 0) {
    return { xml, changed: false, rewritten: 0, removed: 0 };
  }
  const rootPrefix = partNamePrefix(xml);
  const mergeTag = qualified(rootPrefix, 'mergeCell');
  const mergeCellsTag = qualified(rootPrefix, 'mergeCells');
  const pattern = new RegExp(
    `<${mergeTag}\\b([^>]*?)\\bref\\s*=\\s*(["'])([^"']*)\\2([^>]*?)\\/>`,
    'g',
  );
  const byBefore = new Map<string, GridRange | null>();
  for (const rewrite of rewrites) {
    byBefore.set(rangeKey(rewrite.before), rewrite.after);
  }

  let rewritten = 0;
  let removed = 0;
  const next = xml.replace(
    pattern,
    (whole: string, before: string, quote: string, ref: string, after: string) => {
      const range = parseRangeRef(ref);
      if (range === null) return whole;
      const key = rangeKey(range);
      if (!byBefore.has(key)) return whole;
      const replacement = byBefore.get(key) ?? null;
      if (replacement === null) {
        removed += 1;
        return '';
      }
      const text = formatRangeRef(replacement);
      if (text === ref) return whole;
      rewritten += 1;
      return `<${mergeTag}${before}ref=${quote}${text}${quote}${after}/>`;
    },
  );
  const countPattern = new RegExp(`(<${mergeCellsTag}\\b[^>]*?\\bcount=)(["'])(\\d+)\\2`);
  const kept = countMergeCells(next, mergeTag);
  const countPatched = next.replace(
    countPattern,
    (_whole: string, before: string, quote: string) => `${before}${quote}${kept}${quote}`,
  );
  return {
    xml: countPatched,
    changed: next !== xml || countPatched !== xml,
    rewritten,
    removed,
  };
}

function countMergeCells(xml: string, mergeTag: string): number {
  const pattern = new RegExp(`<${mergeTag}\\b`, 'g');
  return [...xml.matchAll(pattern)].length;
}

interface DefinedNamePatchResult {
  readonly bytes: Uint8Array;
  readonly changed: boolean;
  readonly namesRewritten: number;
}

function patchDefinedNames(
  sourceBytes: Uint8Array,
  partName: string,
  plan: RangeTransformPlan,
  workbook: Workbook,
): DefinedNamePatchResult {
  const xml = decodeXml(sourceBytes);
  const rewrites = new Map<string, string>();
  for (const rewrite of plan.nameRewrites) {
    rewrites.set(
      nameKey(rewrite.scope === 'workbook' ? null : rewrite.sheetId, rewrite.name),
      rewrite.after,
    );
  }
  let namesRewritten = 0;
  const pattern = /<((?:[\w.-]+:)?)definedName\b([^>]*)>([\s\S]*?)<\/\1definedName>/g;
  const patched = xml.replace(pattern, (whole: string, prefix: string, attributes: string) => {
    const parsedAttributes = parseAttributes(attributes);
    const name = parsedAttributes.get('name');
    if (name === undefined) return whole;
    const localSheetId = parsedAttributes.get('localSheetId');
    const sheetId = localSheetId === undefined ? null : workbook.sheets[Number(localSheetId)]?.id;
    if (sheetId === undefined) return whole;
    const after = rewrites.get(nameKey(sheetId, name));
    if (after === undefined) return whole;
    namesRewritten += 1;
    return `<${prefix}definedName${attributes}>${escapeXmlText(after)}</${prefix}definedName>`;
  });
  if (namesRewritten === 0) {
    return { bytes: sourceBytes, changed: false, namesRewritten: 0 };
  }
  if (!isWellFormedXml(patched)) {
    throw new Error(`The range-transform prototype produced invalid defined names in ${partName}`);
  }
  return { bytes: encodeXml(patched), changed: true, namesRewritten };
}

function applyEdits(
  xml: string,
  edits: readonly { start: number; end: number; text: string }[],
): string {
  const sorted = [...edits].sort((left, right) => right.start - left.start);
  let output = xml;
  for (const edit of sorted) {
    output = output.slice(0, edit.start) + edit.text + output.slice(edit.end);
  }
  return output;
}

function partNamePrefix(xml: string): string {
  const match = /<([A-Za-z_][\w.-]*:)?worksheet\b/.exec(xml);
  return match?.[1]?.slice(0, -1) ?? '';
}

function qualified(prefix: string, local: string): string {
  return prefix === '' ? local : `${prefix}:${local}`;
}

function parseRangeRef(ref: string): GridRange | null {
  const [startText, endText = startText] = ref.replaceAll('$', '').split(':');
  const start = parseCell(startText ?? '');
  const end = parseCell(endText ?? '');
  if (start === null || end === null) return null;
  return {
    startRow: Math.min(start.row, end.row),
    startCol: Math.min(start.col, end.col),
    endRow: Math.max(start.row, end.row),
    endCol: Math.max(start.col, end.col),
  };
}

function parseCell(ref: string): { row: number; col: number } | null {
  const match = /^([A-Z]{1,3})(\d{1,7})$/i.exec(ref);
  if (match === null) return null;
  let col = 0;
  for (const character of match[1]!.toUpperCase()) {
    col = col * 26 + (character.charCodeAt(0) - 64);
  }
  const row = Number(match[2]) - 1;
  if (col < 1 || row < 0) return null;
  return { row, col: col - 1 };
}

function formatRangeRef(range: GridRange): string {
  return `${formatRef(range.startRow, range.startCol)}:${formatRef(range.endRow, range.endCol)}`;
}

function formatRef(row: number, col: number): string {
  return `${columnLetters(col)}${row + 1}`;
}

function columnLetters(column: number): string {
  let value = column;
  let result = '';
  while (value >= 0) {
    const digit = value % 26;
    result = String.fromCharCode(65 + digit) + result;
    value = Math.floor(value / 26) - 1;
  }
  return result;
}

function rangeKey(range: GridRange): string {
  return `${range.startRow}:${range.startCol}:${range.endRow}:${range.endCol}`;
}

function nameKey(sheetId: SheetId | null, name: string): string {
  return `${sheetId ?? 'workbook'}\u0000${name.toUpperCase()}`;
}

function parseAttributes(text: string): Map<string, string> {
  const attributes = new Map<string, string>();
  const pattern = /([A-Za-z_][\w.:-]*)\s*=\s*(["'])(.*?)\2/g;
  for (const match of text.matchAll(pattern)) {
    const name = match[1] ?? '';
    const local = name.includes(':') ? name.slice(name.lastIndexOf(':') + 1) : name;
    attributes.set(local, decodeEntities(match[3] ?? ''));
  }
  return attributes;
}

function decodeEntities(text: string): string {
  return text
    .replaceAll('&quot;', '"')
    .replaceAll('&apos;', "'")
    .replaceAll('&lt;', '<')
    .replaceAll('&gt;', '>')
    .replaceAll('&amp;', '&');
}
