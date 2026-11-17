import type { SheetId, Workbook } from '@model/types';
import { parseWorkbookXml, type WorkbookSheetReference } from '@ooxml/workbook-xml';

export type SheetVisibility = 'visible' | 'hidden' | 'veryHidden';

interface VisibilitySnapshot {
  /** The workbook part this mapping came from. */
  readonly path: string;
  /** The workbook part bytes this mapping came from. */
  readonly source: Uint8Array;
  readonly bySheetId: ReadonlyMap<SheetId, SheetVisibility>;
}

const snapshots = new WeakMap<Workbook, VisibilitySnapshot>();
const WORKBOOK_TAG = /<(?:(?:[A-Za-z_][\w.-]*):)?workbook(?:\s|>)/;

/**
 * Imported sheet visibility, keyed by the model's stable sheet id.
 *
 * The M4B model has no visibility field yet, so the view reads the preserved
 * `state` attribute from the workbook part. This is read-only metadata: it
 * never mutates the workbook, never enters a transaction and never marks the
 * document dirty. Hide/unhide authoring belongs to V2.
 */
export function sheetVisibilityById(workbook: Workbook): ReadonlyMap<SheetId, SheetVisibility> {
  const cached = snapshots.get(workbook);
  if (cached !== undefined && workbook.sourceParts.get(cached.path) === cached.source) {
    return cached.bySheetId;
  }
  const part = findWorkbookPart(workbook);
  if (part === null) return new Map();
  const references = readSheetReferences(part.bytes, part.path);
  if (references === null) return new Map();
  const bySheetId = mapReferencesToSheets(workbook, references);
  snapshots.set(workbook, { path: part.path, source: part.bytes, bySheetId });
  return bySheetId;
}

export function sheetVisibility(workbook: Workbook, sheetId: SheetId): SheetVisibility {
  return sheetVisibilityById(workbook).get(sheetId) ?? 'visible';
}

export function isSheetVisible(workbook: Workbook, sheetId: SheetId): boolean {
  return sheetVisibility(workbook, sheetId) === 'visible';
}

export function visibleSheetCount(workbook: Workbook): number {
  const visibility = sheetVisibilityById(workbook);
  let count = 0;
  for (const sheet of workbook.sheets) {
    if ((visibility.get(sheet.id) ?? 'visible') === 'visible') count += 1;
  }
  return count;
}

export function firstVisibleSheetId(workbook: Workbook): SheetId | null {
  const visibility = sheetVisibilityById(workbook);
  for (const sheet of workbook.sheets) {
    if ((visibility.get(sheet.id) ?? 'visible') === 'visible') return sheet.id;
  }
  return null;
}

function findWorkbookPart(
  workbook: Workbook,
): { readonly path: string; readonly bytes: Uint8Array } | null {
  const decoder = new TextDecoder();
  for (const [path, bytes] of workbook.sourceParts) {
    // The workbook part is the only part whose root element is `<workbook>`.
    // Inspecting a short prefix avoids decoding large worksheet parts.
    const prefix = decoder.decode(bytes.subarray(0, Math.min(bytes.length, 4096)));
    if (WORKBOOK_TAG.test(prefix)) return { path, bytes };
  }
  return null;
}

function readSheetReferences(
  bytes: Uint8Array,
  partPath: string,
): readonly WorkbookSheetReference[] | null {
  try {
    return parseWorkbookXml(bytes, partPath).sheets;
  } catch {
    // A read-only projection must not break a workbook that already opened.
    return null;
  }
}

/**
 * Map the workbook part's ordered sheet elements onto the model's sheet ids.
 * Names are matched first so an inserted sheet does not shift an imported
 * hidden state onto its neighbour; position is the fallback for a rename that
 * happened before the first tab render.
 */
function mapReferencesToSheets(
  workbook: Workbook,
  references: readonly WorkbookSheetReference[],
): ReadonlyMap<SheetId, SheetVisibility> {
  const bySheetId = new Map<SheetId, SheetVisibility>();
  const claimedReferences = new Set<number>();
  for (const [index, reference] of references.entries()) {
    const sheet = workbook.sheets.find(
      (candidate) => candidate.name === reference.name && !bySheetId.has(candidate.id),
    );
    if (sheet === undefined) continue;
    bySheetId.set(sheet.id, reference.state);
    claimedReferences.add(index);
  }
  for (const [index, reference] of references.entries()) {
    if (claimedReferences.has(index)) continue;
    const sheet = workbook.sheets[index];
    if (sheet === undefined || bySheetId.has(sheet.id)) continue;
    bySheetId.set(sheet.id, reference.state);
  }
  return bySheetId;
}
