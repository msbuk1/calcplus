import type { DefinedName, SheetId, Workbook } from './types';
import { findAreaReferences, type AreaReference } from './references';

/**
 * Sheet rename and delete rewrite the display text of references that name the
 * changed sheet. The model still resolves references by sheet name at the
 * engine edge, so the display text is the only representation available; a
 * stable sheet id survives the rename and keeps undo exact (workbook-model
 * §3, command-system §3).
 */

export interface SheetFormulaRewrite {
  readonly sheetId: SheetId;
  readonly row: number;
  readonly col: number;
  readonly before: string;
  readonly after: string;
}

export interface SheetNameChange {
  readonly scope: 'workbook' | 'sheet';
  readonly sheetId: SheetId | null;
  readonly name: string;
  readonly before: string;
  readonly after: string;
}

export interface SheetReferencePlan {
  readonly formulaRewrites: readonly SheetFormulaRewrite[];
  readonly nameChanges: readonly SheetNameChange[];
  /** Set when a read-only formula group cannot be rewritten safely. */
  readonly blockedReason: string | null;
}

export function emptySheetReferencePlan(): SheetReferencePlan {
  return { formulaRewrites: [], nameChanges: [], blockedReason: null };
}

/**
 * Plan the formula and defined-name display rewrites for a sheet rename.
 * The new prefix keeps the reference's quoting style when that is safe; a
 * name that needs quotes is quoted so the formula stays parseable.
 */
export function planSheetRename(
  workbook: Workbook,
  sheetId: SheetId,
  newName: string,
): SheetReferencePlan {
  const sheet = workbook.sheets.find((candidate) => candidate.id === sheetId);
  if (sheet === undefined) return emptySheetReferencePlan();
  return planSheetReferences(workbook, sheet, 'renamed', sheet.names, (formula) =>
    rewriteSheetNameInFormula(formula, sheet.name, newName),
  );
}

/** Plan the rewrites that turn references to a deleted sheet into `#REF!`. */
export function planSheetDeleteReferences(
  workbook: Workbook,
  sheetId: SheetId,
): SheetReferencePlan {
  const sheet = workbook.sheets.find((candidate) => candidate.id === sheetId);
  if (sheet === undefined) return emptySheetReferencePlan();
  return planSheetReferences(workbook, sheet, 'deleted', sheet.names, (formula) =>
    rewriteReferencesToRefError(formula, sheet.name),
  );
}

/**
 * Rewrite only the sheet prefixes that name `oldName`. Everything else in the
 * formula, including string literals, stays byte-for-byte as authored.
 */
export function rewriteSheetNameInFormula(
  formula: string,
  oldName: string,
  newName: string,
): string {
  const references = findAreaReferences(formula);
  if (references.length === 0) return formula;
  let output = '';
  let cursor = 0;
  for (const reference of references) {
    if (!referenceNamesSheet(reference, formula, oldName)) continue;
    const prefixLength = reference.prefix?.length ?? 0;
    if (prefixLength === 0) continue;
    output += formula.slice(cursor, reference.start);
    output += formatSheetPrefix(newName, reference.prefix!.startsWith("'"));
    cursor = reference.start + prefixLength;
  }
  if (cursor === 0) return formula;
  return output + formula.slice(cursor);
}

/**
 * Replace each whole reference that names `sheetName` with `#REF!`, matching
 * Excel's result when a referenced sheet is deleted.
 */
export function rewriteReferencesToRefError(formula: string, sheetName: string): string {
  const references = findAreaReferences(formula);
  if (references.length === 0) return formula;
  let output = '';
  let cursor = 0;
  for (const reference of references) {
    if (!referenceNamesSheet(reference, formula, sheetName)) continue;
    output += formula.slice(cursor, reference.start);
    output += '#REF!';
    // The spill operator belongs to the reference; a deleted reference does
    // not keep it, so consume the trailing `#` rather than leave `#REF!#`.
    cursor = formula[reference.end] === '#' ? reference.end + 1 : reference.end;
  }
  if (cursor === 0) return formula;
  return output + formula.slice(cursor);
}

/** Whether a formula contains a reference whose sheet prefix names `sheetName`. */
export function formulaReferencesSheet(formula: string, sheetName: string): boolean {
  return findAreaReferences(formula).some((reference) =>
    referenceNamesSheet(reference, formula, sheetName),
  );
}

function planSheetReferences(
  workbook: Workbook,
  targetSheet: { readonly id: SheetId; readonly name: string },
  action: 'renamed' | 'deleted',
  scopedNames: readonly DefinedName[],
  transform: (formula: string) => string,
): SheetReferencePlan {
  const formulaRewrites: SheetFormulaRewrite[] = [];
  const nameChanges: SheetNameChange[] = [];
  const scopedNameSet = new Set(scopedNames.map((name) => name.name.toUpperCase()));

  for (const sheet of workbook.sheets) {
    // A deleted sheet's own formulas leave with it; a renamed sheet stays and
    // its self-references must follow the new label.
    if (action === 'deleted' && sheet.id === targetSheet.id) continue;
    for (const [key, cell] of sheet.cells) {
      if (cell.formula === null) continue;
      if (mentionsThreeDimensionalReference(cell.formula, targetSheet.name)) {
        return blocked(targetSheet, action, '3D reference');
      }
      if (mentionsSheetScopedName(cell.formula, targetSheet.name, scopedNameSet)) {
        return blocked(targetSheet, action, 'sheet-scoped defined name reference');
      }
      if (cell.formulaSource?.kind === 'shared-master' || cell.formulaSource?.kind === 'array') {
        if (formulaReferencesSheet(cell.formula, targetSheet.name)) {
          return blocked(targetSheet, action, 'read-only formula group');
        }
        continue;
      }
      if (cell.richText !== null && formulaReferencesSheet(cell.formula, targetSheet.name)) {
        return blocked(targetSheet, action, 'rich-text formula');
      }
      const after = transform(cell.formula);
      if (after === cell.formula) continue;
      const separator = key.indexOf(':');
      formulaRewrites.push({
        sheetId: sheet.id,
        row: Number(key.slice(0, separator)),
        col: Number(key.slice(separator + 1)),
        before: cell.formula,
        after,
      });
    }
  }

  for (const name of workbook.names) {
    if (mentionsThreeDimensionalReference(name.formula, targetSheet.name)) {
      return blocked(targetSheet, action, '3D reference');
    }
    if (mentionsSheetScopedName(name.formula, targetSheet.name, scopedNameSet)) {
      return blocked(targetSheet, action, 'sheet-scoped defined name reference');
    }
    const after = transform(name.formula);
    if (after === name.formula) continue;
    nameChanges.push({
      scope: 'workbook',
      sheetId: null,
      name: name.name,
      before: name.formula,
      after,
    });
  }
  for (const sheet of workbook.sheets) {
    if (action === 'deleted' && sheet.id === targetSheet.id) continue;
    for (const name of sheet.names) {
      if (mentionsThreeDimensionalReference(name.formula, targetSheet.name)) {
        return blocked(targetSheet, action, '3D reference');
      }
      if (mentionsSheetScopedName(name.formula, targetSheet.name, scopedNameSet)) {
        return blocked(targetSheet, action, 'sheet-scoped defined name reference');
      }
      const after = transform(name.formula);
      if (after === name.formula) continue;
      nameChanges.push({
        scope: 'sheet',
        sheetId: sheet.id,
        name: name.name,
        before: name.formula,
        after,
      });
    }
  }

  return { formulaRewrites, nameChanges, blockedReason: null };
}

function blocked(
  sheet: { readonly name: string },
  action: 'renamed' | 'deleted',
  construct: string,
): SheetReferencePlan {
  return {
    formulaRewrites: [],
    nameChanges: [],
    blockedReason: `${sheet.name} is used by a ${construct}, which cannot be rewritten in this milestone; the sheet was not ${action}.`,
  };
}

function referenceNamesSheet(
  reference: AreaReference,
  formula: string,
  sheetName: string,
): boolean {
  if (reference.prefix === null || reference.sheetName === null) return false;
  if (reference.start > 0 && formula[reference.start - 1] === ']') {
    // `[Book1]Sheet1!A1` is an external-workbook reference; never rewrite it.
    return false;
  }
  return reference.sheetName.toUpperCase() === sheetName.toUpperCase();
}

const THREE_DIMENSIONAL_REFERENCE =
  /(?:'(?:[^']|'')*'|[A-Za-z_][\w.]*)\s*:\s*(?:'(?:[^']|'')*'|[A-Za-z_][\w.]*)\s*!/g;

/**
 * A 3D reference (`Sheet1:Sheet3!A1`) has no sheet-id representation yet, so
 * an affected rename or delete is blocked rather than half-rewritten.
 */
export function mentionsThreeDimensionalReference(formula: string, sheetName: string): boolean {
  const upper = sheetName.toUpperCase();
  const stripped = formula.replace(/"(?:[^"]|"")*"/g, '""');
  THREE_DIMENSIONAL_REFERENCE.lastIndex = 0;
  let match: RegExpExecArray | null;
  while ((match = THREE_DIMENSIONAL_REFERENCE.exec(stripped)) !== null) {
    if (match.index > 0 && stripped[match.index - 1] === ']') continue;
    const reference = match[0].slice(0, -1);
    for (const part of reference.split(':')) {
      if (decodeSheetToken(part.trim()).toUpperCase() === upper) return true;
    }
  }
  return false;
}

const SHEET_SCOPED_NAME_REFERENCE =
  /((?:'(?:[^']|'')*'|[A-Za-z_][\w.]*)\s*!)\s*([A-Za-z_\\\u00C0-\u02AF][\w.\u00C0-\u02AF]*)/g;

/**
 * A sheet-scoped defined name is referenced as `Sheet1!LocalName`. The name
 * has no stable-id form in authored text yet, so a rename/delete that would
 * change that prefix is blocked rather than left stale.
 */
function mentionsSheetScopedName(
  formula: string,
  sheetName: string,
  scopedNameSet: ReadonlySet<string>,
): boolean {
  if (scopedNameSet.size === 0) return false;
  const stripped = formula.replace(/"(?:[^"]|"")*"/g, '""');
  SHEET_SCOPED_NAME_REFERENCE.lastIndex = 0;
  let match: RegExpExecArray | null;
  while ((match = SHEET_SCOPED_NAME_REFERENCE.exec(stripped)) !== null) {
    if (match.index > 0 && stripped[match.index - 1] === ']') continue;
    const prefix = match[1]!.slice(0, -1).trim();
    if (decodeSheetToken(prefix).toUpperCase() !== sheetName.toUpperCase()) continue;
    if (scopedNameSet.has(match[2]!.toUpperCase())) return true;
  }
  return false;
}

function decodeSheetToken(token: string): string {
  if (token.startsWith("'") && token.endsWith("'")) {
    return token.slice(1, -1).replaceAll("''", "'");
  }
  return token;
}

function formatSheetPrefix(name: string, wasQuoted: boolean): string {
  if (!wasQuoted && !sheetNameNeedsQuotes(name)) return name;
  return `'${name.replaceAll("'", "''")}'`;
}

function sheetNameNeedsQuotes(name: string): boolean {
  if (!/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name)) return true;
  // A name that is a valid A1 address, R1C1 address or bare column/row is
  // quoted so the parser does not read it as part of the reference.
  if (/^R\d+C\d+$/i.test(name)) return true;
  return /^(?:[A-Za-z]{1,3}\d+|[A-Za-z]{1,3}|\d+)$/.test(name);
}
