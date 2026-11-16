import type {
  AlignmentStyle,
  BorderEdgeStyle,
  BorderStyle,
  FillStyle,
  FontStyle,
  NumberFormatKind,
  StyleRecord,
  Workbook,
} from '@model/types';
import {
  BUILTIN_NUMBER_FORMATS,
  DEFAULT_ALIGNMENT,
  DEFAULT_BORDER,
  DEFAULT_FILL,
  DEFAULT_FONT,
} from '@model/style';
import { escapeXmlAttribute, decodeXml, parseXml, openTagStart } from '../xml';
import { parseStylesXml, type StylesXmlInfo } from '../styles-xml';
import { WorkbookSaveError } from '../errors';

/**
 * Append-only `styles.xml` writer (file-compat §3, number-formats §6).
 *
 * Existing font, fill, border, number-format and cell-format records are never
 * renumbered. A new style reuses a visually identical existing sub-record when
 * it can, otherwise appends a new one at the end of its section. The cellXfs
 * entry is appended last.
 */

export interface AppendStylesInput {
  readonly partName: string;
  readonly sourceBytes: Uint8Array;
  readonly workbook: Workbook;
}

export interface AppendedStyleIds {
  readonly styleId: number;
  readonly fontId: number;
  readonly fillId: number;
  readonly borderId: number;
  readonly xfId: number;
  readonly numFmtId: number;
  readonly numFmtKind: NumberFormatKind;
  readonly rawXml: string;
}

export interface AppendStylesResult {
  readonly bytes: Uint8Array;
  readonly appendedStyles: readonly AppendedStyleIds[];
  readonly appendedFonts: number;
  readonly appendedFills: number;
  readonly appendedBorders: number;
  readonly appendedNumberFormats: number;
}

export function appendStylesXml(input: AppendStylesInput): AppendStylesResult {
  const source = parseStylesXml(input.sourceBytes, input.partName);
  const appendedRecords = input.workbook.styles.filter(
    (style) => style.id >= source.sourceStyleCount,
  );
  for (const [offset, style] of appendedRecords.entries()) {
    if (style.id !== source.sourceStyleCount + offset) {
      throw new WorkbookSaveError(
        'unsupported-edit',
        'The appended style ids are not contiguous; the file was not written.',
      );
    }
  }
  if (appendedRecords.length === 0) {
    return {
      bytes: input.sourceBytes,
      appendedStyles: [],
      appendedFonts: 0,
      appendedFills: 0,
      appendedBorders: 0,
      appendedNumberFormats: 0,
    };
  }

  const tables = createAppendTables(source);
  const assigned: AppendedStyleIds[] = [];

  for (const style of appendedRecords) {
    const font = tables.fontIdFor(style.font);
    const fill = tables.fillIdFor(style.fill);
    const border = tables.borderIdFor(style.border);
    const numberFormat = tables.numberFormatIdFor(style);
    const xfXml = serializeXf(style, {
      fontId: font,
      fillId: fill,
      borderId: border,
      numFmtId: numberFormat.numFmtId,
    });
    tables.cellXfs.push(xfXml);
    assigned.push({
      styleId: style.id,
      fontId: font,
      fillId: fill,
      borderId: border,
      xfId: style.xfId,
      numFmtId: numberFormat.numFmtId,
      numFmtKind: numberFormat.kind,
      rawXml: xfXml,
    });
  }

  const xml = decodeXml(input.sourceBytes);
  const patched = patchStylesSections(xml, input.partName, source, tables);
  return {
    bytes: new TextEncoder().encode(patched),
    appendedStyles: assigned,
    appendedFonts: tables.appendedFonts,
    appendedFills: tables.appendedFills,
    appendedBorders: tables.appendedBorders,
    appendedNumberFormats: tables.appendedNumberFormats,
  };
}

/** Write the assigned sub-record ids back into the model's append-only table. */
export function adoptAppendedStyleIds(
  workbook: Workbook,
  assigned: readonly AppendedStyleIds[],
): void {
  for (const entry of assigned) {
    const index = workbook.styles.findIndex((style) => style.id === entry.styleId);
    if (index < 0) continue;
    const current = workbook.styles[index]!;
    workbook.styles[index] = {
      ...current,
      fontId: entry.fontId,
      fillId: entry.fillId,
      borderId: entry.borderId,
      xfId: entry.xfId,
      numFmtId: entry.numFmtId,
      numFmtKind: entry.numFmtKind,
      rawXml: entry.rawXml,
    };
  }
}

interface AppendTables {
  readonly fonts: FontStyle[];
  readonly fills: FillStyle[];
  readonly borders: BorderStyle[];
  readonly fontXml: string[];
  readonly fillXml: string[];
  readonly borderXml: string[];
  readonly numFmtXml: string[];
  readonly cellXfs: string[];
  readonly numberFormatCodes: Map<string, number>;
  readonly appendedNumberFormatIds: Set<number>;
  nextNumberFormatId: number;
  readonly sourceFontCount: number;
  readonly sourceFillCount: number;
  readonly sourceBorderCount: number;
  appendedFonts: number;
  appendedFills: number;
  appendedBorders: number;
  appendedNumberFormats: number;
  fontIdFor(font: FontStyle): number;
  fillIdFor(fill: FillStyle): number;
  borderIdFor(border: BorderStyle): number;
  numberFormatIdFor(style: StyleRecord): { numFmtId: number; kind: NumberFormatKind };
}

function createAppendTables(source: StylesXmlInfo): AppendTables {
  const fonts = [...source.fonts];
  const fills = [...source.fills];
  const borders = [...source.borders];
  const fontXml: string[] = [];
  const fillXml: string[] = [];
  const borderXml: string[] = [];
  const numFmtXml: string[] = [];
  const cellXfs: string[] = [];
  const fontIds = new Map<string, number>();
  const fillIds = new Map<string, number>();
  const borderIds = new Map<string, number>();
  fonts.forEach((font, index) => {
    if (!fontIds.has(fontSignature(font))) fontIds.set(fontSignature(font), index);
  });
  fills.forEach((fill, index) => {
    if (!fillIds.has(fillSignature(fill))) fillIds.set(fillSignature(fill), index);
  });
  borders.forEach((border, index) => {
    if (!borderIds.has(borderSignature(border))) borderIds.set(borderSignature(border), index);
  });

  // A styles part always needs the default font and the two reserved fills.
  let preludeFonts = 0;
  let preludeFills = 0;
  let preludeBorders = 0;
  if (fonts.length === 0) {
    fonts.push(DEFAULT_FONT);
    fontXml.push(serializeFont(DEFAULT_FONT));
    fontIds.set(fontSignature(DEFAULT_FONT), 0);
    preludeFonts = 1;
  }
  if (fills.length === 0) {
    fills.push(DEFAULT_FILL, { pattern: 'gray125', color: null, backgroundColor: null });
    fillXml.push(serializeFill(DEFAULT_FILL), serializeFill(fills[1]!));
    fillIds.set(fillSignature(DEFAULT_FILL), 0);
    fillIds.set(fillSignature(fills[1]!), 1);
    preludeFills = 2;
  }
  if (borders.length === 0) {
    borders.push(DEFAULT_BORDER);
    borderXml.push(serializeBorder(DEFAULT_BORDER));
    borderIds.set(borderSignature(DEFAULT_BORDER), 0);
    preludeBorders = 1;
  }

  const numberFormatCodes = new Map<string, number>();
  for (const [id, code] of source.customFormats) {
    if (!numberFormatCodes.has(code)) numberFormatCodes.set(code, id);
  }
  const appendedNumberFormatIds = new Set<number>();
  let nextNumberFormatId = nextCustomNumberFormatId(source);

  const tables: AppendTables = {
    fonts,
    fills,
    borders,
    fontXml,
    fillXml,
    borderXml,
    numFmtXml,
    cellXfs,
    numberFormatCodes,
    appendedNumberFormatIds,
    nextNumberFormatId,
    sourceFontCount: source.fonts.length,
    sourceFillCount: source.fills.length,
    sourceBorderCount: source.borders.length,
    appendedFonts: preludeFonts,
    appendedFills: preludeFills,
    appendedBorders: preludeBorders,
    appendedNumberFormats: 0,
    fontIdFor(font) {
      const signature = fontSignature(font);
      const existing = fontIds.get(signature);
      if (existing !== undefined) return existing;
      const id = fonts.length;
      fonts.push(font);
      fontXml.push(serializeFont(font));
      fontIds.set(signature, id);
      tables.appendedFonts += 1;
      return id;
    },
    fillIdFor(fill) {
      const signature = fillSignature(fill);
      const existing = fillIds.get(signature);
      if (existing !== undefined) return existing;
      const id = fills.length;
      fills.push(fill);
      fillXml.push(serializeFill(fill));
      fillIds.set(signature, id);
      tables.appendedFills += 1;
      return id;
    },
    borderIdFor(border) {
      const signature = borderSignature(border);
      const existing = borderIds.get(signature);
      if (existing !== undefined) return existing;
      const id = borders.length;
      borders.push(border);
      borderXml.push(serializeBorder(border));
      borderIds.set(signature, id);
      tables.appendedBorders += 1;
      return id;
    },
    numberFormatIdFor(style) {
      if (style.numFmtKind === 'custom' && style.numFmtCode !== null) {
        // A style that already points at a source custom record keeps that id.
        if (style.numFmtId !== 0 && source.customFormats.get(style.numFmtId) === style.numFmtCode) {
          return { numFmtId: style.numFmtId, kind: 'custom' };
        }
        const existing = numberFormatCodes.get(style.numFmtCode);
        if (existing !== undefined) return { numFmtId: existing, kind: 'custom' };
        const id = nextNumberFormatId;
        nextNumberFormatId += 1;
        numberFormatCodes.set(style.numFmtCode, id);
        appendedNumberFormatIds.add(id);
        numFmtXml.push(serializeNumberFormat(id, style.numFmtCode));
        tables.appendedNumberFormats += 1;
        return { numFmtId: id, kind: 'custom' };
      }
      if (style.numFmtKind === 'builtin' || style.numFmtKind === 'unknown') {
        return { numFmtId: style.numFmtId, kind: style.numFmtKind };
      }
      return { numFmtId: 0, kind: 'general' };
    },
  };
  return tables;
}

function nextCustomNumberFormatId(source: StylesXmlInfo): number {
  let maximum = 163;
  for (const id of source.customFormats.keys()) {
    if (id > maximum) maximum = id;
  }
  for (const id of source.styles.map((style) => style.numFmtId)) {
    if (styleIsCustomId(source, id) && id > maximum) maximum = id;
  }
  return maximum + 1;
}

function styleIsCustomId(source: StylesXmlInfo, id: number): boolean {
  return source.customFormats.has(id) || (id > 81 && !(id in BUILTIN_NUMBER_FORMATS));
}

interface SectionRange {
  readonly local: string;
  readonly start: number;
  readonly openEnd: number;
  readonly contentEnd: number;
  readonly end: number;
  readonly selfClosing: boolean;
  readonly prefix: string;
}

const SECTION_ORDER = [
  'numFmts',
  'fonts',
  'fills',
  'borders',
  'cellStyleXfs',
  'cellXfs',
  'cellStyles',
  'dxfs',
  'tableStyles',
  'colors',
  'extLst',
] as const;

function patchStylesSections(
  xml: string,
  partName: string,
  source: StylesXmlInfo,
  tables: AppendTables,
): string {
  const ranges = collectSectionRanges(xml, partName);
  const styleSheet = findRootRange(xml, partName);
  const edits: { start: number; end: number; text: string; kind: 'replace' | 'insert' }[] = [];
  const itemsBySection: ReadonlyMap<string, readonly string[]> = new Map([
    ['numFmts', tables.numFmtXml],
    ['fonts', tables.fontXml],
    ['fills', tables.fillXml],
    ['borders', tables.borderXml],
    ['cellXfs', tables.cellXfs],
  ]);
  const counts = new Map<string, number>([
    ['numFmts', source.customFormats.size + tables.appendedNumberFormats],
    ['fonts', source.fonts.length + tables.appendedFonts],
    ['fills', source.fills.length + tables.appendedFills],
    ['borders', source.borders.length + tables.appendedBorders],
    ['cellXfs', source.sourceStyleCount + tables.cellXfs.length],
  ]);

  for (const [local, items] of itemsBySection) {
    if (items.length === 0) continue;
    const range = ranges.get(local);
    const count = counts.get(local) ?? items.length;
    const body = items.map((item) => `\n    ${item}`).join('');
    if (range === undefined) {
      const prefix = styleSheet.prefix;
      const qualified = qualify(prefix, local);
      const position = missingSectionInsertPosition(local, ranges, styleSheet, xml);
      edits.push({
        start: position,
        end: position,
        text: `<${qualified} count="${count}">${body}\n  </${qualified}>`,
        kind: 'insert',
      });
      continue;
    }
    const openText = xml.slice(range.start, range.openEnd);
    if (range.selfClosing) {
      const open = withCount(openText.replace(/\/>\s*$/, '>'), count);
      edits.push({
        start: range.start,
        end: range.end,
        text: `${open}${body}\n  </${qualify(range.prefix, local)}>`,
        kind: 'replace',
      });
    } else {
      edits.push({
        start: range.start,
        end: range.openEnd,
        text: withCount(openText, count),
        kind: 'replace',
      });
      edits.push({ start: range.contentEnd, end: range.contentEnd, text: body, kind: 'insert' });
    }
  }

  return applyEdits(xml, edits);
}

function collectSectionRanges(xml: string, partName: string): Map<string, SectionRange> {
  const ranges = new Map<string, SectionRange>();
  const stack: string[] = [];
  const pending = new Map<
    string,
    { start: number; openEnd: number; prefix: string; selfClosing: boolean }
  >();
  parseXml(
    xml,
    {
      onOpenTag(tag, position) {
        stack.push(tag.local);
        if (stack.length === 2 && isSectionName(tag.local)) {
          pending.set(tag.local, {
            start: openTagStart(xml, position),
            openEnd: position,
            prefix: tag.prefix,
            selfClosing: false,
          });
        }
      },
      onCloseTag(tag, start, end) {
        const active = pending.get(tag.local);
        if (active !== undefined && stack.length === 2) {
          ranges.set(tag.local, {
            local: tag.local,
            start: active.start,
            openEnd: active.openEnd,
            contentEnd: start,
            end,
            selfClosing: active.openEnd === end,
            prefix: active.prefix,
          });
          pending.delete(tag.local);
        }
        stack.pop();
      },
    },
    partName,
  );
  return ranges;
}

function findRootRange(
  xml: string,
  partName: string,
): { start: number; openEnd: number; end: number; prefix: string } {
  let root: { start: number; openEnd: number; end: number; prefix: string } | undefined = undefined;
  const stack: string[] = [];
  parseXml(
    xml,
    {
      onOpenTag(tag, position) {
        stack.push(tag.local);
        if (stack.length === 1 && tag.local === 'styleSheet') {
          root = {
            start: openTagStart(xml, position),
            openEnd: position,
            end: -1,
            prefix: tag.prefix,
          };
        }
      },
      onCloseTag(tag, _start, end) {
        if (stack.length === 1 && tag.local === 'styleSheet' && root !== undefined) {
          root = { ...root, end };
        }
        stack.pop();
      },
    },
    partName,
  );
  if (root === undefined) {
    throw new WorkbookSaveError('unsafe-part', `${partName} is not a style sheet.`);
  }
  return root;
}

function missingSectionInsertPosition(
  local: string,
  ranges: ReadonlyMap<string, SectionRange>,
  root: { start: number; openEnd: number; end: number },
  xml: string,
): number {
  const order = SECTION_ORDER.indexOf(local as (typeof SECTION_ORDER)[number]);
  for (let index = order + 1; index < SECTION_ORDER.length; index += 1) {
    const next = ranges.get(SECTION_ORDER[index]!);
    if (next !== undefined) return next.start;
  }
  for (let index = order - 1; index >= 0; index -= 1) {
    const previous = ranges.get(SECTION_ORDER[index]!);
    if (previous !== undefined) return previous.end;
  }
  const close = xml.lastIndexOf('</', root.end - 1);
  return close < 0 ? root.end : close;
}

function isSectionName(local: string): boolean {
  return (SECTION_ORDER as readonly string[]).includes(local);
}

function withCount(openText: string, count: number): string {
  if (/\bcount="[^"]*"/.test(openText)) {
    return openText.replace(/\bcount="[^"]*"/, `count="${count}"`);
  }
  return openText.replace(/>$/, ` count="${count}">`);
}

function qualify(prefix: string, local: string): string {
  return prefix === '' ? local : `${prefix}:${local}`;
}

function applyEdits(
  xml: string,
  edits: readonly { start: number; end: number; text: string; kind: 'replace' | 'insert' }[],
): string {
  const sorted = [...edits].sort((a, b) => {
    if (b.start !== a.start) return b.start - a.start;
    if (a.kind !== b.kind) return a.kind === 'replace' ? -1 : 1;
    return 0;
  });
  let output = xml;
  for (const edit of sorted) {
    output = output.slice(0, edit.start) + edit.text + output.slice(edit.end);
  }
  return output;
}

function fontSignature(font: FontStyle): string {
  return JSON.stringify([
    font.name,
    font.size,
    font.bold,
    font.italic,
    font.underline,
    font.underlineStyle,
    font.strike,
    font.color === null ? null : font.color.toUpperCase(),
  ]);
}

function fillSignature(fill: FillStyle): string {
  return JSON.stringify([
    fill.pattern,
    fill.color === null ? null : fill.color.toUpperCase(),
    fill.backgroundColor === null || fill.backgroundColor === undefined
      ? null
      : fill.backgroundColor.toUpperCase(),
  ]);
}

function borderSignature(border: BorderStyle): string {
  return JSON.stringify([
    edgeSignature(border.left),
    edgeSignature(border.right),
    edgeSignature(border.top),
    edgeSignature(border.bottom),
    edgeSignature(border.diagonal),
    border.diagonalUp,
    border.diagonalDown,
  ]);
}

function edgeSignature(edge: BorderEdgeStyle): string {
  const color = edge.color === null ? '' : edge.color.toUpperCase();
  return `${edge.style}\u0000${color}\u0000${edge.weight ?? ''}`;
}

function serializeFont(font: FontStyle): string {
  const parts: string[] = [];
  if (font.bold) parts.push('<b/>');
  if (font.italic) parts.push('<i/>');
  if (font.underline) {
    parts.push(font.underlineStyle === 'single' ? '<u/>' : `<u val="${font.underlineStyle}"/>`);
  }
  if (font.strike) parts.push('<strike/>');
  if (font.color !== null) parts.push(`<color rgb="${toArgb(font.color)}"/>`);
  parts.push(`<sz val="${formatNumber(font.size)}"/>`);
  parts.push(`<name val="${escapeXmlAttribute(font.name)}"/>`);
  return `<font>${parts.join('')}</font>`;
}

function serializeFill(fill: FillStyle): string {
  const children: string[] = [];
  if (fill.color !== null) children.push(`<fgColor rgb="${toArgb(fill.color)}"/>`);
  const background = fill.backgroundColor ?? null;
  if (background !== null) {
    children.push(`<bgColor rgb="${toArgb(background)}"/>`);
  }
  if (children.length === 0) return `<fill><patternFill patternType="${fill.pattern}"/></fill>`;
  return `<fill><patternFill patternType="${fill.pattern}">${children.join('')}</patternFill></fill>`;
}

function serializeBorder(border: BorderStyle): string {
  const attributes: string[] = [];
  if (border.diagonalUp) attributes.push('diagonalUp="1"');
  if (border.diagonalDown) attributes.push('diagonalDown="1"');
  const open = attributes.length === 0 ? '<border>' : `<border ${attributes.join(' ')}>`;
  return `${open}${serializeBorderEdge('left', border.left)}${serializeBorderEdge(
    'right',
    border.right,
  )}${serializeBorderEdge('top', border.top)}${serializeBorderEdge(
    'bottom',
    border.bottom,
  )}${serializeBorderEdge('diagonal', border.diagonal)}</border>`;
}

function serializeBorderEdge(name: string, edge: BorderEdgeStyle): string {
  if (edge.style === 'none') return `<${name}/>`;
  const color = edge.color === null ? '' : `<color rgb="${toArgb(edge.color)}"/>`;
  return `<${name} style="${escapeXmlAttribute(edge.style)}">${color}</${name}>`;
}

function serializeNumberFormat(id: number, code: string): string {
  return `<numFmt numFmtId="${id}" formatCode="${escapeXmlAttribute(code)}"/>`;
}

function serializeXf(
  style: StyleRecord,
  ids: { fontId: number; fillId: number; borderId: number; numFmtId: number },
): string {
  const attributes = [
    `numFmtId="${ids.numFmtId}"`,
    `fontId="${ids.fontId}"`,
    `fillId="${ids.fillId}"`,
    `borderId="${ids.borderId}"`,
    `xfId="${style.xfId}"`,
  ];
  if (ids.numFmtId !== 0) attributes.push('applyNumberFormat="1"');
  if (!sameFont(style.font, DEFAULT_FONT)) attributes.push('applyFont="1"');
  if (!sameFill(style.fill, DEFAULT_FILL)) attributes.push('applyFill="1"');
  if (!sameBorder(style.border, DEFAULT_BORDER)) attributes.push('applyBorder="1"');
  const alignment = serializeAlignment(style.alignment);
  if (alignment !== null) attributes.push('applyAlignment="1"');
  const open = `<xf ${attributes.join(' ')}`;
  return alignment === null ? `${open}/>` : `${open}>${alignment}</xf>`;
}

function serializeAlignment(alignment: AlignmentStyle): string | null {
  if (sameAlignment(alignment, DEFAULT_ALIGNMENT)) return null;
  const attributes: string[] = [];
  if (alignment.horizontal !== 'general') {
    attributes.push(`horizontal="${alignment.horizontal}"`);
  }
  if (alignment.vertical !== 'bottom') attributes.push(`vertical="${alignment.vertical}"`);
  if (alignment.wrapText) attributes.push('wrapText="1"');
  if (alignment.shrinkToFit) attributes.push('shrinkToFit="1"');
  if (alignment.indent !== 0) attributes.push(`indent="${alignment.indent}"`);
  if (alignment.textRotation !== 0) attributes.push(`textRotation="${alignment.textRotation}"`);
  return `<alignment ${attributes.join(' ')}/>`;
}

function sameFont(left: FontStyle, right: FontStyle): boolean {
  return fontSignature(left) === fontSignature(right);
}

function sameFill(left: FillStyle, right: FillStyle): boolean {
  return fillSignature(left) === fillSignature(right);
}

function sameBorder(left: BorderStyle, right: BorderStyle): boolean {
  return borderSignature(left) === borderSignature(right);
}

function sameAlignment(left: AlignmentStyle, right: AlignmentStyle): boolean {
  return (
    left.horizontal === right.horizontal &&
    left.vertical === right.vertical &&
    left.wrapText === right.wrapText &&
    left.indent === right.indent &&
    left.textRotation === right.textRotation &&
    left.shrinkToFit === right.shrinkToFit
  );
}

function toArgb(color: string): string {
  const hex = color.startsWith('#') ? color.slice(1) : color;
  if (/^[0-9a-f]{3}$/i.test(hex)) {
    const expanded = hex
      .split('')
      .map((character) => `${character}${character}`)
      .join('');
    return `FF${expanded.toUpperCase()}`;
  }
  if (/^[0-9a-f]{8}$/i.test(hex)) return hex.toUpperCase();
  if (/^[0-9a-f]{6}$/i.test(hex)) return `FF${hex.toUpperCase()}`;
  return hex.toUpperCase();
}

function formatNumber(value: number): string {
  return Number.isInteger(value) ? String(value) : String(value);
}
