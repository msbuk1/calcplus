import type {
  GridRange,
  PageMarginsRecord,
  PrintOptionsRecord,
  PrintSetupRecord,
  Sheet,
} from '@model/types';
import { parseWholeArea } from '@model/references';
import { WorkbookSaveError } from './errors';
import { decodeXml, encodeXml, escapeXmlAttribute, isWellFormedXml, parseXml } from './xml';

/**
 * Print setup: `<printOptions>`, `<pageMargins>`, `<pageSetup>`, the
 * `<sheetPr><pageSetUpPr fitToPage>` switch and the `_xlnm.Print_Area` defined
 * name (file-compat §3, §9 F16).
 *
 * Reading maps margins, print options, orientation, fit-to-width, scale,
 * paper size and the fit-to-page switch onto `Sheet.printSetup`; the print area
 * is not duplicated in the model because it is a sheet-scoped defined name
 * (`workbook-model §1`) and `printAreaOf` reads it back. Every other `pageSetup`
 * attribute (`firstPageNumber`, `cellComments`, …) and every unmodelled child
 * stays verbatim: a save only rewrites the modelled attributes, in place, when
 * the model differs from the parsed source (the `autofilter-xml.ts` raw +
 * state-key pattern).
 *
 * The state key is per element, so a change to the orientation never rewrites
 * `<printOptions>` or `pageSetUpPr`, and an untouched sheet is byte-identical.
 * Honouring these settings in print preview and PDF export is M7 consumption;
 * this module only stores them faithfully.
 */

export const PRINT_AREA_NAME = '_xlnm.Print_Area';

/** Excel's defaults for the margins a declared element omits. */
const DEFAULT_MARGINS: PageMarginsRecord = {
  left: 0.7,
  right: 0.7,
  top: 0.75,
  bottom: 0.75,
  header: 0.3,
  footer: 0.3,
};

/** Children of `<worksheet>` that follow `<printOptions>`. */
const AFTER_PRINT_OPTIONS: ReadonlySet<string> = new Set([
  'pageMargins',
  'pageSetup',
  'headerFooter',
  'rowBreaks',
  'colBreaks',
  'customProperties',
  'cellWatches',
  'ignoredErrors',
  'smartTags',
  'drawing',
  'legacyDrawing',
  'legacyDrawingHF',
  'picture',
  'oleObjects',
  'controls',
  'webPublishItems',
  'tableParts',
  'extLst',
]);

/** Children of `<worksheet>` that follow `<pageSetup>` in the schema order. */
const AFTER_PAGE_SETUP: ReadonlySet<string> = new Set([
  'headerFooter',
  'rowBreaks',
  'colBreaks',
  'customProperties',
  'cellWatches',
  'ignoredErrors',
  'smartTags',
  'drawing',
  'legacyDrawing',
  'legacyDrawingHF',
  'picture',
  'oleObjects',
  'controls',
  'webPublishItems',
  'tableParts',
  'extLst',
]);

/** Children of `<worksheet>` that follow `<pageMargins>`. */
const AFTER_PAGE_MARGINS: ReadonlySet<string> = new Set(['pageSetup', ...AFTER_PAGE_SETUP]);

/**
 * The parsed source of the modelled elements, stored in
 * `PrintSetupRecord.raw`. The keys record the model values at parse time: while
 * an element's key still matches, that source element is never rewritten, which
 * keeps an untouched save byte-identical and a one-field edit to one value.
 */
export interface PrintSetupElementKeys {
  readonly margins: string;
  readonly options: string;
  readonly setup: string;
  readonly fitToPage: string;
}

export interface PrintSetupRawSource {
  readonly kind: 'print-setup-source';
  readonly stateKey: string;
  /** Per-element keys at parse time; absent on a restored older snapshot. */
  readonly elements?: PrintSetupElementKeys;
}

export function isPrintSetupRawSource(raw: unknown): raw is PrintSetupRawSource {
  if (raw === null || typeof raw !== 'object') return false;
  const candidate = raw as Partial<PrintSetupRawSource>;
  return candidate.kind === 'print-setup-source' && typeof candidate.stateKey === 'string';
}

export function isPrintAreaName(name: string): boolean {
  return name.toLowerCase() === PRINT_AREA_NAME.toLowerCase();
}

/**
 * The sheet's print areas, from its `_xlnm.Print_Area` name. A name may carry
 * several comma-separated areas; an area on another sheet is ignored.
 */
export function printAreaOf(sheet: Sheet): readonly GridRange[] {
  const name = sheet.names.find((candidate) => isPrintAreaName(candidate.name));
  if (name === undefined) return [];
  const ranges: GridRange[] = [];
  for (const areaText of splitAreas(name.formula)) {
    const area = parseWholeArea(areaText);
    if (area === null) continue;
    if (area.sheetName !== null && !sheetsNameEqual(area.sheetName, sheet.name)) continue;
    ranges.push({
      startRow: area.startRow,
      startCol: area.startCol,
      endRow: area.endRow,
      endCol: area.endCol,
    });
  }
  return ranges;
}

/** Parse the sheet's print-setup elements. `null` when none is declared. */
export function parsePrintSetup(xml: string, partName: string): PrintSetupRecord | null {
  const children = directChildren(xml, partName);
  const optionsChild = children.find((child) => child.local === 'printOptions') ?? null;
  const marginsChild = children.find((child) => child.local === 'pageMargins') ?? null;
  const setupChild = children.find((child) => child.local === 'pageSetup') ?? null;
  const sheetPrChild = children.find((child) => child.local === 'sheetPr') ?? null;
  const pageSetUpPr = sheetPrChild === null ? null : pageSetUpPrOf(xml, partName);
  if (
    optionsChild === null &&
    marginsChild === null &&
    setupChild === null &&
    pageSetUpPr === null
  ) {
    return null;
  }

  const options =
    optionsChild === null
      ? null
      : parsePrintOptions(openTagOf(xml, optionsChild.start, optionsChild.openEnd));
  const margins =
    marginsChild === null
      ? null
      : marginsOf(openTagOf(xml, marginsChild.start, marginsChild.openEnd));
  const setup =
    setupChild === null
      ? null
      : parseSetupAttributes(openTagOf(xml, setupChild.start, setupChild.openEnd));
  const record: PrintSetupRecord = {
    margins,
    options,
    orientation: setup?.orientation ?? null,
    fitToWidth: setup?.fitToWidth ?? null,
    fitToHeight: setup?.fitToHeight ?? null,
    scale: setup?.scale ?? null,
    fitToPage: pageSetUpPr === null ? null : parseFitToPage(pageSetUpPr.openTag),
    paperSize: setup?.paperSize ?? null,
    raw: { kind: 'print-setup-source', stateKey: '' },
  };
  const elements = printSetupElementKeys(record);
  return {
    ...record,
    raw: {
      kind: 'print-setup-source',
      stateKey: printSetupStateKey(record),
      elements,
    },
  };
}

/** Populate `Sheet.printSetup` from the worksheet part. */
export function applyPrintSetup(sheet: Sheet, xml: string, partName: string): void {
  sheet.printSetup = parsePrintSetup(xml, partName);
}

/**
 * A deterministic key for print setup's documented state. `raw.stateKey`
 * records the key at parse time; equality means the source elements are still
 * exactly what the model describes.
 */
export function printSetupStateKey(setup: PrintSetupRecord): string {
  const keys = printSetupElementKeys(setup);
  return [keys.margins, keys.options, keys.setup, keys.fitToPage].join('|');
}

/**
 * The state key of each modelled element. A change to one element rewrites only
 * that element's attributes (print-setup-xml header).
 */
export function printSetupElementKeys(setup: PrintSetupRecord): PrintSetupElementKeys {
  const fitToPage = setup.fitToPage ?? null;
  return {
    margins: marginsKey(setup.margins ?? null),
    options: optionsKey(setup.options ?? null),
    setup: setupKey(setup),
    fitToPage: fitToPage === null ? 'none' : fitToPage ? '1' : '0',
  };
}

function marginsKey(margins: PageMarginsRecord | null): string {
  if (margins === null) return 'none';
  return [margins.left, margins.right, margins.top, margins.bottom, margins.header, margins.footer]
    .map(plainNumber)
    .join(',');
}

function optionsKey(options: PrintOptionsRecord | null): string {
  if (options === null) return 'none';
  return [
    options.gridLines,
    options.gridLinesSet,
    options.headings,
    options.horizontalCentered,
    options.verticalCentered,
  ]
    .map((value) => (value ? '1' : '0'))
    .join(',');
}

function setupKey(setup: PrintSetupRecord): string {
  return [
    setup.orientation ?? 'none',
    setup.fitToWidth === null || setup.fitToWidth === undefined
      ? 'none'
      : plainNumber(setup.fitToWidth),
    setup.fitToHeight === null || setup.fitToHeight === undefined
      ? 'none'
      : plainNumber(setup.fitToHeight),
    setup.scale === null || setup.scale === undefined ? 'none' : plainNumber(setup.scale),
    setup.paperSize === null || setup.paperSize === undefined
      ? 'none'
      : plainNumber(setup.paperSize),
  ].join(',');
}

export interface PrintSetupSavePlan {
  readonly changed: boolean;
  readonly bytes: Uint8Array;
}

/**
 * True when the model's print setup no longer matches the source it was parsed
 * from. Cheap by design: an untouched sheet never reaches the XML parser.
 */
export function printSetupChanged(sheet: Sheet): boolean {
  const setup = sheet.printSetup;
  if (setup === undefined || setup === null) return false;
  return !isPrintSetupRawSource(setup.raw) || setup.raw.stateKey !== printSetupStateKey(setup);
}

/**
 * Write the model's print setup into the worksheet part. Attribute-level only:
 * an authored element keeps its prefix, attribute order and every attribute
 * the model does not own, so the diff of a changed orientation is one value.
 */
export function planPrintSetupSave(
  sheet: Sheet,
  sourceBytes: Uint8Array,
  partName: string,
): PrintSetupSavePlan {
  const setup = sheet.printSetup;
  if (setup === undefined || setup === null) {
    return { changed: false, bytes: sourceBytes };
  }
  if (isPrintSetupRawSource(setup.raw) && setup.raw.stateKey === printSetupStateKey(setup)) {
    return { changed: false, bytes: sourceBytes };
  }

  const xml = decodeXml(sourceBytes);
  const children = directChildren(xml, partName);
  const optionsChild = children.find((child) => child.local === 'printOptions') ?? null;
  const marginsChild = children.find((child) => child.local === 'pageMargins') ?? null;
  const setupChild = children.find((child) => child.local === 'pageSetup') ?? null;
  const changed = changedElements(setup);
  const edits: { start: number; end: number; text: string }[] = [];
  const options = setup.options ?? null;
  const fitToPageValue = setup.fitToPage ?? null;

  if (changed.options) {
    const attributes = optionsAttributes(options);
    if (options === null) {
      if (optionsChild !== null) {
        edits.push({ start: optionsChild.start, end: optionsChild.end, text: '' });
      }
    } else if (optionsChild !== null) {
      const openTag = openTagOf(xml, optionsChild.start, optionsChild.openEnd);
      const updated = updateOpenTagAttributes(openTag, attributes);
      if (updated !== openTag) {
        edits.push({ start: optionsChild.start, end: optionsChild.openEnd, text: updated });
      }
    } else {
      const prefix = rootPrefix(xml);
      const position = insertionPosition(children, AFTER_PRINT_OPTIONS, xml);
      edits.push({
        start: position,
        end: position,
        text: `<${qualified(prefix, 'printOptions')}${attributesText(attributes)}/>`,
      });
    }
  }

  if (changed.margins) {
    if (setup.margins === null) {
      if (marginsChild !== null) {
        edits.push({ start: marginsChild.start, end: marginsChild.end, text: '' });
      }
    } else {
      const attributes = new Map<string, string | null>([
        ['left', plainNumber(setup.margins.left)],
        ['right', plainNumber(setup.margins.right)],
        ['top', plainNumber(setup.margins.top)],
        ['bottom', plainNumber(setup.margins.bottom)],
        ['header', plainNumber(setup.margins.header)],
        ['footer', plainNumber(setup.margins.footer)],
      ]);
      if (marginsChild !== null) {
        const openTag = openTagOf(xml, marginsChild.start, marginsChild.openEnd);
        const updated = updateOpenTagAttributes(openTag, attributes);
        if (updated !== openTag) {
          edits.push({
            start: marginsChild.start,
            end: marginsChild.openEnd,
            text: updated,
          });
        }
      } else {
        const prefix = rootPrefix(xml);
        const position = insertionPosition(children, AFTER_PAGE_MARGINS, xml);
        edits.push({
          start: position,
          end: position,
          text: `<${qualified(prefix, 'pageMargins')}${attributesText(attributes)}/>`,
        });
      }
    }
  }

  if (changed.setup) {
    const setupAttributes = new Map<string, string | null>([
      ['orientation', setup.orientation ?? null],
      ['fitToWidth', setup.fitToWidth == null ? null : plainNumber(setup.fitToWidth)],
      ['fitToHeight', setup.fitToHeight == null ? null : plainNumber(setup.fitToHeight)],
      ['scale', setup.scale == null ? null : plainNumber(setup.scale)],
      ['paperSize', setup.paperSize == null ? null : plainNumber(setup.paperSize)],
    ]);
    if (setupChild !== null) {
      if (setupAttributesHaveValues(setupAttributes)) {
        const openTag = openTagOf(xml, setupChild.start, setupChild.openEnd);
        const updated = updateOpenTagAttributes(openTag, setupAttributes);
        if (updated !== openTag) {
          edits.push({
            start: setupChild.start,
            end: setupChild.openEnd,
            text: updated,
          });
        }
      }
    } else if (setupAttributesHaveValues(setupAttributes)) {
      const prefix = rootPrefix(xml);
      const position = insertionPosition(children, AFTER_PAGE_SETUP, xml);
      edits.push({
        start: position,
        end: position,
        text: `<${qualified(prefix, 'pageSetup')}${attributesText(setupAttributes)}/>`,
      });
    }
  }

  if (changed.fitToPage) {
    const fitToPage = fitToPageValue === null ? null : fitToPageValue ? '1' : '0';
    const sheetPrChild = children.find((child) => child.local === 'sheetPr') ?? null;
    const pageSetUpPr = sheetPrChild === null ? null : pageSetUpPrOf(xml, partName);
    if (pageSetUpPr !== null) {
      const updated = updateOpenTagAttributes(
        pageSetUpPr.openTag,
        new Map<string, string | null>([['fitToPage', fitToPage]]),
      );
      if (updated !== pageSetUpPr.openTag) {
        edits.push({
          start: pageSetUpPr.start,
          end: pageSetUpPr.openEnd,
          text: updated,
        });
      }
    } else if (fitToPage !== null && sheetPrChild !== null) {
      const prefix = sheetPrChild.prefix;
      const element = `<${qualified(prefix, 'pageSetUpPr')} fitToPage="${fitToPage}"/>`;
      if (isSelfClosing(xml, sheetPrChild)) {
        const expanded = `${xml.slice(sheetPrChild.start, sheetPrChild.openEnd - 2)}>${element}</${qualified(prefix, 'sheetPr')}>`;
        edits.push({ start: sheetPrChild.start, end: sheetPrChild.end, text: expanded });
      } else {
        const close = xml.lastIndexOf('</', sheetPrChild.end - 1);
        const position = close < 0 ? sheetPrChild.openEnd : close;
        edits.push({ start: position, end: position, text: element });
      }
    } else if (fitToPage !== null) {
      const prefix = rootPrefix(xml);
      const element = `<${qualified(prefix, 'sheetPr')}><${qualified(prefix, 'pageSetUpPr')} fitToPage="${fitToPage}"/></${qualified(prefix, 'sheetPr')}>`;
      const position =
        children.length > 0 ? children[0]!.start : insertionPosition(children, new Set(), xml);
      edits.push({ start: position, end: position, text: element });
    }
  }

  if (edits.length === 0) return { changed: false, bytes: sourceBytes };
  let patched = xml;
  for (const edit of [...edits].sort((left, right) => right.start - left.start)) {
    patched = patched.slice(0, edit.start) + edit.text + patched.slice(edit.end);
  }
  if (!isWellFormedXml(patched)) {
    throw new WorkbookSaveError(
      'unsafe-part',
      `The worksheet ${partName} could not store its print setup safely. The file was not written.`,
    );
  }
  return { changed: true, bytes: encodeXml(patched) };
}

interface ChangedElements {
  readonly margins: boolean;
  readonly options: boolean;
  readonly setup: boolean;
  readonly fitToPage: boolean;
}

/**
 * Which modelled elements differ from the parsed source. An older snapshot
 * without per-element keys cannot tell, so it rewrites all of them; the values
 * are still the model's, never lost.
 */
function changedElements(setup: PrintSetupRecord): ChangedElements {
  const current = printSetupElementKeys(setup);
  const raw = setup.raw;
  const before = isPrintSetupRawSource(raw) ? raw.elements : undefined;
  if (before === undefined) {
    return { margins: true, options: true, setup: true, fitToPage: true };
  }
  return {
    margins: before.margins !== current.margins,
    options: before.options !== current.options,
    setup: before.setup !== current.setup,
    fitToPage: before.fitToPage !== current.fitToPage,
  };
}

function optionsAttributes(options: PrintOptionsRecord | null): Map<string, string | null> {
  if (options === null) return new Map();
  const flag = (value: boolean): string => (value ? '1' : '0');
  return new Map<string, string | null>([
    ['gridLines', flag(options.gridLines)],
    ['gridLinesSet', flag(options.gridLinesSet)],
    ['headings', flag(options.headings)],
    ['horizontalCentered', flag(options.horizontalCentered)],
    ['verticalCentered', flag(options.verticalCentered)],
  ]);
}

function isSelfClosing(xml: string, child: DirectChild): boolean {
  return xml.slice(Math.max(0, child.openEnd - 2), child.openEnd) === '/>';
}

/**
 * The `<sheetPr><pageSetUpPr>` element. It is the only nested element print
 * reads, so it gets its own scan instead of a general tree walk.
 */
function pageSetUpPrOf(
  xml: string,
  partName: string,
): { readonly openTag: string; readonly start: number; readonly openEnd: number } | null {
  const holder: {
    found: { start: number; openEnd: number } | null;
  } = { found: null };
  const stack: string[] = [];
  parseXml(
    xml,
    {
      onOpenTag(tag, position) {
        if (
          holder.found === null &&
          tag.local === 'pageSetUpPr' &&
          stack.length === 2 &&
          stack[1] === 'sheetPr'
        ) {
          holder.found = {
            start: xml.lastIndexOf('<', Math.max(0, position - 1)),
            openEnd: position,
          };
        }
        stack.push(tag.local);
      },
      onCloseTag() {
        stack.pop();
      },
    },
    partName,
  );
  const found = holder.found;
  if (found === null) return null;
  return {
    openTag: xml.slice(found.start, found.openEnd),
    start: found.start,
    openEnd: found.openEnd,
  };
}

// ---------------------------------------------------------------------------
// Element scanning
// ---------------------------------------------------------------------------

interface DirectChild {
  readonly local: string;
  readonly prefix: string;
  readonly start: number;
  readonly openEnd: number;
  readonly end: number;
}

function directChildren(xml: string, partName: string): DirectChild[] {
  const children: DirectChild[] = [];
  const stack: { local: string; prefix: string; start: number; openEnd: number }[] = [];
  parseXml(
    xml,
    {
      onOpenTag(tag, position) {
        stack.push({
          local: tag.local,
          prefix: tag.prefix,
          start: xml.lastIndexOf('<', Math.max(0, position - 1)),
          openEnd: position,
        });
      },
      onCloseTag(_tag, _start, end) {
        const open = stack.pop();
        if (open === undefined) return;
        // The root is the only element left once a direct child closes.
        if (stack.length === 1) {
          children.push({
            local: open.local,
            prefix: open.prefix,
            start: open.start,
            openEnd: open.openEnd,
            end,
          });
        }
      },
    },
    partName,
  );
  return children;
}

function insertionPosition(
  children: readonly DirectChild[],
  after: ReadonlySet<string>,
  xml: string,
): number {
  const next = children.find((child) => after.has(child.local));
  if (next !== undefined) return next.start;
  const closing = xml.lastIndexOf('</');
  return closing < 0 ? xml.length : closing;
}

function rootPrefix(xml: string): string {
  const match = /<([A-Za-z_][\w.-]*:)?worksheet\b/.exec(xml);
  return match?.[1]?.slice(0, -1) ?? '';
}

function qualified(prefix: string, local: string): string {
  return prefix === '' ? local : `${prefix}:${local}`;
}

// ---------------------------------------------------------------------------
// Attribute parsing and rewriting
// ---------------------------------------------------------------------------

function openTagOf(xml: string, start: number, openEnd: number): string {
  return xml.slice(start, openEnd);
}

function marginsOf(openTag: string): PageMarginsRecord {
  const attributes = parseAttributes(openTag);
  const number = (local: keyof PageMarginsRecord): number => {
    const value = attributes.get(local);
    if (value === undefined) return DEFAULT_MARGINS[local];
    const parsed = Number(value);
    return Number.isFinite(parsed) ? parsed : DEFAULT_MARGINS[local];
  };
  return {
    left: number('left'),
    right: number('right'),
    top: number('top'),
    bottom: number('bottom'),
    header: number('header'),
    footer: number('footer'),
  };
}

interface SetupAttributes {
  readonly orientation: 'portrait' | 'landscape' | null;
  readonly fitToWidth: number | null;
  readonly fitToHeight: number | null;
  readonly scale: number | null;
  readonly paperSize: number | null;
}

function parseSetupAttributes(openTag: string): SetupAttributes {
  const attributes = parseAttributes(openTag);
  const orientationText = attributes.get('orientation');
  const integer = (local: string): number | null => {
    const value = attributes.get(local);
    if (value === undefined) return null;
    const parsed = Number(value);
    return Number.isInteger(parsed) ? parsed : null;
  };
  return {
    orientation:
      orientationText === 'portrait' || orientationText === 'landscape' ? orientationText : null,
    fitToWidth: integer('fitToWidth'),
    fitToHeight: integer('fitToHeight'),
    scale: integer('scale'),
    paperSize: integer('paperSize'),
  };
}

/** `<printOptions>`: an absent flag means `false`, as the schema defaults. */
function parsePrintOptions(openTag: string): PrintOptionsRecord {
  const attributes = parseAttributes(openTag);
  const flag = (local: string): boolean => xsdBoolean(attributes.get(local)) ?? false;
  return {
    gridLines: flag('gridLines'),
    gridLinesSet: flag('gridLinesSet'),
    headings: flag('headings'),
    horizontalCentered: flag('horizontalCentered'),
    verticalCentered: flag('verticalCentered'),
  };
}

function parseFitToPage(openTag: string): boolean | null {
  return xsdBoolean(parseAttributes(openTag).get('fitToPage'));
}

/** OOXML `xsd:boolean`: `0`/`1` and `false`/`true`; anything else is unread. */
function xsdBoolean(value: string | undefined): boolean | null {
  if (value === undefined) return null;
  const text = value.trim().toLowerCase();
  if (text === '1' || text === 'true') return true;
  if (text === '0' || text === 'false') return false;
  return null;
}

/**
 * Rewrite modelled attributes inside an open tag without touching any
 * unmodelled attribute, its order, the element prefix or the trailing slash.
 * A `null` value removes the attribute.
 */
function updateOpenTagAttributes(
  openTag: string,
  updates: ReadonlyMap<string, string | null>,
): string {
  const pattern = /\s+([^\s=/>]+)\s*=\s*("([^"]*)"|'([^']*)')/g;
  const applied = new Set<string>();
  let output = '';
  let lastIndex = 0;
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(openTag)) !== null) {
    const name = match[1]!;
    const local = name.includes(':') ? name.slice(name.lastIndexOf(':') + 1) : name;
    if (!updates.has(local)) continue;
    output += openTag.slice(lastIndex, match.index);
    const value = updates.get(local);
    if (value !== null && value !== undefined) {
      output += ` ${name}="${escapeXmlAttribute(value)}"`;
    }
    lastIndex = pattern.lastIndex;
    applied.add(local);
  }
  output += openTag.slice(lastIndex);
  const missing = [...updates.entries()].filter(
    ([name, value]) => value !== null && !applied.has(name),
  );
  if (missing.length === 0) return output;
  const close = output.lastIndexOf('>');
  if (close < 0) return output;
  const insertAt = output[close - 1] === '/' ? close - 1 : close;
  return `${output.slice(0, insertAt)}${attributesText(new Map(missing))}${output.slice(insertAt)}`;
}

function attributesText(attributes: ReadonlyMap<string, string | null>): string {
  return [...attributes.entries()]
    .filter(([, value]) => value !== null && value !== undefined)
    .map(([name, value]) => ` ${name}="${escapeXmlAttribute(value!)}"`)
    .join('');
}

function setupAttributesHaveValues(attributes: ReadonlyMap<string, string | null>): boolean {
  for (const value of attributes.values()) {
    if (value !== null && value !== undefined) return true;
  }
  return false;
}

function parseAttributes(openTag: string): Map<string, string> {
  const attributes = new Map<string, string>();
  const pattern = /([A-Za-z_][\w.:-]*)\s*=\s*("([^"]*)"|'([^']*)')/g;
  for (const match of openTag.matchAll(pattern)) {
    const name = match[1] ?? '';
    const local = name.includes(':') ? name.slice(name.lastIndexOf(':') + 1) : name;
    attributes.set(local, decodeEntities(match[3] ?? match[4] ?? ''));
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

// ---------------------------------------------------------------------------
// Ranges and numbers
// ---------------------------------------------------------------------------

function splitAreas(formula: string): string[] {
  const areas: string[] = [];
  let current = '';
  let quoted = false;
  for (let index = 0; index < formula.length; index += 1) {
    const character = formula[index] ?? '';
    if (character === "'") {
      if (quoted && formula[index + 1] === "'") {
        current += "''";
        index += 1;
        continue;
      }
      quoted = !quoted;
      current += character;
      continue;
    }
    if (character === ',' && !quoted) {
      if (current.trim() !== '') areas.push(current.trim());
      current = '';
      continue;
    }
    current += character;
  }
  if (current.trim() !== '') areas.push(current.trim());
  return areas;
}

function sheetsNameEqual(left: string, right: string): boolean {
  return left.toUpperCase() === right.toUpperCase();
}

function plainNumber(value: number): string {
  return Object.is(value, -0) ? '0' : String(value);
}
