import type { Sheet } from '@model/types';
import { columnToLetters } from '@model/types';
import { decodeXml, encodeXml, openTagStart, parseXml } from './xml';
import type { SaxesTagNS } from 'saxes';

/**
 * The sheet's frozen-pane split in the worksheet part (file-compat §3,
 * workbook-model §1 "sheet view settings that round-trip").
 *
 * Reading the pane is already done by `worksheet-xml.ts`. This module owns the
 * write side: it compares the model's split with the source pane and rewrites
 * only the `<pane>` child of the first `<sheetView>` when it differs. An
 * untouched split leaves the part byte-identical. A file with no sheetViews
 * receives one at the schema position before `sheetFormatPr`, `cols` or
 * `sheetData`.
 */

export interface PaneSavePlan {
  readonly changed: boolean;
  readonly bytes: Uint8Array;
}

interface PaneScan {
  readonly rootPrefix: string;
  readonly sheetViewsPrefix: string;
  readonly sheetViewsStart: number;
  readonly sheetViewsOpenEnd: number;
  readonly sheetViewsSelfClosing: boolean;
  readonly sheetViewPrefix: string;
  readonly sheetViewStart: number;
  readonly sheetViewOpenEnd: number;
  readonly sheetViewSelfClosing: boolean;
  readonly paneStart: number;
  readonly paneEnd: number;
  readonly frozenRows: number;
  readonly frozenCols: number;
}

export function planPaneSave(sheet: Sheet, sourceBytes: Uint8Array | null): PaneSavePlan {
  const rows = normalizeSplit(sheet.view.freezeRows);
  const cols = normalizeSplit(sheet.view.freezeCols);
  if (sourceBytes === null) {
    // A new sheet is written whole by `new-worksheet.ts`; there is no existing
    // part for this plan to patch.
    return { changed: false, bytes: new Uint8Array() };
  }
  const xml = decodeXml(sourceBytes);
  const scan = scanPane(xml, sheet.partPath);
  if (scan !== null && scan.frozenRows === rows && scan.frozenCols === cols) {
    return { changed: false, bytes: sourceBytes };
  }
  const next = writePane(xml, rows, cols, scan);
  return { changed: true, bytes: encodeXml(next) };
}

function normalizeSplit(value: number): number {
  if (!Number.isFinite(value) || value <= 0) return 0;
  return Math.floor(value);
}

/** Locate the first sheetView's pane and read its frozen split. */
function scanPane(xml: string, partName: string): PaneScan | null {
  let rootPrefix = '';
  let sheetViewsPrefix = '';
  let sheetViewsStart = -1;
  let sheetViewsOpenEnd = -1;
  let sheetViewsSelfClosing = false;
  let sheetViewPrefix = '';
  let sheetViewStart = -1;
  let sheetViewOpenEnd = -1;
  let sheetViewSelfClosing = false;
  let paneStart = -1;
  let paneEnd = -1;
  let frozenRows = 0;
  let frozenCols = 0;
  let depth = 0;
  let sheetViewsDepth = -1;
  let sheetViewDepth = -1;

  parseXml(
    xml,
    {
      onOpenTag(tag, position) {
        const local = tag.local;
        const start = openTagStart(xml, position);
        if (local === 'worksheet' && depth === 0) rootPrefix = tagPrefix(tag.name);
        if (local === 'sheetViews' && sheetViewsDepth < 0) {
          sheetViewsPrefix = tagPrefix(tag.name);
          sheetViewsStart = start;
          sheetViewsOpenEnd = position;
          sheetViewsSelfClosing = xml.slice(start, position).endsWith('/>');
          sheetViewsDepth = depth;
        } else if (
          local === 'sheetView' &&
          sheetViewsDepth >= 0 &&
          sheetViewStart < 0 &&
          depth === sheetViewsDepth + 1
        ) {
          sheetViewPrefix = tagPrefix(tag.name);
          sheetViewStart = start;
          sheetViewOpenEnd = position;
          sheetViewSelfClosing = xml.slice(start, position).endsWith('/>');
          sheetViewDepth = depth;
        } else if (
          local === 'pane' &&
          sheetViewStart >= 0 &&
          paneStart < 0 &&
          depth === sheetViewDepth + 1
        ) {
          paneStart = start;
          paneEnd = -1;
          const state = attribute(tag, 'state');
          if (state === 'frozen' || state === 'frozenSplit') {
            frozenCols = splitAttribute(tag, 'xSplit');
            frozenRows = splitAttribute(tag, 'ySplit');
          }
        }
        depth += 1;
      },
      onCloseTag(tag, start, end) {
        depth -= 1;
        if (tag.local === 'pane' && paneStart >= 0 && paneEnd < 0) {
          paneEnd = end;
        } else if (tag.local === 'sheetView' && sheetViewStart >= 0 && sheetViewDepth === depth) {
          sheetViewSelfClosing = sheetViewStart === start;
        } else if (
          tag.local === 'sheetViews' &&
          sheetViewsStart >= 0 &&
          sheetViewsDepth === depth
        ) {
          sheetViewsSelfClosing = sheetViewsStart === start;
        }
      },
    },
    partName,
  );

  if (sheetViewsStart < 0) {
    return {
      rootPrefix,
      sheetViewsPrefix,
      sheetViewsStart,
      sheetViewsOpenEnd,
      sheetViewsSelfClosing,
      sheetViewPrefix,
      sheetViewStart,
      sheetViewOpenEnd,
      sheetViewSelfClosing,
      paneStart,
      paneEnd,
      frozenRows: 0,
      frozenCols: 0,
    };
  }
  return {
    rootPrefix,
    sheetViewsPrefix,
    sheetViewsStart,
    sheetViewsOpenEnd,
    sheetViewsSelfClosing,
    sheetViewPrefix,
    sheetViewStart,
    sheetViewOpenEnd,
    sheetViewSelfClosing,
    paneStart,
    paneEnd,
    frozenRows,
    frozenCols,
  };
}

function writePane(xml: string, rows: number, cols: number, scan: PaneScan | null): string {
  const hasSplit = rows > 0 || cols > 0;
  const prefix =
    scan === null ? '' : scan.sheetViewPrefix || scan.sheetViewsPrefix || scan.rootPrefix;
  const pane = hasSplit ? paneElement(rows, cols, prefix) : '';

  if (scan === null || scan.sheetViewsStart < 0) {
    if (!hasSplit) return xml;
    const rootPrefix = scan?.rootPrefix ?? rootPrefixOf(xml);
    return insertSheetViews(xml, pane, rootPrefix);
  }

  if (scan.sheetViewsSelfClosing) {
    if (!hasSplit) return xml;
    const replacement = sheetViewsBlock(pane, scan.sheetViewsPrefix || scan.rootPrefix);
    return xml.slice(0, scan.sheetViewsStart) + replacement + xml.slice(scan.sheetViewsOpenEnd);
  }

  if (scan.sheetViewStart < 0) {
    if (!hasSplit) return xml;
    const view = sheetViewElement(pane, scan.sheetViewsPrefix || scan.rootPrefix);
    return xml.slice(0, scan.sheetViewsOpenEnd) + view + xml.slice(scan.sheetViewsOpenEnd);
  }

  if (scan.paneStart >= 0) {
    return xml.slice(0, scan.paneStart) + pane + xml.slice(scan.paneEnd);
  }
  if (!hasSplit) return xml;
  if (scan.sheetViewSelfClosing) {
    const open = xml.slice(scan.sheetViewStart, scan.sheetViewOpenEnd).replace(/\/\s*>$/, '>');
    return (
      xml.slice(0, scan.sheetViewStart) +
      open +
      pane +
      `</${qualified('sheetView', prefix)}>` +
      xml.slice(scan.sheetViewOpenEnd)
    );
  }
  return xml.slice(0, scan.sheetViewOpenEnd) + pane + xml.slice(scan.sheetViewOpenEnd);
}

function insertSheetViews(xml: string, pane: string, rootPrefix: string): string {
  const block = sheetViewsBlock(pane, rootPrefix);
  const next = /<(?:[A-Za-z_][\w.-]*:)?(?:sheetFormatPr|cols|sheetData)\b/.exec(xml);
  const position = next === null ? xml.lastIndexOf('</') : next.index;
  if (position < 0) return xml;
  return xml.slice(0, position) + block + xml.slice(position);
}

function sheetViewsBlock(pane: string, prefix: string): string {
  return `<${qualified('sheetViews', prefix)}>${sheetViewElement(pane, prefix)}</${qualified('sheetViews', prefix)}>`;
}

function sheetViewElement(pane: string, prefix: string): string {
  const open = `<${qualified('sheetView', prefix)} workbookViewId="0"`;
  if (pane === '') return `${open}/>`;
  return `${open}>${pane}</${qualified('sheetView', prefix)}>`;
}

function paneElement(rows: number, cols: number, prefix: string): string {
  const topLeftCell = `${columnToLetters(cols)}${rows + 1}`;
  const xSplit = cols > 0 ? ` xSplit="${cols}"` : '';
  const ySplit = rows > 0 ? ` ySplit="${rows}"` : '';
  const activePane = rows > 0 && cols > 0 ? 'bottomRight' : rows > 0 ? 'bottomLeft' : 'topRight';
  return `<${qualified('pane', prefix)}${xSplit}${ySplit} topLeftCell="${topLeftCell}" activePane="${activePane}" state="frozen"/>`;
}

function qualified(local: string, prefix: string): string {
  return prefix === '' ? local : `${prefix}:${local}`;
}

function tagPrefix(name: string): string {
  const separator = name.lastIndexOf(':');
  return separator < 0 ? '' : name.slice(0, separator);
}

function rootPrefixOf(xml: string): string {
  const match = /<([A-Za-z_][\w.-]*:)?worksheet\b/.exec(xml);
  if (match === null || match[1] === undefined) return '';
  return match[1].slice(0, -1);
}

function attribute(tag: SaxesTagNS, local: string): string | undefined {
  for (const candidate of Object.values(tag.attributes)) {
    if (candidate.local === local) return candidate.value;
  }
  return undefined;
}

function splitAttribute(tag: SaxesTagNS, local: string): number {
  const value = Number(attribute(tag, local) ?? '0');
  return Number.isFinite(value) && value > 0 ? Math.floor(value) : 0;
}
