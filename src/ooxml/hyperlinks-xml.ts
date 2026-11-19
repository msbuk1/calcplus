import type { HyperlinkRecord, Sheet } from '@model/types';
import { cellKey, parseRangeRef } from '@model/types';
import { findRelationshipById } from './relationships';
import type { Relationship } from './types';

/**
 * The `hyperlinks` element of a worksheet part (file-compat §3, §9 F13).
 *
 * v1 models a hyperlink per cell and never authors one (feature matrix
 * "Hyperlinks: read-only; insert/edit dialog is V2"). An external target is the
 * relationship's target URL; an internal jump keeps its `location` text, which
 * is how Excel stores `#Sheet2!A1`. The element itself is preserved verbatim:
 * a save only rewrites it when a structural edit moves its `ref` or transforms
 * its internal `location`, which the source-part renumber does (see
 * `transformHyperlinkLocation` in `@model/range-transform`).
 *
 * The `location` grammar is `[Sheet!]A1[:B2]`, with an optional quoted sheet
 * name and optional `$` markers; a bare cell is relative to the sheet that
 * holds the hyperlink.
 *
 * Following a link is a surface decision outside this milestone; this module
 * only makes the target available to the model.
 */

export interface ParsedHyperlinkElement {
  /** The `ref` text as authored. */
  readonly ref: string;
  /** The external URL or part path, or `''` for a same-workbook location. */
  readonly target: string;
  readonly location?: string;
  readonly tooltip?: string;
}

export interface HyperlinkParseInput {
  readonly relationships: readonly Relationship[];
}

/** Parse every `hyperlink` element of an already-sliced `<hyperlinks>` element. */
export function parseHyperlinkElements(
  elementXml: string,
  input: HyperlinkParseInput,
): ParsedHyperlinkElement[] {
  const links: ParsedHyperlinkElement[] = [];
  for (const element of scanElements(elementXml, 'hyperlink')) {
    const openTag = openTagOf(element.text);
    if (openTag === null) continue;
    const attributes = parseAttributes(openTag.text);
    const ref = attributes.get('ref');
    if (ref === undefined || ref === '') continue;
    const relationshipId = attributes.get('id');
    const relationship =
      relationshipId === undefined
        ? undefined
        : findRelationshipById(input.relationships, relationshipId);
    const location = attributes.get('location');
    const tooltip = attributes.get('tooltip');
    const target =
      relationship === undefined
        ? ''
        : relationship.targetMode === 'External'
          ? relationship.target
          : relationship.resolvedTarget;
    links.push({
      ref,
      target,
      ...(location === undefined || location === '' ? {} : { location }),
      ...(tooltip === undefined || tooltip === '' ? {} : { tooltip }),
    });
  }
  return links;
}

/**
 * Populate `Sheet.hyperlinks` from the worksheet part. A range `ref` anchors one
 * record per cell, because the model's key is a cell (workbook-model §3).
 */
export function applyHyperlinks(
  sheet: Sheet,
  xml: string,
  rootChildren: readonly { readonly local: string; readonly start: number; readonly end: number }[],
  input: HyperlinkParseInput,
): void {
  const hyperlinks = new Map<string, HyperlinkRecord>();
  for (const child of rootChildren) {
    if (child.local !== 'hyperlinks') continue;
    for (const link of parseHyperlinkElements(xml.slice(child.start, child.end), input)) {
      const record: HyperlinkRecord = {
        target: link.target,
        ...(link.location === undefined ? {} : { location: link.location }),
        ...(link.tooltip === undefined ? {} : { tooltip: link.tooltip }),
      };
      for (const key of keysOf(link.ref)) hyperlinks.set(key, record);
    }
  }
  sheet.hyperlinks.clear();
  for (const [key, record] of hyperlinks) sheet.hyperlinks.set(key, record);
}

function keysOf(ref: string): string[] {
  try {
    const range = parseRangeRef(ref);
    const keys: string[] = [];
    for (let row = range.startRow; row <= range.endRow; row += 1) {
      for (let col = range.startCol; col <= range.endCol; col += 1) keys.push(cellKey(row, col));
    }
    return keys;
  } catch {
    return [];
  }
}

// ---------------------------------------------------------------------------
// Text helpers
// ---------------------------------------------------------------------------

interface OpenTag {
  readonly text: string;
}

interface ScannedElement {
  readonly text: string;
}

function openTagOf(xml: string): OpenTag | null {
  const start = xml.indexOf('<');
  if (start < 0) return null;
  const end = findTagEnd(xml, start);
  if (end < 0) return null;
  return { text: xml.slice(start, end) };
}

function scanElements(xml: string, local: string): ScannedElement[] {
  const elements: ScannedElement[] = [];
  const pattern = new RegExp(`<(?:[\\w.-]+:)?${local}\\b`, 'g');
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(xml)) !== null) {
    const start = match.index;
    const openEnd = findTagEnd(xml, start);
    if (openEnd < 0) break;
    const openTag = xml.slice(start, openEnd);
    if (/\/\s*>$/.test(openTag)) {
      elements.push({ text: openTag });
      pattern.lastIndex = openEnd;
      continue;
    }
    const qualified = /^<([^\s/>]+)/.exec(openTag)?.[1] ?? local;
    const close = `</${qualified}>`;
    const closeIndex = xml.indexOf(close, openEnd);
    if (closeIndex < 0) break;
    const end = closeIndex + close.length;
    elements.push({ text: xml.slice(start, end) });
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
