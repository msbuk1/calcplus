import { SaxesParser } from 'saxes';
import type { SaxesAttributeNS, SaxesTagNS } from 'saxes';
import { invalidXmlCharacterMessage } from '@model/text';
import { WorkbookOpenError, WorkbookSaveError } from './errors';

export const SPREADSHEET_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
export const STRICT_SPREADSHEET_NS = 'http://purl.oclc.org/ooxml/spreadsheetml/main';
export const RELATIONSHIPS_NS =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export const PACKAGE_RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
export const CONTENT_TYPES_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
export const XML_DECLARATION_UTF8 = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
export const MAX_TREE_PATCH_BYTES = 16 * 1024 * 1024;

const utf8Decoder = new TextDecoder('utf-8');
const utf8Encoder = new TextEncoder();

export function decodeXml(bytes: Uint8Array): string {
  return utf8Decoder.decode(bytes);
}

export function encodeXml(text: string): Uint8Array {
  return utf8Encoder.encode(text);
}

export function assertXmlCharacters(value: string, context: string): void {
  const message = invalidXmlCharacterMessage(value, context);
  if (message !== null) {
    throw new WorkbookSaveError('unsupported-edit', message);
  }
}

export function escapeXmlText(value: string): string {
  assertXmlCharacters(value, 'Cell text');
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&apos;');
}

export function escapeXmlAttribute(value: string): string {
  return escapeXmlText(value);
}

export function assertSafeXml(xml: string, partName: string): void {
  if (/<!DOCTYPE/i.test(xml) || /<!ENTITY/i.test(xml)) {
    throw new WorkbookOpenError(
      'unsafe-xml',
      `The workbook contains a DTD or entity declaration in ${partName}; calcplus rejects it for safety.`,
    );
  }
}

export interface XmlEventHandlers {
  onOpenTag?(tag: SaxesTagNS, position: number): void;
  onCloseTag?(tag: SaxesTagNS, startPosition: number, endPosition: number): void;
  onText?(text: string): void;
  onCdata?(text: string): void;
  onComment?(text: string): void;
  onProcessingInstruction?(target: string, data: string): void;
}

/**
 * Parse a part with saxes. Positions are JavaScript string indexes; the
 * position reported for an open tag is the index just after its closing `>`.
 * DTDs and entity declarations are rejected before parsing.
 */
export function parseXml(xml: string, handlers: XmlEventHandlers, partName: string): void {
  assertSafeXml(xml, partName);
  const parser = new SaxesParser({ xmlns: true, position: true });
  parser.on('opentag', (tag) => {
    handlers.onOpenTag?.(tag, parser.position);
  });
  parser.on('closetag', (tag) => {
    const end = parser.position;
    const start = xml.lastIndexOf('<', Math.max(0, end - 1));
    handlers.onCloseTag?.(tag, start, end);
  });
  parser.on('text', (text) => handlers.onText?.(text));
  parser.on('cdata', (text) => handlers.onCdata?.(text));
  parser.on('comment', (text) => handlers.onComment?.(text));
  parser.on('processinginstruction', (instruction) =>
    handlers.onProcessingInstruction?.(instruction.target, instruction.body),
  );
  parser.on('error', (error) => {
    throw new WorkbookOpenError('corrupt', `Could not parse ${partName}: ${error.message}`);
  });
  try {
    parser.write(xml).close();
  } catch (error) {
    if (error instanceof WorkbookOpenError) throw error;
    const message = error instanceof Error ? error.message : String(error);
    throw new WorkbookOpenError('corrupt', `Could not parse ${partName}: ${message}`);
  }
}

export function isWellFormedXml(xml: string): boolean {
  try {
    parseXml(xml, {}, 'patched part');
    return true;
  } catch {
    return false;
  }
}

/** Validate a generated part before it can enter the package. */
export function validatePatchedPart(bytes: Uint8Array, partName: string): void {
  try {
    parseXml(decodeXml(bytes), {}, partName);
  } catch (error) {
    if (error instanceof WorkbookSaveError) throw error;
    const message = error instanceof Error ? error.message : String(error);
    throw new WorkbookSaveError(
      'unsafe-part',
      `The patched ${partName} is not valid XML (${message}). The file was not written.`,
    );
  }
}

export function tagLocalName(name: string): string {
  const separator = name.lastIndexOf(':');
  return separator < 0 ? name : name.slice(separator + 1);
}

export function attributeValue(tag: SaxesTagNS, localName: string): string | undefined {
  for (const attribute of Object.values(tag.attributes) as SaxesAttributeNS[]) {
    if (attribute.local === localName) return attribute.value;
  }
  return undefined;
}

export function attributeMap(tag: SaxesTagNS): Map<string, SaxesAttributeNS> {
  const map = new Map<string, SaxesAttributeNS>();
  for (const attribute of Object.values(tag.attributes) as SaxesAttributeNS[]) {
    map.set(attribute.name, attribute);
  }
  return map;
}

export function openTagStart(xml: string, position: number): number {
  return xml.lastIndexOf('<', Math.max(0, position - 1));
}

export function closeTagStart(xml: string, position: number): number {
  return xml.lastIndexOf('<', Math.max(0, position - 1));
}

export function resolvePartPath(basePartPath: string, target: string): string {
  if (target.startsWith('/')) return normalizePartName(target.slice(1));
  const baseParts = basePartPath.split('/');
  baseParts.pop();
  const targetParts = target.split('/');
  const combined = [...baseParts, ...targetParts];
  return normalizePartName(combined.join('/'));
}

function normalizePartName(path: string): string {
  const segments: string[] = [];
  for (const segment of path.split('/')) {
    if (segment === '' || segment === '.') continue;
    if (segment === '..') {
      if (segments.length === 0) {
        throw new WorkbookOpenError('corrupt', `Package relationship escapes the package: ${path}`);
      }
      segments.pop();
      continue;
    }
    try {
      segments.push(decodeURIComponent(segment));
    } catch {
      segments.push(segment);
    }
  }
  return segments.join('/');
}
