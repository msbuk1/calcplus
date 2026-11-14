import type { SharedStringEntry, SharedStringsTable } from './types';
import { attributeValue, decodeXml, openTagStart, parseXml } from './xml';

interface SharedStringsStructure {
  readonly xml: string;
  readonly entries: SharedStringEntry[];
  readonly count: number;
  readonly uniqueCount: number;
  readonly rootStart: number;
  readonly rootOpenEnd: number;
  readonly rootCloseStart: number;
  readonly rootPrefix: string;
  readonly selfClosing: boolean;
}

interface OpenElement {
  readonly local: string;
  readonly start: number;
  readonly openEnd: number;
  readonly prefix: string;
}

export function parseSharedStringsXml(
  bytes: Uint8Array,
  partName: string,
): { table: SharedStringsTable; structure: SharedStringsStructure } {
  const xml = decodeXml(bytes);
  const entries: SharedStringEntry[] = [];
  const stack: OpenElement[] = [];
  const rootHolder: { root: OpenElement | null } = { root: null };
  let rootCloseStart = 0;
  let rootSelfClosing = false;
  let count = 0;
  let uniqueCount = 0;
  let inSi = false;
  let siStart = 0;
  let siText = '';
  let siRich = false;
  let inText = false;
  let inPhonetic = false;

  parseXml(
    xml,
    {
      onOpenTag(tag, position) {
        const start = openTagStart(xml, position);
        stack.push({ local: tag.local, start, openEnd: position, prefix: tag.prefix });
        if (tag.local === 'sst' && rootHolder.root === null) {
          rootHolder.root = { local: tag.local, start, openEnd: position, prefix: tag.prefix };
          const countText = attributeValue(tag, 'count');
          const uniqueText = attributeValue(tag, 'uniqueCount');
          count = countText === undefined ? 0 : Number(countText) || 0;
          uniqueCount = uniqueText === undefined ? entries.length : Number(uniqueText) || 0;
          return;
        }
        if (tag.local === 'si') {
          inSi = true;
          siStart = start;
          siText = '';
          siRich = false;
          inPhonetic = false;
        } else if (inSi) {
          if (tag.local === 'r') siRich = true;
          if (tag.local === 'rPh') inPhonetic = true;
          if (tag.local === 't') inText = true;
        }
      },
      onCloseTag(tag, start, end) {
        const open = stack.pop();
        if (tag.local === 'sst' && open !== undefined) {
          rootCloseStart = start;
          rootSelfClosing = open.openEnd === end;
        }
        if (tag.local === 't') inText = false;
        if (tag.local === 'rPh') inPhonetic = false;
        if (tag.local === 'si' && open !== undefined) {
          entries.push({
            text: siText,
            richText: siRich,
            rawXml: xml.slice(siStart, end),
          });
          inSi = false;
        }
      },
      onText(text) {
        if (inSi && inText && !inPhonetic) siText += text;
      },
      onCdata(text) {
        if (inSi && inText && !inPhonetic) siText += text;
      },
    },
    partName,
  );

  if (rootHolder.root === null) {
    throw new Error(`${partName} has no sst root`);
  }
  const root = rootHolder.root;
  const table: SharedStringsTable = {
    partPath: partName,
    entries,
    count: Math.max(count, entries.length),
    uniqueCount: Math.max(uniqueCount, entries.length),
  };
  return {
    table,
    structure: {
      xml,
      entries,
      count: table.count,
      uniqueCount: table.uniqueCount,
      rootStart: root.start,
      rootOpenEnd: root.openEnd,
      rootCloseStart,
      rootPrefix: root.prefix,
      selfClosing: rootSelfClosing,
    },
  };
}

export type SharedStringsStructureInternal = SharedStringsStructure;
