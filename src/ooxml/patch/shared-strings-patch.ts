import type { SharedStringEntry, SharedStringsTable } from '../types';
import type { SharedStringsStructureInternal } from '../shared-strings';
import { escapeXmlText } from '../xml';

export interface SharedStringsAppender {
  indexFor(text: string): number;
  readonly addedCount: number;
  readonly changed: boolean;
  bytes(): Uint8Array;
  table(): SharedStringsTable;
}

export function createSharedStringsAppender(
  structure: SharedStringsStructureInternal,
  partPath: string,
): SharedStringsAppender {
  const entries: SharedStringEntry[] = structure.entries.map((entry) => ({ ...entry }));
  const plainIndex = new Map<string, number>();
  entries.forEach((entry, index) => {
    if (!entry.richText && !plainIndex.has(entry.text)) plainIndex.set(entry.text, index);
  });
  let added = 0;
  const prefix = structure.rootPrefix;
  const tag = (local: string): string => (prefix === '' ? local : `${prefix}:${local}`);

  return {
    indexFor(text: string): number {
      const existing = plainIndex.get(text);
      if (existing !== undefined) return existing;
      const index = entries.length;
      const rawXml = `<${tag('si')}><${tag('t')} xml:space="preserve">${escapeXmlText(
        text,
      )}</${tag('t')}></${tag('si')}>`;
      entries.push({ text, richText: false, rawXml });
      plainIndex.set(text, index);
      added += 1;
      return index;
    },
    get addedCount(): number {
      return added;
    },
    get changed(): boolean {
      return added > 0;
    },
    bytes(): Uint8Array {
      if (added === 0) {
        return new TextEncoder().encode(structure.xml);
      }
      const xml = structure.xml;
      const newCount = structure.count + added;
      const newUnique = structure.uniqueCount + added;
      const openTag = updateCounts(
        structure.xml.slice(structure.rootStart, structure.rootOpenEnd),
        newCount,
        newUnique,
      );
      if (structure.selfClosing) {
        const nonSelfClosing = openTag.replace(/\/>$/, '>');
        const closing = `</${tag('sst')}>`;
        const appended = entries
          .slice(structure.entries.length)
          .map((entry) => entry.rawXml)
          .join('');
        return new TextEncoder().encode(
          `${xml.slice(0, structure.rootStart)}${nonSelfClosing}${appended}${closing}${xml.slice(
            structure.rootOpenEnd,
          )}`,
        );
      }
      const appended = entries
        .slice(structure.entries.length)
        .map((entry) => entry.rawXml)
        .join('');
      return new TextEncoder().encode(
        `${xml.slice(0, structure.rootStart)}${openTag}${xml.slice(
          structure.rootOpenEnd,
          structure.rootCloseStart,
        )}${appended}${xml.slice(structure.rootCloseStart)}`,
      );
    },
    table(): SharedStringsTable {
      return {
        partPath,
        entries,
        count: structure.count + added,
        uniqueCount: structure.uniqueCount + added,
      };
    },
  };
}

function updateCounts(openTag: string, count: number, uniqueCount: number): string {
  const selfClosing = /\/>\s*$/.test(openTag);
  let content = selfClosing ? openTag.replace(/\/>\s*$/, '') : openTag.replace(/>\s*$/, '');
  let foundCount = false;
  let foundUnique = false;
  content = content.replace(
    /(\s)([\w:.-]+)\s*=\s*("([^"]*)"|'([^']*)')/g,
    (match, space: string, name: string) => {
      if (name === 'count') {
        foundCount = true;
        return `${space}count="${count}"`;
      }
      if (name === 'uniqueCount') {
        foundUnique = true;
        return `${space}uniqueCount="${uniqueCount}"`;
      }
      return match;
    },
  );
  if (!foundCount) content += ` count="${count}"`;
  if (!foundUnique) content += ` uniqueCount="${uniqueCount}"`;
  return `${content}${selfClosing ? '/>' : '>'}`;
}
