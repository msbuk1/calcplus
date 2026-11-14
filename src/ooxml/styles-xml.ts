import type { NamedStyleRecord, StyleRecord } from '@model/types';
import { createDefaultStyle } from '@model/workbook';
import { decodeXml, parseXml, openTagStart } from './xml';

export interface StylesXmlInfo {
  readonly styles: StyleRecord[];
  readonly namedStyles: NamedStyleRecord[];
}

interface OpenElement {
  readonly local: string;
  readonly start: number;
}

export function parseStylesXml(bytes: Uint8Array, partName: string): StylesXmlInfo {
  const xml = decodeXml(bytes);
  const styles: StyleRecord[] = [];
  const namedStyles: NamedStyleRecord[] = [];
  const stack: OpenElement[] = [];
  let inCellXfs = false;
  let inCellStyleXfs = false;
  let cellXfsDepth = -1;
  let cellStyleXfsDepth = -1;

  parseXml(
    xml,
    {
      onOpenTag(tag, position) {
        const start = openTagStart(xml, position);
        stack.push({ local: tag.local, start });
        const depth = stack.length;
        if (tag.local === 'cellXfs') {
          inCellXfs = true;
          cellXfsDepth = depth;
        } else if (tag.local === 'cellStyleXfs') {
          inCellStyleXfs = true;
          cellStyleXfsDepth = depth;
        } else if (inCellXfs && tag.local === 'xf') {
          styles.push({
            id: styles.length,
            numFmtId: numberAttribute(tag, 'numFmtId', 0),
            fontId: numberAttribute(tag, 'fontId', 0),
            fillId: numberAttribute(tag, 'fillId', 0),
            borderId: numberAttribute(tag, 'borderId', 0),
            xfId: numberAttribute(tag, 'xfId', 0),
            rawXml: '',
          });
        } else if (inCellStyleXfs && tag.local === 'xf') {
          namedStyles.push({
            name: '',
            xfId: namedStyles.length,
            builtinId: null,
          });
        }
      },
      onCloseTag(tag, _start, end) {
        const depth = stack.length;
        const open = stack.pop();
        if (inCellXfs && tag.local === 'xf' && open !== undefined) {
          const index = styles.length - 1;
          const last = styles[index];
          if (last !== undefined) {
            styles[index] = { ...last, rawXml: xml.slice(open.start, end) };
          }
        }
        if (tag.local === 'cellXfs' && depth === cellXfsDepth) {
          inCellXfs = false;
          cellXfsDepth = -1;
        }
        if (tag.local === 'cellStyleXfs' && depth === cellStyleXfsDepth) {
          inCellStyleXfs = false;
          cellStyleXfsDepth = -1;
        }
      },
    },
    partName,
  );

  if (styles.length === 0) styles.push(createDefaultStyle(0));
  return {
    styles: styles.map((style, index) => ({ ...style, id: index })),
    namedStyles,
  };
}

function numberAttribute(
  tag: { attributes: Record<string, { local: string; value: string }> },
  local: string,
  fallback: number,
): number {
  for (const attribute of Object.values(tag.attributes)) {
    if (attribute.local === local) {
      const value = Number(attribute.value);
      return Number.isFinite(value) ? value : fallback;
    }
  }
  return fallback;
}
