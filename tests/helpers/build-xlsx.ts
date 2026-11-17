import { SaxesParser } from 'saxes';
import { unzipSync, zipSync } from 'fflate';

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const SPREADSHEET_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PACKAGE_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

export interface BuildXlsxOptions {
  readonly sheetName?: string;
  readonly sharedStrings?: string[];
  readonly stylesXml?: string;
  readonly extraParts?: Record<string, string>;
  readonly strict?: boolean;
  readonly workbookExtras?: string;
  readonly workbookRelsExtras?: string;
  readonly contentTypeOverrides?: readonly (readonly [string, string])[];
}

export function buildXlsx(worksheetXml: string, options: BuildXlsxOptions = {}): Uint8Array {
  const sheetName = options.sheetName ?? 'Sheet1';
  const ns = options.strict ? 'http://purl.oclc.org/ooxml/spreadsheetml/main' : SPREADSHEET_NS;
  const sharedStrings =
    options.sharedStrings === undefined
      ? null
      : `${XML_HEADER}
<sst xmlns="${ns}" count="${options.sharedStrings.length}" uniqueCount="${options.sharedStrings.length}">
${options.sharedStrings.map((text) => `  <si><t>${escapeXml(text)}</t></si>`).join('\n')}
</sst>`;
  const workbook = `${XML_HEADER}
<workbook xmlns="${ns}" xmlns:r="${REL_NS}">
  <workbookPr date1904="0"/>
  <sheets><sheet name="${escapeXml(sheetName)}" sheetId="1" r:id="rId1"/></sheets>
  ${options.workbookExtras ?? ''}
</workbook>`;
  const workbookRels = `${XML_HEADER}
<Relationships xmlns="${PACKAGE_REL_NS}">
  <Relationship Id="rId1" Type="${REL_NS}/worksheet" Target="worksheets/sheet1.xml"/>
  ${sharedStrings === null ? '' : `<Relationship Id="rId2" Type="${REL_NS}/sharedStrings" Target="sharedStrings.xml"/>`}
  ${options.stylesXml === undefined ? '' : `<Relationship Id="rId3" Type="${REL_NS}/styles" Target="styles.xml"/>`}
  ${options.workbookRelsExtras ?? ''}
</Relationships>`;
  const overrides = [
    [
      'xl/workbook.xml',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
    ],
    [
      'xl/worksheets/sheet1.xml',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
    ],
    ...(sharedStrings === null
      ? []
      : [
          [
            'xl/sharedStrings.xml',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
          ] as [string, string],
        ]),
    ...(options.stylesXml === undefined
      ? []
      : [
          [
            'xl/styles.xml',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
          ] as [string, string],
        ]),
    ...(options.contentTypeOverrides ?? []),
  ];
  const contentTypes = `${XML_HEADER}
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  ${overrides.map(([part, type]) => `<Override PartName="/${part}" ContentType="${type}"/>`).join('\n  ')}
</Types>`;
  const rootRels = `${XML_HEADER}
<Relationships xmlns="${PACKAGE_REL_NS}">
  <Relationship Id="rId1" Type="${REL_NS}/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;
  const files: Record<string, string> = {
    '[Content_Types].xml': contentTypes,
    '_rels/.rels': rootRels,
    'xl/workbook.xml': workbook,
    'xl/_rels/workbook.xml.rels': workbookRels,
    'xl/worksheets/sheet1.xml': worksheetXml,
    ...(sharedStrings === null ? {} : { 'xl/sharedStrings.xml': sharedStrings }),
    ...(options.stylesXml === undefined ? {} : { 'xl/styles.xml': options.stylesXml }),
    ...(options.extraParts ?? {}),
  };
  return zipSync(
    Object.fromEntries(
      Object.entries(files).map(([name, content]) => [name, new TextEncoder().encode(content)]),
    ),
    { level: 6, mtime: new Date('1980-01-01T00:00:00Z') },
  );
}

export interface MultiSheetSpec {
  readonly name: string;
  /** Worksheet XML including the `<worksheet>` root. */
  readonly worksheetXml: string;
  readonly state?: 'visible' | 'hidden' | 'veryHidden';
  /** Defaults to `xl/worksheets/sheetN.xml`. */
  readonly partPath?: string;
}

export interface BuildMultiSheetOptions {
  readonly sharedStrings?: string[];
  readonly stylesXml?: string;
  readonly extraParts?: Record<string, string>;
  readonly workbookExtras?: string;
}

/**
 * A multi-sheet package with the tab order, hidden state and section
 * relationships a workbook-structure test needs. Kept byte-shape compatible
 * with `buildXlsx` so the same helpers can read it back.
 */
export function buildMultiSheetXlsx(
  sheets: readonly MultiSheetSpec[],
  options: BuildMultiSheetOptions = {},
): Uint8Array {
  if (sheets.length === 0) throw new Error('A workbook needs at least one sheet');
  const ns = SPREADSHEET_NS;
  const sheetRelEntries = sheets.map(
    (sheet, index) =>
      `<Relationship Id="rId${index + 1}" Type="${REL_NS}/worksheet" Target="worksheets/${sheetFile(sheet, index)}"/>`,
  );
  const sheetElements = sheets
    .map((sheet, index) => {
      const state =
        sheet.state === undefined || sheet.state === 'visible' ? '' : ` state="${sheet.state}"`;
      return `    <sheet name="${escapeXml(sheet.name)}" sheetId="${index + 1}" r:id="rId${index + 1}"${state}/>`;
    })
    .join('\n');
  const workbook = `${XML_HEADER}
<workbook xmlns="${ns}" xmlns:r="${REL_NS}">
  <workbookPr date1904="0"/>
  <sheets>
${sheetElements}
  </sheets>
  ${options.workbookExtras ?? ''}
</workbook>`;

  const sharedStrings =
    options.sharedStrings === undefined
      ? null
      : `${XML_HEADER}
<sst xmlns="${ns}" count="${options.sharedStrings.length}" uniqueCount="${options.sharedStrings.length}">
${options.sharedStrings.map((text) => `  <si><t>${escapeXml(text)}</t></si>`).join('\n')}
</sst>`;
  const nextRelationshipId = sheets.length + 1;
  const workbookRels = `${XML_HEADER}
<Relationships xmlns="${PACKAGE_REL_NS}">
  ${sheetRelEntries.join('\n  ')}
  ${sharedStrings === null ? '' : `<Relationship Id="rId${nextRelationshipId}" Type="${REL_NS}/sharedStrings" Target="sharedStrings.xml"/>`}
  ${options.stylesXml === undefined ? '' : `<Relationship Id="rId${nextRelationshipId + (sharedStrings === null ? 0 : 1)}" Type="${REL_NS}/styles" Target="styles.xml"/>`}
</Relationships>`;

  const overrides = [
    [
      'xl/workbook.xml',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
    ],
    ...sheets.map(
      (sheet, index) =>
        [
          sheetPartName(sheet, index),
          'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
        ] as [string, string],
    ),
    ...(sharedStrings === null
      ? []
      : [
          [
            'xl/sharedStrings.xml',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
          ] as [string, string],
        ]),
    ...(options.stylesXml === undefined
      ? []
      : [
          [
            'xl/styles.xml',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
          ] as [string, string],
        ]),
  ];
  const contentTypes = `${XML_HEADER}
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  ${overrides.map(([part, type]) => `<Override PartName="/${part}" ContentType="${type}"/>`).join('\n  ')}
</Types>`;
  const rootRels = `${XML_HEADER}
<Relationships xmlns="${PACKAGE_REL_NS}">
  <Relationship Id="rId1" Type="${REL_NS}/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;
  const files: Record<string, string> = {
    '[Content_Types].xml': contentTypes,
    '_rels/.rels': rootRels,
    'xl/workbook.xml': workbook,
    'xl/_rels/workbook.xml.rels': workbookRels,
    ...(sharedStrings === null ? {} : { 'xl/sharedStrings.xml': sharedStrings }),
    ...(options.stylesXml === undefined ? {} : { 'xl/styles.xml': options.stylesXml }),
    ...(options.extraParts ?? {}),
  };
  sheets.forEach((sheet, index) => {
    files[sheetPartName(sheet, index)] = sheet.worksheetXml;
  });
  return zipSync(
    Object.fromEntries(
      Object.entries(files).map(([name, content]) => [name, new TextEncoder().encode(content)]),
    ),
    { level: 6, mtime: new Date('1980-01-01T00:00:00Z') },
  );
}

function sheetPartName(sheet: MultiSheetSpec, index: number): string {
  return sheet.partPath ?? `xl/worksheets/sheet${index + 1}.xml`;
}

function sheetFile(sheet: MultiSheetSpec, index: number): string {
  const path = sheetPartName(sheet, index);
  return path.slice(path.lastIndexOf('/') + 1);
}

export function unzipPart(bytes: Uint8Array, name: string): string {
  const files = unzipParts(bytes);
  const part = files.get(name);
  if (part === undefined) throw new Error(`Missing part ${name}`);
  return new TextDecoder().decode(part);
}

export function unzipParts(bytes: Uint8Array): Map<string, Uint8Array> {
  return new Map(Object.entries(unzipSync(bytes)));
}

export function directChildNames(xml: string, parent = 'worksheet'): string[] {
  const parser = new SaxesParser({ xmlns: true });
  const names: string[] = [];
  let depth = 0;
  let parentDepth = -1;
  parser.on('opentag', (tag) => {
    depth += 1;
    if (tag.local === parent) parentDepth = depth;
    else if (parentDepth > 0 && depth === parentDepth + 1) names.push(tag.local);
  });
  parser.on('closetag', () => {
    depth -= 1;
  });
  parser.write(xml).close();
  return names;
}

export function elementRanges(xml: string, local: string): { start: number; end: number }[] {
  const parser = new SaxesParser({ xmlns: true, position: true });
  const ranges: { start: number; end: number }[] = [];
  const stack: { local: string; start: number }[] = [];
  parser.on('opentag', (tag) => {
    const start = xml.lastIndexOf('<', Math.max(0, parser.position - 1));
    stack.push({ local: tag.local, start });
  });
  parser.on('closetag', (tag) => {
    const open = stack.pop();
    if (tag.local === local && open !== undefined) {
      ranges.push({ start: open.start, end: parser.position });
    }
  });
  parser.write(xml).close();
  return ranges;
}

function escapeXml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;');
}
