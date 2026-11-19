import { describe, expect, it } from 'vitest';
import {
  isMacroPartPath,
  partCategoryOf,
  planPartPolicy,
  sourceFormatOf,
  type PackagePartLocations,
} from '@ooxml/part-policy';
import { openFixture } from './helpers/roundtrip';

/**
 * M6 lane 4 — the part-policy table (file-compat §1-§4).
 *
 * The policy is the single statement of which package parts the save pipeline
 * rewrites and which it copies verbatim. These tests pin the table; the
 * fixture-level behaviour lives in tests/m6-preserved-parts.test.ts.
 */

const LOCATIONS: PackagePartLocations = {
  contentTypesPath: '[Content_Types].xml',
  workbookPartPath: 'xl/workbook.xml',
  workbookRelsPath: 'xl/_rels/workbook.xml.rels',
  stylesPath: 'xl/styles.xml',
  sharedStringsPath: 'xl/sharedStrings.xml',
  sheetPartPaths: new Map([
    ['sheet-1', 'xl/worksheets/sheet1.xml'],
    ['sheet-2', 'xl/worksheets/sheet2.xml'],
  ]),
};

describe('part categories', () => {
  it.each([
    ['[Content_Types].xml', 'contentTypes'],
    ['xl/workbook.xml', 'workbook'],
    ['xl/_rels/workbook.xml.rels', 'relationships'],
    ['xl/worksheets/sheet1.xml', 'worksheet'],
    ['xl/worksheets/sheet9.xml', 'worksheet'],
    ['xl/worksheets/_rels/sheet1.xml.rels', 'relationships'],
    ['xl/drawings/_rels/drawing1.xml.rels', 'relationships'],
    ['xl/pivotTables/_rels/pivotTable1.xml.rels', 'relationships'],
    ['xl/styles.xml', 'styles'],
    ['xl/sharedStrings.xml', 'sharedStrings'],
    ['xl/calcChain.xml', 'calcChain'],
    ['xl/theme/theme1.xml', 'theme'],
    ['docProps/core.xml', 'documentProperties'],
    ['docProps/app.xml', 'documentProperties'],
    ['docProps/custom.xml', 'customProperties'],
    ['xl/charts/chart1.xml', 'chart'],
    ['xl/chartsheets/sheet1.xml', 'chartSheet'],
    ['xl/drawings/drawing1.xml', 'drawing'],
    ['xl/drawings/vmlDrawing1.vml', 'vmlDrawing'],
    ['xl/media/image1.png', 'media'],
    ['xl/embeddings/oleObject1.bin', 'oleObject'],
    ['xl/pivotTables/pivotTable1.xml', 'pivotTable'],
    ['xl/pivotCache/pivotCacheDefinition1.xml', 'pivotCache'],
    ['xl/externalLinks/externalLink1.xml', 'externalLink'],
    ['xl/vbaProject.bin', 'vbaProject'],
    ['xl/vbaData.xml', 'vbaProject'],
    ['_xmlsignatures/sig1.xml', 'signature'],
    ['xl/tables/table1.xml', 'table'],
    ['xl/slicers/slicer1.xml', 'slicer'],
    ['xl/timelines/timeline1.xml', 'slicer'],
    ['customXml/item1.xml', 'customXml'],
  ])('classifies %s as %s', (path, category) => {
    expect(partCategoryOf(path, LOCATIONS)).toBe(category);
  });

  it('recognises macro parts by name', () => {
    expect(isMacroPartPath('xl/vbaProject.bin')).toBe(true);
    expect(isMacroPartPath('xl/vbaData.xml')).toBe(true);
    expect(isMacroPartPath('XL/VBAPROJECT.BIN')).toBe(true);
    expect(isMacroPartPath('xl/media/image1.png')).toBe(false);
  });
});

describe('the verbatim table', () => {
  it('derives the source format from the macro parts', () => {
    expect(sourceFormatOf(new Map())).toBe('xlsx');
    expect(sourceFormatOf(new Map([['xl/vbaProject.bin', new Uint8Array([1])]]))).toBe('xlsm');
  });

  it('splits pipeline-owned from verbatim parts on the chart fixture', () => {
    const opened = openFixture('charts-images.xlsx');
    const policy = planPartPolicy({ package: opened.package });
    expect([...policy.pipelineOwnedPartPaths].sort()).toEqual([
      '[Content_Types].xml',
      'xl/_rels/workbook.xml.rels',
      'xl/calcChain.xml',
      'xl/workbook.xml',
      'xl/worksheets/sheet1.xml',
    ]);
    for (const path of [
      'xl/charts/chart1.xml',
      'xl/drawings/drawing1.xml',
      'xl/drawings/vmlDrawing1.vml',
      'xl/media/image1.png',
      'xl/embeddings/oleObject1.bin',
      'docProps/core.xml',
      'docProps/app.xml',
    ]) {
      expect(policy.verbatimPartPaths.has(path), path).toBe(true);
    }
    expect(policy.removedPartPaths.size).toBe(0);
  });

  it('treats a macro workbook as xlsm unless the caller asks for xlsx', () => {
    const opened = openFixture('macros.xlsm');
    const retained = planPartPolicy({ package: opened.package });
    expect(retained.sourceFormat).toBe('xlsm');
    expect(retained.format).toBe('xlsm');
    expect(retained.verbatimPartPaths.has('xl/vbaProject.bin')).toBe(true);

    const converted = planPartPolicy({ package: opened.package, format: 'xlsx' });
    expect(converted.removedPartPaths.has('xl/vbaProject.bin')).toBe(true);
    expect(converted.verbatimPartPaths.has('xl/vbaProject.bin')).toBe(false);
    // Everything else keeps its disposition; the sheet and workbook stay owned.
    expect(converted.pipelineOwnedPartPaths.has('xl/worksheets/sheet1.xml')).toBe(true);
  });

  it('does not mark a plain workbook as converted', () => {
    const opened = openFixture('pivot-external.xlsx');
    const policy = planPartPolicy({ package: opened.package, format: 'xlsx' });
    expect(policy.sourceFormat).toBe('xlsx');
    expect(policy.removedPartPaths.size).toBe(0);
    expect(policy.verbatimPartPaths.has('xl/externalLinks/externalLink1.xml')).toBe(true);
  });
});
