import { describe, expect, it } from 'vitest';
import { createEmptyWorkbook, createSheet } from '@model/workbook';
import { errorValue, numberValue, textValue } from '@model/value';
import { createCell } from '@model/types';
import type { Sheet, Workbook } from '@model/types';
import type { CompatibilityInfo } from '@ooxml/open';
import type { PackagePatchReport, WorkbookPackage } from '@ooxml/types';
import {
  EXPORT_REPORT_CATEGORIES,
  buildCompatibilityReport,
  buildExportReport,
  classifyPreservedPart,
  compatibilityReportSummary,
  exportReportSummary,
  formatCompatibilityReportText,
  formatExportReportText,
  isReportablePart,
} from '@ooxml/report';
import type { PreservedPartStatus } from '@ooxml/report';
import { compatibilitySections, exportSections } from '@chrome/report-view';

/**
 * M6 lane 6 — compatibility and export reports (file-compat §7, §10 criteria
 * 9 and 10).
 *
 * The tests are synthetic where a hook is involved: packages are built as
 * plain part maps, so a stale/transformed/removed classification is asserted
 * without depending on the save pipeline or on lane 4's listing. The real
 * fixtures are covered by tests/m6-producers.test.ts.
 */

const encoder = new TextEncoder();

function packageOf(parts: Record<string, string>): WorkbookPackage {
  return {
    originalBytes: null,
    entries: new Map(Object.entries(parts).map(([name, text]) => [name, encoder.encode(text)])),
    contentTypesPath: '[Content_Types].xml',
    rootRelsPath: '_rels/.rels',
    workbookPartPath: 'xl/workbook.xml',
    workbookRelsPath: 'xl/_rels/workbook.xml.rels',
    sharedStringsPath: null,
    stylesPath: null,
    themePath: null,
    sheetPartPaths: new Map(),
    sheetRelationshipIds: new Map(),
    relationships: new Map(),
    sharedStrings: null,
    externalLinkParts: [],
  };
}

function compatibilityOf(overrides: Partial<CompatibilityInfo> = {}): CompatibilityInfo {
  return {
    opaquePartPaths: [],
    externalLinkParts: [],
    hasMacros: false,
    hasDigitalSignatures: false,
    notes: [],
    ...overrides,
  };
}

function saveReport(overrides: Partial<PackagePatchReport> = {}): PackagePatchReport {
  return {
    patchedParts: [],
    modeByPart: new Map(),
    appendedSharedStrings: 0,
    removedParts: [],
    ...overrides,
  };
}

function workbookWith(parts: string[], sheet: Sheet | null = null): Workbook {
  const workbook = createEmptyWorkbook();
  for (const path of parts) {
    workbook.opaqueParts.set(path, {
      path,
      bytes: encoder.encode(`<part path="${path}"/>`),
      anchor: null,
    });
  }
  if (sheet !== null) workbook.sheets.push(sheet);
  return workbook;
}

describe('the five export categories', () => {
  it('are exactly §7, in report order', () => {
    expect(EXPORT_REPORT_CATEGORIES).toEqual([
      'unchanged',
      'transformed',
      'stale',
      'removed',
      'warning',
    ]);
  });
});

describe('export report', () => {
  const chart = 'xl/charts/chart1.xml';
  const image = 'xl/media/image1.png';
  const link = 'xl/externalLinks/externalLink1.xml';
  const macro = 'xl/vbaProject.bin';

  it('reports nothing for a no-op save', () => {
    const workbook = workbookWith([chart]);
    const before = packageOf({ [chart]: 'a' });
    const report = buildExportReport({
      workbook,
      before,
      after: before,
      report: saveReport(),
      compatibility: compatibilityOf(),
      noOp: true,
    });
    expect(report.entries).toEqual([]);
    expect(report.needsAttention).toBe(false);
    expect(report.text).toContain('Nothing to report');
  });

  it('separates unchanged, transformed, stale, removed and warning entries', () => {
    const workbook = workbookWith([chart, image, link, macro]);
    const sheet = createSheet('Sheet1', 'xl/worksheets/sheet1.xml');
    sheet.protection = { rawXml: '<sheetProtection/>' };
    workbook.sheets.push(sheet);
    const before = packageOf({
      [chart]: 'chart-v1',
      [image]: 'png',
      [link]: 'link-v1',
      [macro]: 'vba',
    });
    const after = packageOf({ [chart]: 'chart-v2', [image]: 'png', [link]: 'link-v1' });
    const report = buildExportReport({
      workbook,
      before,
      after,
      report: saveReport({ removedParts: [macro] }),
      compatibility: compatibilityOf({
        hasMacros: true,
        hasDigitalSignatures: true,
        externalLinkParts: [link],
      }),
      noOp: false,
      structuralChange: true,
      formatChange: { from: 'xlsm', to: 'xlsx' },
    });

    expect(report.counts).toEqual({
      unchanged: 1,
      transformed: 1,
      stale: 1,
      removed: 1,
      warning: 2,
    });
    expect(report.needsAttention).toBe(true);
    expect(report.entries.map((entry) => entry.category)).toEqual([
      'unchanged',
      'transformed',
      'stale',
      'removed',
      'warning',
      'warning',
    ]);
    expect(report.entries[0]!.part).toBe(image);
    expect(report.entries[1]!.part).toBe(chart);
    expect(report.entries[2]!.part).toBe(link);
    expect(report.entries[2]!.detail).toMatch(/not propagated to the external reference/);
    expect(report.entries[3]!.part).toBe(macro);
    expect(report.entries[3]!.detail).toMatch(/saved as \.xlsx/);
    expect(report.entries.some((entry) => entry.title.includes('signature'))).toBe(true);
    expect(report.entries.some((entry) => entry.title === 'Protection not enforced')).toBe(true);
    expect(report.text).toContain('preserved but stale (1)');
    expect(report.text).toContain('removed (1)');
  });

  it('does not open for unchanged and transformed parts alone', () => {
    const workbook = workbookWith([chart]);
    const before = packageOf({ [chart]: 'chart-v1' });
    const after = packageOf({ [chart]: 'chart-v2' });
    const report = buildExportReport({
      workbook,
      before,
      after,
      report: saveReport({ patchedParts: [chart], modeByPart: new Map([[chart, 'tree']]) }),
      compatibility: compatibilityOf(),
      noOp: false,
    });
    expect(report.counts.transformed).toBe(1);
    expect(report.needsAttention).toBe(false);
  });

  it('does not invent staleness when no structural edit happened', () => {
    const workbook = workbookWith([link]);
    const before = packageOf({ [link]: 'link' });
    const report = buildExportReport({
      workbook,
      before,
      after: before,
      report: saveReport(),
      compatibility: compatibilityOf({
        externalLinkParts: [link],
        opaquePartPaths: [link],
      }),
      noOp: false,
    });
    expect(report.counts.stale).toBe(0);
    expect(report.counts.unchanged).toBe(1);
  });

  it('consumes lane 4’s preserved-part listing, including its stale and warning statuses', () => {
    const workbook = workbookWith([chart, link]);
    const before = packageOf({ [chart]: 'chart', [link]: 'link' });
    const listing: PreservedPartStatus[] = [
      {
        path: chart,
        category: 'chart',
        status: 'transformed',
        description: 'Chart 1',
        note: 'Series rewritten.',
      },
      {
        path: link,
        category: 'externalLink',
        status: 'stale',
        description: 'External link 1',
        note: 'Never refreshed.',
      },
      {
        path: 'xl/pivotCache/pivotCacheDefinition1.xml',
        category: 'pivotCache',
        status: 'warning',
        description: 'Pivot cache',
        note: 'Not refreshed.',
      },
    ];
    const report = buildExportReport({
      workbook,
      before,
      after: before,
      report: saveReport(),
      compatibility: compatibilityOf(),
      noOp: false,
      preserved: listing,
    });
    const transformed = report.entries.find((entry) => entry.category === 'transformed');
    expect(transformed?.title).toBe('Chart 1');
    expect(transformed?.detail).toBe('Series rewritten.');
    expect(report.entries.some((entry) => entry.category === 'stale' && entry.part === link)).toBe(
      true,
    );
    // A part the model does not hold still reaches the warning list through the
    // listing, because the listing is the save pipeline's own account.
    expect(report.counts.stale).toBe(1);
    expect(report.counts.warning).toBe(1);
    expect(report.entries.some((entry) => entry.title === 'Pivot cache')).toBe(true);
  });

  it('ignores package plumbing', () => {
    const workbook = workbookWith([
      '[Content_Types].xml',
      '_rels/.rels',
      'xl/_rels/workbook.xml.rels',
      'xl/calcChain.xml',
      'docProps/core.xml',
      'docProps/app.xml',
      'docProps/custom.xml',
    ]);
    const before = packageOf({ 'docProps/custom.xml': 'custom' });
    const after = packageOf({ 'docProps/custom.xml': 'custom' });
    const report = buildExportReport({
      workbook,
      before,
      after,
      report: saveReport(),
      compatibility: compatibilityOf(),
      noOp: false,
    });
    expect(report.entries.map((entry) => entry.part)).toEqual(['docProps/custom.xml']);
  });
});

describe('compatibility report', () => {
  it('is empty for a clean workbook', () => {
    const workbook = createEmptyWorkbook();
    workbook.sheets.push(createSheet('Sheet1', 'xl/worksheets/sheet1.xml'));
    const report = buildCompatibilityReport({
      workbook,
      compatibility: compatibilityOf(),
    });
    expect(report.hasContent).toBe(false);
    expect(report.items).toEqual([]);
    expect(report.text).toContain('Nothing to report');
  });

  it('lists Tier B parts, links, macros, signatures, protection, unsupported rules and stale cells', () => {
    const workbook = workbookWith([
      'xl/charts/chart1.xml',
      'xl/drawings/drawing1.xml',
      'xl/media/image1.png',
      'xl/pivotTables/pivotTable1.xml',
      'xl/externalLinks/externalLink1.xml',
      'xl/vbaProject.bin',
      '_xmlsignatures/sig1.xml',
      'xl/tables/table1.xml',
    ]);
    const sheet = createSheet('Sheet1', 'xl/worksheets/sheet1.xml');
    sheet.protection = { rawXml: '<sheetProtection/>' };
    const staleCell = {
      ...createCell(numberValue(42)),
      formula: '_xlfn.XLOOKUP(A1,B:B,C:C)',
      savedResult: numberValue(42),
      stale: true,
    };
    sheet.cells.set('0:0', staleCell);
    const nameError = { ...createCell(errorValue('#NAME?')), stale: true };
    sheet.cells.set('1:1', nameError);
    workbook.sheets.push(sheet);
    const protectedWorkbook: Workbook = {
      ...workbook,
      protection: { rawXml: '<workbookProtection/>' },
    };
    const compatibility = compatibilityOf({
      opaquePartPaths: [
        'xl/charts/chart1.xml',
        'xl/drawings/drawing1.xml',
        'xl/media/image1.png',
        'xl/pivotTables/pivotTable1.xml',
        'xl/externalLinks/externalLink1.xml',
        'xl/vbaProject.bin',
        '_xmlsignatures/sig1.xml',
        'xl/tables/table1.xml',
      ],
      externalLinkParts: ['xl/externalLinks/externalLink1.xml'],
      hasMacros: true,
      hasDigitalSignatures: true,
      notes: ['Charts are preserved but not rendered in this milestone.'],
    });

    const report = buildCompatibilityReport({
      workbook: protectedWorkbook,
      compatibility,
      unsupportedConditionalFormats: [
        {
          sheetId: sheet.id,
          sheetName: 'Sheet1',
          ranges: [{ startRow: 0, startCol: 0, endRow: 2, endCol: 1 }],
          ruleType: 'aboveAverage',
          reason: 'aboveAverage',
          source: 'base',
        },
      ],
    });

    expect(report.hasContent).toBe(true);
    expect(report.counts.chart).toBe(1);
    expect(report.counts.drawing).toBe(1);
    expect(report.counts.image).toBe(1);
    expect(report.counts.pivotTable).toBe(1);
    expect(report.counts.externalLink).toBe(1);
    expect(report.counts.macro).toBe(1);
    expect(report.counts.signature).toBe(1);
    expect(report.counts.table).toBe(1);
    expect(report.counts.protection).toBe(2);
    expect(report.counts.unsupportedConditionalFormat).toBe(1);
    expect(report.counts.unsupportedFunction).toBe(2);
    expect(report.items.find((item) => item.kind === 'unsupportedConditionalFormat')?.anchor).toBe(
      'A1:B3',
    );
    expect(report.items.find((item) => item.anchor === 'A1')?.detail).toMatch(/cached result/);
    expect(report.items.find((item) => item.anchor === 'B2')?.detail).toMatch(/#NAME\?/);
    expect(report.text).toContain('Compatibility report');
    expect(report.text).toContain('Notes');
    expect(compatibilityReportSummary(report)).toMatch(/1 chart/);
  });

  it('ignores plumbing and takes lane 4’s description and note when supplied', () => {
    const workbook = workbookWith(['xl/charts/chart1.xml']);
    const compatibility = compatibilityOf({
      opaquePartPaths: [
        '[Content_Types].xml',
        '_rels/.rels',
        'docProps/core.xml',
        'docProps/app.xml',
        'xl/calcChain.xml',
        'xl/charts/chart1.xml',
      ],
    });
    const report = buildCompatibilityReport({
      workbook,
      compatibility,
      preserved: [
        {
          path: 'xl/charts/chart1.xml',
          category: 'chart',
          status: 'unchanged',
          description: 'Chart (from the part policy)',
          note: 'Lane 4 note.',
        },
      ],
    });
    expect(report.items).toHaveLength(1);
    expect(report.items[0]!.title).toBe('Chart (from the part policy)');
    expect(report.items[0]!.detail).toBe('Lane 4 note.');
  });

  it('keeps notes but does not count them as content', () => {
    const workbook = createEmptyWorkbook();
    const report = buildCompatibilityReport({
      workbook,
      compatibility: compatibilityOf({
        notes: [
          'Locale-specific built-in number formats are shown in their Gregorian en-US form; the original numFmtId is preserved.',
        ],
      }),
    });
    expect(report.hasContent).toBe(false);
    expect(report.text).toMatch(/Locale-specific/);
  });
});

describe('classification helpers', () => {
  it('classifies preserved parts by path', () => {
    expect(classifyPreservedPart('xl/charts/chart1.xml')).toBe('chart');
    expect(classifyPreservedPart('xl/chartsheets/sheet1.xml')).toBe('chartSheet');
    expect(classifyPreservedPart('xl/media/image1.png')).toBe('image');
    expect(classifyPreservedPart('xl/drawings/drawing1.xml')).toBe('drawing');
    expect(classifyPreservedPart('xl/pivotTables/pivotTable1.xml')).toBe('pivotTable');
    expect(classifyPreservedPart('xl/pivotCache/pivotCacheDefinition1.xml')).toBe('pivotCache');
    expect(classifyPreservedPart('xl/externalLinks/externalLink1.xml')).toBe('externalLink');
    expect(classifyPreservedPart('xl/vbaProject.bin')).toBe('macro');
    expect(classifyPreservedPart('_xmlsignatures/sig1.xml')).toBe('signature');
    expect(classifyPreservedPart('xl/tables/table1.xml')).toBe('table');
    expect(classifyPreservedPart('customXml/item1.xml')).toBe('customXml');
    expect(classifyPreservedPart('docProps/custom.xml')).toBe('customProperty');
    expect(classifyPreservedPart('xl/unknown.xml')).toBe('other');
  });

  it('keeps package plumbing out of the reports', () => {
    expect(isReportablePart('[Content_Types].xml')).toBe(false);
    expect(isReportablePart('_rels/.rels')).toBe(false);
    expect(isReportablePart('xl/workbook.xml.rels')).toBe(false);
    expect(isReportablePart('xl/calcChain.xml')).toBe(false);
    expect(isReportablePart('docProps/core.xml')).toBe(false);
    expect(isReportablePart('docProps/custom.xml')).toBe(true);
    expect(isReportablePart('xl/charts/chart1.xml')).toBe(true);
  });
});

describe('dialog views', () => {
  it('group compatibility items by kind and export entries by category', () => {
    const workbook = workbookWith(['xl/charts/chart1.xml', 'xl/media/image1.png']);
    const sheet = createSheet('Sheet1', 'xl/worksheets/sheet1.xml');
    sheet.cells.set('0:0', { ...createCell(textValue('x')), stale: true });
    workbook.sheets.push(sheet);
    const compatibility = buildCompatibilityReport({
      workbook,
      compatibility: compatibilityOf({
        opaquePartPaths: ['xl/charts/chart1.xml', 'xl/media/image1.png'],
      }),
    });
    const compatibilityView = compatibilitySections(compatibility);
    expect(compatibilityView.map((section) => section.id)).toEqual([
      'chart',
      'image',
      'unsupportedFunction',
    ]);
    expect(compatibilityView[0]!.entries[0]!.title).toContain('xl/charts/chart1.xml');
    expect(compatibilityView[2]!.entries[0]!.where).toBe('Sheet1!A1');

    const before = packageOf({ 'xl/charts/chart1.xml': 'a', 'xl/media/image1.png': 'b' });
    const after = packageOf({ 'xl/charts/chart1.xml': 'c', 'xl/media/image1.png': 'b' });
    const exportReport = buildExportReport({
      workbook,
      before,
      after,
      report: saveReport(),
      compatibility: compatibilityOf(),
      noOp: false,
    });
    const exportView = exportSections(exportReport);
    expect(exportView.map((section) => section.title)).toEqual([
      'Unchanged (1)',
      'Transformed (1)',
    ]);
  });
});

describe('copy-as-text forms', () => {
  it('are stable and name each category', () => {
    const workbook = workbookWith(['xl/charts/chart1.xml']);
    const before = packageOf({ 'xl/charts/chart1.xml': 'a' });
    const after = packageOf({ 'xl/charts/chart1.xml': 'b' });
    const exportReport = buildExportReport({
      workbook,
      before,
      after,
      report: saveReport(),
      compatibility: compatibilityOf(),
      noOp: false,
    });
    expect(formatExportReportText(exportReport)).toContain('Export report');
    expect(formatExportReportText(exportReport)).toContain('transformed (1)');
    expect(exportReportSummary(exportReport)).toBe('1 transformed');

    const compatibility = buildCompatibilityReport({
      workbook,
      compatibility: compatibilityOf({ opaquePartPaths: ['xl/charts/chart1.xml'] }),
    });
    expect(formatCompatibilityReportText(compatibility)).toContain('[Chart]');
    expect(formatCompatibilityReportText(compatibility)).toContain('xl/charts/chart1.xml');
  });
});
