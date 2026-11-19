import { spawnSync } from 'node:child_process';
import { Buffer } from 'node:buffer';
import { describe, expect, it } from 'vitest';
import { DocumentSession } from '@commands/session';
import type { GridRange, Sheet } from '@model/types';
import { PIVOT_SOURCE_BLOCK_MESSAGE } from '@model/range-transform';
import { numberValue } from '@model/value';
import { openWorkbook, type OpenedWorkbook } from '@ooxml/open';
import { saveWorkbook } from '@ooxml/save';
import {
  listPreservedParts,
  listStaleParts,
  planPartPolicy,
  type PreservedPartReport,
} from '@ooxml/part-policy';
import { planMacroSave } from '@io/confirm';
import {
  assertByteIdentical,
  assertDirtyPartPreserved,
  assertOnlyPartsChanged,
  assertRelationshipsResolve,
  bytesEqual,
  fixtureBytes,
  openFixture,
  partBytesMap,
  partText,
  saveNoOp,
} from './helpers/roundtrip';
import { directChildNames } from './helpers/build-xlsx';

/**
 * M6 lane 4 — preserved parts and macros (file-compat §9 F9-F11, §10.2, .6,
 * .15, .16).
 *
 * Fixtures come from scripts/fixtures/groups/objects.mjs. The identity,
 * part-inventory and dirty-part measurements live in tests/helpers/roundtrip.ts
 * and never use the workbook model or the OOXML part parsers.
 */

const SHEET1 = 'xl/worksheets/sheet1.xml';
const WORKBOOK_PART = 'xl/workbook.xml';
const WORKBOOK_RELS = 'xl/_rels/workbook.xml.rels';
const CONTENT_TYPES = '[Content_Types].xml';
const CHART = 'xl/charts/chart1.xml';
const DRAWING = 'xl/drawings/drawing1.xml';
const DRAWING_RELS = 'xl/drawings/_rels/drawing1.xml.rels';
const MEDIA = 'xl/media/image1.png';
const VML = 'xl/drawings/vmlDrawing1.vml';
const OLE = 'xl/embeddings/oleObject1.bin';
const PIVOT_TABLE = 'xl/pivotTables/pivotTable1.xml';
const PIVOT_CACHE = 'xl/pivotCache/pivotCacheDefinition1.xml';
const PIVOT_RECORDS = 'xl/pivotCache/pivotCacheRecords1.xml';
const EXTERNAL_LINK = 'xl/externalLinks/externalLink1.xml';
const EXTERNAL_LINK_RELS = 'xl/externalLinks/_rels/externalLink1.xml.rels';
const VBA = 'xl/vbaProject.bin';

function range(startRow: number, startCol: number, endRow: number, endCol: number): GridRange {
  return { startRow, startCol, endRow, endCol };
}

function editCell(
  session: DocumentSession,
  sheet: Sheet,
  row: number,
  col: number,
  value: ReturnType<typeof numberValue>,
): void {
  session.dispatch('cell.set', { sheetId: sheet.id, row, col, value });
}

function saveSession(opened: OpenedWorkbook, session: DocumentSession) {
  return saveWorkbook({
    workbook: opened.workbook,
    package: opened.package,
    dirtyCells: session.dirtyCellMap(),
    dirtyDimensions: session.dirtyDimensionSnapshot(),
    dirtyMerges: new Set(session.dirtyMergeSnapshot().keys()),
  });
}

/** Every one of these parts must keep identical uncompressed bytes. */
function expectPartsIdentical(
  before: Uint8Array,
  after: Uint8Array,
  paths: readonly string[],
): void {
  const left = partBytesMap(before);
  const right = partBytesMap(after);
  for (const path of paths) {
    const original = left.get(path);
    const saved = right.get(path);
    expect(original, `${path} is missing from the source`).toBeDefined();
    expect(saved, `${path} is missing from the saved package`).toBeDefined();
    expect(bytesEqual(original!, saved!), `${path} is not byte-identical`).toBe(true);
  }
}

function listingByPath(listing: readonly PreservedPartReport[]): Map<string, PreservedPartReport> {
  return new Map(listing.map((entry) => [entry.path, entry]));
}

// ---------------------------------------------------------------------------
// F9 charts-images.xlsx
// ---------------------------------------------------------------------------

describe('F9 charts-images.xlsx — chart, image and shape parts survive an edit', () => {
  const original = fixtureBytes('charts-images.xlsx');
  const preserved = [CHART, DRAWING, DRAWING_RELS, MEDIA, VML, OLE];

  it('opens the chart, drawing, image, VML and OLE parts as opaque parts', () => {
    const opened = openFixture('charts-images.xlsx');
    const paths = new Set([
      ...opened.workbook.opaqueParts.keys(),
      ...[...opened.workbook.sheets[0]!.opaqueParts.keys()],
    ]);
    for (const path of preserved) expect(paths.has(path), path).toBe(true);
    expect([...opened.workbook.sheets[0]!.opaqueParts.keys()]).toContain(DRAWING);
  });

  it('is byte-identical on a no-op save and lists every preserved part', () => {
    const opened = openWorkbook(original);
    const saved = saveNoOp(opened);
    expect(saved.noOp).toBe(true);
    assertByteIdentical(original, saved.bytes, 'charts-images no-op save');
    assertOnlyPartsChanged(original, saved.bytes, []);
    assertRelationshipsResolve(saved.bytes);

    const listing = listingByPath(
      listPreservedParts({ package: opened.package, workbook: opened.workbook }),
    );
    for (const path of [CHART, DRAWING, MEDIA, VML, OLE]) {
      expect(listing.get(path)?.status, path).toBe('unchanged');
    }
    expect(listing.get(CHART)?.category).toBe('chart');
    expect(listing.get(MEDIA)?.category).toBe('media');
    expect(listing.get(OLE)?.category).toBe('oleObject');
  });

  it('keeps every preserved part byte-identical after an unrelated cell edit', () => {
    const opened = openFixture('charts-images.xlsx');
    const sheet = opened.workbook.sheets[0]!;
    const session = new DocumentSession(opened.workbook);
    editCell(session, sheet, 1, 3, numberValue(7));

    const saved = saveSession(opened, session);
    expect(saved.noOp).toBe(false);
    assertRelationshipsResolve(saved.bytes);
    assertOnlyPartsChanged(original, saved.bytes, [SHEET1]);
    expectPartsIdentical(original, saved.bytes, preserved);
  });

  it('keeps the dirty worksheet children (extLst, legacyDrawing, oleObjects) and their order', () => {
    const opened = openFixture('charts-images.xlsx');
    const sheet = opened.workbook.sheets[0]!;
    const session = new DocumentSession(opened.workbook);
    editCell(session, sheet, 1, 3, numberValue(7));
    const saved = saveSession(opened, session);

    const beforeXml = partText(original, SHEET1);
    const afterXml = partText(saved.bytes, SHEET1);
    // §10.15: every unmodelled child element and its order survives.
    expect(directChildNames(afterXml)).toEqual(directChildNames(beforeXml));
    expect(directChildNames(afterXml)).toEqual([
      'dimension',
      'sheetViews',
      'sheetFormatPr',
      'cols',
      'sheetData',
      'drawing',
      'legacyDrawing',
      'oleObjects',
      'extLst',
    ]);
    expect(afterXml).toContain('<legacyDrawing r:id="rId2"/>');
    expect(afterXml).toContain('<oleObject progId="Package" shapeId="1025" r:id="rId3"/>');
    expect(afterXml.slice(afterXml.indexOf('<extLst>'))).toBe(
      beforeXml.slice(beforeXml.indexOf('<extLst>')),
    );
    expect(afterXml).toContain('preserve-chart-me');
    assertDirtyPartPreserved(beforeXml, afterXml, {
      changedCells: ['D2'],
      changedRows: [2],
    });
    expect(afterXml).toContain('<c r="D2"><v>7</v></c>');
  });

  it('rewrites the chart series and the drawing anchors on an insert above the charted range', () => {
    const opened = openFixture('charts-images.xlsx');
    const sheet = opened.workbook.sheets[0]!;
    const session = new DocumentSession(opened.workbook);
    const result = session.dispatch(
      'structure.insertRows',
      {},
      {
        activeSheetId: sheet.id,
        selection: range(0, 0, 0, 0),
        activeCell: { row: 0, col: 0 },
      },
    );
    expect(result.structure?.state).toBe('applied');

    const saved = saveSession(opened, session);
    assertRelationshipsResolve(saved.bytes);
    expect(saved.noOp).toBe(false);
    // The chart series and the drawing anchors are transformed; the image,
    // VML and OLE parts are not anchored by this edit and stay verbatim.
    assertOnlyPartsChanged(original, saved.bytes, [SHEET1, CHART, DRAWING]);
    expectPartsIdentical(original, saved.bytes, [DRAWING_RELS, MEDIA, VML, OLE]);

    const chartAfter = partText(saved.bytes, CHART);
    expect(chartAfter).toContain('<c:f>Sheet1!$A$3:$A$5</c:f>');
    expect(chartAfter).toContain('<c:f>Sheet1!$B$3:$B$5</c:f>');
    expect(chartAfter).not.toContain('<c:f>Sheet1!$A$2:$A$4</c:f>');
    // The chart's cached points are preserved: the transform moves the
    // reference, never the last-rendered snapshot.
    expect(chartAfter).toContain('<c:pt idx="0"><c:v>10</c:v></c:pt>');

    const drawingAfter = partText(saved.bytes, DRAWING);
    expect(drawingAfter).toContain('<xdr:row>2</xdr:row>');
    expect(drawingAfter).toContain('<xdr:row>13</xdr:row>');
    expect(drawingAfter).not.toContain('<xdr:row>1</xdr:row>');

    // The dirty worksheet keeps its unmodelled children through the
    // structural patch as well.
    const beforeXml = partText(original, SHEET1);
    const afterXml = partText(saved.bytes, SHEET1);
    expect(directChildNames(afterXml)).toEqual(directChildNames(beforeXml));
    expect(afterXml.slice(afterXml.indexOf('<extLst>'))).toBe(
      beforeXml.slice(beforeXml.indexOf('<extLst>')),
    );

    const listing = listingByPath(
      listPreservedParts({
        package: opened.package,
        workbook: opened.workbook,
        transformedParts: saved.report.patchedParts,
        entries: opened.package.entries,
      }),
    );
    expect(listing.get(CHART)?.status).toBe('transformed');
    expect(listing.get(DRAWING)?.status).toBe('transformed');
    expect(listing.get(MEDIA)?.status).toBe('unchanged');
  });
});

// ---------------------------------------------------------------------------
// F10 pivot-external.xlsx
// ---------------------------------------------------------------------------

describe('F10 pivot-external.xlsx — pivots and external links are preserved', () => {
  const original = fixtureBytes('pivot-external.xlsx');
  const preserved = [
    PIVOT_TABLE,
    'xl/pivotTables/_rels/pivotTable1.xml.rels',
    PIVOT_CACHE,
    'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels',
    PIVOT_RECORDS,
    EXTERNAL_LINK,
    EXTERNAL_LINK_RELS,
  ];

  it('opens the pivot and external-link parts as opaque parts', () => {
    const opened = openFixture('pivot-external.xlsx');
    expect(opened.compatibility.externalLinkParts).toContain(EXTERNAL_LINK);
    const paths = new Set([
      ...opened.workbook.opaqueParts.keys(),
      ...[...opened.workbook.sheets[0]!.opaqueParts.keys()],
    ]);
    for (const path of preserved) expect(paths.has(path), path).toBe(true);
    expect(opened.workbook.sheets[0]!.cells.get('4:9')?.formula).toBe('[1]Linked!$A$1');
  });

  it('is byte-identical on a no-op save', () => {
    const opened = openWorkbook(original);
    const saved = saveNoOp(opened);
    expect(saved.noOp).toBe(true);
    assertByteIdentical(original, saved.bytes, 'pivot-external no-op save');
    assertOnlyPartsChanged(original, saved.bytes, []);
    assertRelationshipsResolve(saved.bytes);
  });

  it('keeps the pivot and external-link parts byte-identical after an unrelated edit', () => {
    const opened = openFixture('pivot-external.xlsx');
    const sheet = opened.workbook.sheets[0]!;
    const session = new DocumentSession(opened.workbook);
    // F2 is outside the pivot source range A1:C5.
    editCell(session, sheet, 1, 5, numberValue(5));

    const saved = saveSession(opened, session);
    assertRelationshipsResolve(saved.bytes);
    expect(saved.noOp).toBe(false);
    // The workbook part changes only for the recalculation flag the fixture's
    // external-reference formula requires.
    assertOnlyPartsChanged(original, saved.bytes, [WORKBOOK_PART, SHEET1]);
    expectPartsIdentical(original, saved.bytes, preserved);
    // External-workbook references are not rewritten (§6.2).
    expect(partText(saved.bytes, SHEET1)).toContain('<f>[1]Linked!$A$1</f>');
    expect(partText(saved.bytes, SHEET1)).toContain('<c r="J5"><f>[1]Linked!$A$1</f><v>99</v></c>');
  });

  it('blocks a pivot source-range edit and leaves the file untouched', () => {
    const opened = openFixture('pivot-external.xlsx');
    const sheet = opened.workbook.sheets[0]!;
    const session = new DocumentSession(opened.workbook);

    const insert = session.dispatch(
      'structure.insertRows',
      {},
      {
        activeSheetId: sheet.id,
        selection: range(3, 0, 3, 0),
        activeCell: { row: 3, col: 0 },
      },
    );
    expect(insert.structure?.state).toBe('blocked');
    expect(insert.structure?.message).toContain(PIVOT_SOURCE_BLOCK_MESSAGE);
    expect(insert.ui).toMatchObject({ type: 'toast' });

    const remove = session.dispatch(
      'structure.deleteColumns',
      {},
      {
        activeSheetId: sheet.id,
        selection: range(0, 1, 0, 1),
        activeCell: { row: 0, col: 1 },
      },
    );
    expect(remove.structure?.state).toBe('blocked');
    expect(remove.structure?.message).toContain(PIVOT_SOURCE_BLOCK_MESSAGE);
    expect(session.transactionCount).toBe(0);

    const saved = saveNoOp(opened);
    expect(saved.noOp).toBe(true);
    assertByteIdentical(original, saved.bytes, 'blocked pivot edit keeps the file');
  });

  it('allows a structural edit outside the pivot source range', () => {
    const opened = openFixture('pivot-external.xlsx');
    const sheet = opened.workbook.sheets[0]!;
    const session = new DocumentSession(opened.workbook);
    const result = session.dispatch(
      'structure.insertRows',
      {},
      {
        activeSheetId: sheet.id,
        selection: range(40, 0, 40, 0),
        activeCell: { row: 40, col: 0 },
      },
    );
    expect(result.structure?.state).toBe('applied');
    const saved = saveSession(opened, session);
    assertRelationshipsResolve(saved.bytes);
    expectPartsIdentical(original, saved.bytes, preserved);
  });

  it('marks the pivot cache stale after an edit inside its source range', () => {
    const opened = openFixture('pivot-external.xlsx');
    const sheet = opened.workbook.sheets[0]!;
    const session = new DocumentSession(opened.workbook);
    editCell(session, sheet, 2, 2, numberValue(21));

    const saved = saveSession(opened, session);
    // The cache itself is never rewritten; the listing reports the staleness.
    expectPartsIdentical(original, saved.bytes, preserved);

    const listing = listingByPath(
      listPreservedParts({
        package: opened.package,
        workbook: opened.workbook,
        dirtyCells: session.dirtyCellMap(),
      }),
    );
    expect(listing.get(PIVOT_CACHE)?.status).toBe('stale');
    expect(listing.get(PIVOT_CACHE)?.reason).toBe('pivot-cache-not-refreshed');
    expect(listing.get(PIVOT_TABLE)?.status).toBe('unchanged');
    expect(listing.get(EXTERNAL_LINK)?.status).toBe('stale');
    expect(listing.get(EXTERNAL_LINK)?.reason).toBe('external-link-not-refreshed');

    const stale = listStaleParts({
      package: opened.package,
      workbook: opened.workbook,
      dirtyCells: session.dirtyCellMap(),
    }).map((entry) => entry.path);
    expect(stale).toContain(PIVOT_CACHE);
    expect(stale).toContain(EXTERNAL_LINK);
    expect(stale).not.toContain(PIVOT_TABLE);
  });

  it('does not mark the pivot cache stale for an edit outside its source range', () => {
    const opened = openFixture('pivot-external.xlsx');
    const sheet = opened.workbook.sheets[0]!;
    const session = new DocumentSession(opened.workbook);
    editCell(session, sheet, 1, 5, numberValue(5));
    const listing = listingByPath(
      listPreservedParts({
        package: opened.package,
        workbook: opened.workbook,
        dirtyCells: session.dirtyCellMap(),
      }),
    );
    expect(listing.get(PIVOT_CACHE)?.status).toBe('unchanged');
  });
});

// ---------------------------------------------------------------------------
// F11 macros.xlsm
// ---------------------------------------------------------------------------

describe('F11 macros.xlsm — vbaProject.bin survives and never executes', () => {
  const original = fixtureBytes('macros.xlsm');

  it('opens as a macro workbook with the VBA project preserved as an opaque part', () => {
    const opened = openWorkbook(original, 'macros.xlsm');
    expect(opened.compatibility.hasMacros).toBe(true);
    expect(opened.compatibility.notes.join(' ')).toContain('Macros are preserved');
    const part = opened.workbook.opaqueParts.get(VBA);
    expect(part).toBeDefined();
    expect(bytesEqual(part!.bytes, partBytesMap(original).get(VBA)!)).toBe(true);

    // The macro payload never reaches the model: it is not inspected, parsed
    // into cells or executed — it is bytes in an opaque part.
    const model = JSON.stringify(opened.workbook.sheets.map((sheet) => [...sheet.cells.values()]));
    expect(model).not.toContain('vbaProject');
    expect(model).not.toContain('calcplus vbaProject fixture');
  });

  it('is byte-identical on a no-op save and after an unrelated edit', () => {
    const opened = openFixture('macros.xlsm');
    const noOp = saveNoOp(opened);
    expect(noOp.noOp).toBe(true);
    assertByteIdentical(original, noOp.bytes, 'macros no-op save');

    const sheet = opened.workbook.sheets[0]!;
    const session = new DocumentSession(opened.workbook);
    editCell(session, sheet, 1, 0, numberValue(21));
    const saved = saveSession(opened, session);
    assertRelationshipsResolve(saved.bytes);
    expect(saved.noOp).toBe(false);
    assertOnlyPartsChanged(original, saved.bytes, [WORKBOOK_PART, SHEET1]);
    expectPartsIdentical(original, saved.bytes, [VBA]);
    expect(partText(saved.bytes, CONTENT_TYPES)).toContain(
      'application/vnd.ms-excel.sheet.macroEnabled.main+xml',
    );

    const reopened = openWorkbook(saved.bytes, 'macros.xlsm');
    expect(reopened.compatibility.hasMacros).toBe(true);
    expect(
      bytesEqual(reopened.workbook.opaqueParts.get(VBA)!.bytes, partBytesMap(original).get(VBA)!),
    ).toBe(true);
  });

  it('removes the macro parts only for an explicit .xlsx save', () => {
    const opened = openFixture('macros.xlsm');
    const sheet = opened.workbook.sheets[0]!;
    const session = new DocumentSession(opened.workbook);
    editCell(session, sheet, 1, 0, numberValue(21));

    const saved = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
      options: { format: 'xlsx' },
    });
    expect(saved.noOp).toBe(false);
    assertRelationshipsResolve(saved.bytes);
    expect(saved.package.entries.has(VBA)).toBe(false);
    expect(saved.report.removedParts).toContain(VBA);
    assertOnlyPartsChanged(original, saved.bytes, [
      CONTENT_TYPES,
      WORKBOOK_RELS,
      WORKBOOK_PART,
      SHEET1,
      VBA,
    ]);

    const contentTypes = partText(saved.bytes, CONTENT_TYPES);
    expect(contentTypes).not.toContain('macroEnabled');
    expect(contentTypes).toContain(
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
    );
    expect(contentTypes).not.toContain('vbaProject');
    expect(partText(saved.bytes, WORKBOOK_RELS)).not.toContain('vbaProject');

    const reopened = openWorkbook(saved.bytes, 'macros.xlsx');
    expect(reopened.compatibility.hasMacros).toBe(false);
    expect(reopened.workbook.opaqueParts.has(VBA)).toBe(false);
    expect(reopened.workbook.sheets[0]!.cells.get('1:1')?.formula).toBe('A2*2');

    const listing = listingByPath(
      listPreservedParts({
        package: opened.package,
        workbook: opened.workbook,
        format: 'xlsx',
      }),
    );
    expect(listing.get(VBA)?.status).toBe('removed');
    expect(listing.get(VBA)?.reason).toBe('macros-removed');

    // The export report passes the post-save entries; the removed part is
    // still listed because the source package is the classification source.
    const afterSaveListing = listingByPath(
      listPreservedParts({
        package: opened.package,
        entries: saved.package.entries,
        workbook: opened.workbook,
        format: 'xlsx',
      }),
    );
    expect(afterSaveListing.get(VBA)?.status).toBe('removed');
  });

  it('never resurrects a macro part from an unadopted model mirror', () => {
    const opened = openFixture('macros.xlsm');
    const first = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: new Map(),
      options: { format: 'xlsx' },
    });
    // No adoptSavedPackage here: the model still holds the VBA bytes, so this
    // proves the policy keeps them out of the second write.
    expect(opened.workbook.opaqueParts.has(VBA)).toBe(true);
    const second = saveWorkbook({
      workbook: opened.workbook,
      package: first.package,
      dirtyCells: new Map(),
      options: { format: 'xlsx' },
    });
    expect(second.package.entries.has(VBA)).toBe(false);
    expect(second.bytes.byteLength).toBe(first.bytes.byteLength);
  });
});

describe('F11 macros.xlsm — the .xlsm to .xlsx confirmation', () => {
  it('keeps .xlsm on a plain Save without asking', () => {
    const asked: string[] = [];
    const plan = planMacroSave({
      fileName: 'macros.xlsm',
      hasMacros: true,
      saveAs: false,
      confirmFn: (message) => {
        asked.push(message);
        return true;
      },
    });
    expect(plan).toEqual({ format: 'xlsm', fileName: 'macros.xlsm' });
    expect(asked).toEqual([]);
  });

  it('warns and converts on a confirmed Save As', () => {
    const asked: string[] = [];
    const plan = planMacroSave({
      fileName: 'macros.xlsm',
      hasMacros: true,
      saveAs: true,
      confirmFn: (message) => {
        asked.push(message);
        return true;
      },
    });
    expect(plan).toEqual({ format: 'xlsx', fileName: 'macros.xlsx' });
    expect(asked).toHaveLength(1);
    expect(asked[0]).toContain('removes them');
  });

  it('writes nothing when the confirmation is declined', () => {
    const plan = planMacroSave({
      fileName: 'macros.xlsm',
      hasMacros: true,
      saveAs: true,
      confirmFn: () => false,
    });
    expect(plan).toBeNull();
  });

  it('does not ask for a workbook without macros', () => {
    const asked: string[] = [];
    const plan = planMacroSave({
      fileName: 'book.xlsx',
      hasMacros: false,
      saveAs: true,
      confirmFn: (message) => {
        asked.push(message);
        return true;
      },
    });
    expect(plan).toEqual({ format: 'xlsx', fileName: 'book.xlsx' });
    expect(asked).toEqual([]);
  });
});

// ---------------------------------------------------------------------------
// Part policy: the verbatim-versus-rewritten table
// ---------------------------------------------------------------------------

describe('part policy classifies the package', () => {
  it('marks the pipeline-owned parts rewritten and every Tier B part verbatim', () => {
    const opened = openFixture('charts-images.xlsx');
    const policy = planPartPolicy({ package: opened.package });
    expect(policy.sourceFormat).toBe('xlsx');
    expect(policy.format).toBe('xlsx');
    for (const path of [CONTENT_TYPES, WORKBOOK_PART, WORKBOOK_RELS, SHEET1]) {
      expect(policy.pipelineOwnedPartPaths.has(path), path).toBe(true);
    }
    for (const path of [CHART, DRAWING, MEDIA, VML, OLE, 'docProps/core.xml', 'docProps/app.xml']) {
      expect(policy.verbatimPartPaths.has(path), path).toBe(true);
      expect(policy.pipelineOwnedPartPaths.has(path), path).toBe(false);
    }
    expect(policy.removedPartPaths.size).toBe(0);
  });

  it('sources the format from the package and removes macro parts for xlsx', () => {
    const opened = openFixture('macros.xlsm');
    const retained = planPartPolicy({ package: opened.package });
    expect(retained.sourceFormat).toBe('xlsm');
    expect(retained.format).toBe('xlsm');
    expect(retained.verbatimPartPaths.has(VBA)).toBe(true);

    const converted = planPartPolicy({ package: opened.package, format: 'xlsx' });
    expect(converted.format).toBe('xlsx');
    expect(converted.removedPartPaths.has(VBA)).toBe(true);
    expect(converted.verbatimPartPaths.has(VBA)).toBe(false);
  });
});

// ---------------------------------------------------------------------------
// Cross-check: every fixture opens in an independent reader (openpyxl)
// ---------------------------------------------------------------------------

const hasOpenpyxl =
  spawnSync('python3', ['-c', 'import openpyxl'], { encoding: 'utf8' }).status === 0;

function openpyxlReads(bytes: Uint8Array, keepVba: boolean): string | null {
  const script = `import io, sys, openpyxl
data = sys.stdin.buffer.read()
workbook = openpyxl.load_workbook(io.BytesIO(data), keep_vba=${keepVba ? 'True' : 'False'}, data_only=True)
print(",".join(workbook.sheetnames))`;
  const result = spawnSync('python3', ['-c', script], { input: Buffer.from(bytes) });
  if (result.status === 0) return null;
  return String(result.stderr || result.stdout);
}

describe.skipIf(!hasOpenpyxl)('fixtures open in openpyxl', () => {
  it('charts-images.xlsx loads', () => {
    expect(openpyxlReads(fixtureBytes('charts-images.xlsx'), false)).toBeNull();
  });

  it('pivot-external.xlsx loads', () => {
    expect(openpyxlReads(fixtureBytes('pivot-external.xlsx'), false)).toBeNull();
  });

  it('macros.xlsm loads with the VBA project kept', () => {
    expect(openpyxlReads(fixtureBytes('macros.xlsm'), true)).toBeNull();
  });

  it('a converted macros.xlsx loads without the VBA project', () => {
    const opened = openFixture('macros.xlsm');
    const converted = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: new Map(),
      options: { format: 'xlsx' },
    });
    expect(openpyxlReads(converted.bytes, false)).toBeNull();
  });
});
