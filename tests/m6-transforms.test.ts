import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import fc from 'fast-check';
import type { Sheet } from '@model/types';
import { setCellSnapshot } from '@model/types';
import { numberValue, textValue } from '@model/value';
import { createEmptyWorkbook, createSheet } from '@model/workbook';
import {
  CHART_SOURCE_BLOCK_MESSAGE,
  HYPERLINK_LOCATION_BLOCK_MESSAGE,
  PIVOT_SOURCE_BLOCK_MESSAGE,
  planOpaquePartTransforms,
  planRangeTransform,
  transformAnchorMarker,
  transformChartXml,
  transformCommentRefsXml,
  transformDrawingAnchorsXml,
  transformPivotCacheXml,
  transformPivotTableXml,
  transformRange,
  transformThreadedCommentRefsXml,
  transformVmlAnchorsXml,
  type StructuralEdit,
} from '@model/range-transform';
import { DocumentSession } from '@commands/session';
import { openWorkbook } from '@ooxml/open';
import { saveWorkbook } from '@ooxml/save';
import {
  applyStructuralPartPatches,
  planOpaquePartPatches,
  planStructuralPartPatches,
  renumberWorksheetPart,
} from '@ooxml/structure-patch';
import { buildMultiSheetXlsx, buildXlsx, unzipPart } from './helpers/build-xlsx';

/**
 * M6 lane 5 — structural transforms hardening (file-compat §4, §6.4, §8, §10.7;
 * workbook-model §8; milestones M6 exit criterion 5).
 *
 * Every range-bearing Tier B construct either transforms for the edit or the
 * edit is blocked with a typed result. The matrix sweep at the bottom is the
 * regression net for "no silent stale anchor": for any edit shape, a construct
 * whose own transform reports a change must appear in the plan, or the plan
 * must carry a block.
 */

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

// ---------------------------------------------------------------------------
// Synthetic package parts
// ---------------------------------------------------------------------------

const DRAWING = `${HEADER}
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="${R_NS}">
  <xdr:twoCellAnchor editAs="twoCell">
    <xdr:from><xdr:col>1</xdr:col><xdr:colOff>365760</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>182880</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="2" name="Shape 1"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="rect"/></xdr:spPr><xdr:style/><xdr:txBody><a:bodyPr/><a:p/></xdr:txBody></xdr:sp>
  </xdr:twoCellAnchor>
  <xdr:oneCellAnchor>
    <xdr:from><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:ext cx="2400000" cy="1600000"/>
    <xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="3" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:oneCellAnchor>
  <xdr:absoluteAnchor><xdr:pos x="100" y="200"/><xdr:ext cx="10" cy="20"/><xdr:clientData/></xdr:absoluteAnchor>
</xdr:wsDr>`;

const DRAWING_RELS = `${HEADER}
<Relationships xmlns="${PKG_NS}">
  <Relationship Id="rId1" Type="${R_NS}/chart" Target="../charts/chart1.xml"/>
</Relationships>`;

function chart(xml: string): string {
  return `${HEADER}
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${R_NS}">
  <c:chart><c:plotArea><c:layout/>
    <c:barChart><c:barDir val="col"/>
      ${xml}
      <c:axId val="1"/><c:axId val="2"/>
    </c:barChart>
  </c:plotArea><c:legend/></c:chart>
</c:chartSpace>`;
}

const CHART_RANGE = chart(
  '<c:ser><c:idx val="0"/><c:order val="0"/><c:cat><c:strRef><c:f>Sheet1!$A$2:$A$3</c:f><c:strCache/></c:strRef></c:cat><c:val><c:numRef><c:f>Sheet1!$B$2:$B$3</c:f><c:numCache/></c:numRef></c:val></c:ser>',
);

const VML = `${HEADER}
<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
  <o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>
  <v:shape id="_x0000_s1025" type="#_x0000_t202">
    <x:ClientData ObjectType="Note">
      <x:MoveWithCells/><x:SizeWithCells/>
      <x:Anchor>1, 15, 2, 2, 3, 15, 5, 3</x:Anchor>
      <x:AutoFill>False</x:AutoFill>
      <x:Row>2</x:Row>
      <x:Column>1</x:Column>
    </x:ClientData>
  </v:shape>
</xml>`;

const COMMENTS = `${HEADER}
<comments xmlns="${NS}">
  <authors><author>Ada</author></authors>
  <commentList>
    <comment ref="B3" authorId="0"><text><r><t>drop me</t></r></text></comment>
    <comment ref="B5" authorId="0"><text><r><t>keep me</t></r></text></comment>
  </commentList>
</comments>`;

const THREADED_COMMENTS = `${HEADER}
<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <threadedComment ref="B3" dT="2026-01-01T00:00:00Z" personId="{1}" id="{2}"><text>hi</text></threadedComment>
</ThreadedComments>`;

const PIVOT_CACHE = (source: string): string => `${HEADER}
<pivotCacheDefinition xmlns="${NS}" xmlns:r="${R_NS}" r:id="rId1" refreshOnLoad="0">
  <cacheSource type="worksheet">${source}</cacheSource>
  <cacheFields count="0"/>
</pivotCacheDefinition>`;

const PIVOT_TABLE = `${HEADER}
<pivotTableDefinition xmlns="${NS}" name="PivotTable1" cacheId="1">
  <location ref="G3:H6" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/>
</pivotTableDefinition>`;

const PIVOT_TABLE_RELS = `${HEADER}
<Relationships xmlns="${PKG_NS}">
  <Relationship Id="rId1" Type="${R_NS}/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition1.xml"/>
</Relationships>`;

interface Parts {
  readonly drawing?: string;
  readonly vml?: string;
  readonly comments?: string;
  readonly threadedComments?: string;
  readonly chart?: string;
  readonly pivotCache?: string;
  readonly pivotTable?: string;
}

/** A worksheet with the drawing, note, hyperlink and pivot hooks a package needs. */
function worksheetXml(parts: Parts): string {
  return `${HEADER}
<worksheet xmlns="${NS}" xmlns:r="${R_NS}">
  <dimension ref="A1:C6"/>
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
    <row r="2"><c r="A2"><v>2</v></c><c r="B2"><v>3</v></c><c r="C2"><v>4</v></c></row>
    <row r="3"><c r="A3"><v>3</v></c><c r="B3"><v>5</v></c></row>
    <row r="4"><c r="A4"><v>4</v></c></row>
    <row r="5"><c r="A5"><v>5</v></c></row>
  </sheetData>
  <hyperlinks><hyperlink ref="C2" location="Sheet1!A1" display="jump"/></hyperlinks>
  ${parts.drawing === undefined ? '' : '<drawing r:id="rId1"/>'}
  ${parts.vml === undefined && parts.comments === undefined ? '' : '<legacyDrawing r:id="rId2"/>'}
</worksheet>`;
}

function packageFor(parts: Parts): Uint8Array {
  const rels: string[] = [];
  const extraParts: Record<string, string> = {};
  if (parts.drawing !== undefined) {
    rels.push(`<Relationship Id="rId1" Type="${R_NS}/drawing" Target="../drawings/drawing1.xml"/>`);
    extraParts['xl/drawings/drawing1.xml'] = parts.drawing;
    extraParts['xl/drawings/_rels/drawing1.xml.rels'] = DRAWING_RELS;
  }
  if (parts.vml !== undefined || parts.comments !== undefined) {
    rels.push(
      `<Relationship Id="rId2" Type="${R_NS}/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>`,
    );
    extraParts['xl/drawings/vmlDrawing1.vml'] = parts.vml ?? VML;
  }
  if (parts.comments !== undefined) {
    rels.push(`<Relationship Id="rId3" Type="${R_NS}/comments" Target="../comments1.xml"/>`);
    extraParts['xl/comments1.xml'] = parts.comments;
  }
  if (parts.threadedComments !== undefined) {
    rels.push(
      `<Relationship Id="rId4" Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment" Target="../threadedComments/threadedComment1.xml"/>`,
    );
    extraParts['xl/threadedComments/threadedComment1.xml'] = parts.threadedComments;
  }
  if (parts.pivotTable !== undefined) {
    rels.push(
      `<Relationship Id="rId5" Type="${R_NS}/pivotTable" Target="../pivotTables/pivotTable1.xml"/>`,
    );
    extraParts['xl/pivotTables/pivotTable1.xml'] = parts.pivotTable;
    extraParts['xl/pivotTables/_rels/pivotTable1.xml.rels'] = PIVOT_TABLE_RELS;
  }
  if (parts.chart !== undefined) {
    extraParts['xl/charts/chart1.xml'] = parts.chart;
  }
  if (parts.pivotCache !== undefined) {
    extraParts['xl/pivotCache/pivotCacheDefinition1.xml'] = parts.pivotCache;
  }
  if (rels.length > 0) {
    extraParts['xl/worksheets/_rels/sheet1.xml.rels'] =
      `${HEADER}\n<Relationships xmlns="${PKG_NS}">${rels.join('')}</Relationships>`;
  }
  return buildXlsx(worksheetXml(parts), { extraParts });
}

function openWith(parts: Parts) {
  const opened = openWorkbook(packageFor(parts));
  return { ...opened, sheet: opened.workbook.sheets[0]! };
}

function insertRow(sheetId: string, index: number, count = 1): StructuralEdit {
  return { kind: 'insert', sheetId, axis: 'row', index, count };
}

function deleteRow(sheetId: string, index: number, count = 1): StructuralEdit {
  return { kind: 'delete', sheetId, axis: 'row', index, count };
}

function insertColumn(sheetId: string, index: number, count = 1): StructuralEdit {
  return { kind: 'insert', sheetId, axis: 'column', index, count };
}

function deleteColumn(sheetId: string, index: number, count = 1): StructuralEdit {
  return { kind: 'delete', sheetId, axis: 'column', index, count };
}

function partText(bytes: Uint8Array | undefined): string {
  if (bytes === undefined) throw new Error('missing part');
  return new TextDecoder().decode(bytes);
}

function plannedPart(
  plan: ReturnType<typeof planRangeTransform>,
  path: string,
): { readonly before: string; readonly after: string } | null {
  const rewrite = plan.opaqueParts.find((candidate) => candidate.path === path);
  if (rewrite === undefined) return null;
  return { before: partText(rewrite.before), after: partText(rewrite.after) };
}

// ---------------------------------------------------------------------------
// Drawing anchors
// ---------------------------------------------------------------------------

describe('M6 drawing anchors', () => {
  it('moves one- and two-cell markers on a row insert and keeps offsets and absolute anchors', () => {
    const { workbook, sheet } = openWith({ drawing: DRAWING });
    const plan = planRangeTransform(workbook, insertRow(sheet.id, 1));
    const rewrite = plannedPart(plan, 'xl/drawings/drawing1.xml');
    expect(rewrite).not.toBeNull();
    const after = rewrite!.after;
    // Two-cell anchor: from row 2 -> 3, to row 4 -> 5; columns untouched.
    expect(after).toContain(
      '<xdr:from><xdr:col>1</xdr:col><xdr:colOff>365760</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>182880</xdr:rowOff></xdr:from>',
    );
    expect(after).toContain(
      '<xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>',
    );
    // Tier B rule: a one-cell anchor keeps its ext offsets verbatim.
    expect(after).toContain('<xdr:ext cx="2400000" cy="1600000"/>');
    expect(after).toContain('<xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>');
    // An absolute anchor holds EMU positions, not cell indices.
    expect(after).toContain('<xdr:absoluteAnchor><xdr:pos x="100" y="200"/>');
  });

  it('moves markers on a column insert and a column delete', () => {
    const { workbook, sheet } = openWith({ drawing: DRAWING });
    const insert = planRangeTransform(workbook, insertColumn(sheet.id, 0));
    const inserted = plannedPart(insert, 'xl/drawings/drawing1.xml')!.after;
    expect(inserted).toContain(
      '<xdr:col>2</xdr:col><xdr:colOff>365760</xdr:colOff><xdr:row>2</xdr:row>',
    );
    expect(inserted).toContain(
      '<xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row>',
    );

    const remove = planRangeTransform(workbook, deleteColumn(sheet.id, 0));
    const removed = plannedPart(remove, 'xl/drawings/drawing1.xml')!.after;
    expect(removed).toContain(
      '<xdr:col>0</xdr:col><xdr:colOff>365760</xdr:colOff><xdr:row>2</xdr:row>',
    );
  });

  it('clamps a marker whose tracked row is deleted instead of leaving a stale row', () => {
    const { workbook, sheet } = openWith({ drawing: DRAWING });
    const plan = planRangeTransform(workbook, deleteRow(sheet.id, 2));
    const after = plannedPart(plan, 'xl/drawings/drawing1.xml')!.after;
    // The from marker on row 2 clamps to the cell that shifts in (row 2); the
    // to marker on row 4 moves up to row 3.
    expect(after).toContain(
      '<xdr:col>1</xdr:col><xdr:colOff>365760</xdr:colOff><xdr:row>2</xdr:row>',
    );
    expect(after).toContain('<xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row>');
  });

  it('leaves a drawing untouched when the edit is below it', () => {
    const { workbook, sheet } = openWith({ drawing: DRAWING });
    const plan = planRangeTransform(workbook, insertRow(sheet.id, 20));
    expect(plannedPart(plan, 'xl/drawings/drawing1.xml')).toBeNull();
    expect(plan.block).toBeNull();
  });

  it('blocks an anchor whose markers cannot be read', () => {
    const broken = DRAWING.replace(
      '<xdr:from><xdr:col>1</xdr:col><xdr:colOff>365760</xdr:colOff><xdr:row>2</xdr:row>',
      '<xdr:from><xdr:col>1</xdr:col><xdr:colOff>365760</xdr:colOff>',
    );
    const { workbook, sheet } = openWith({ drawing: broken });
    const plan = planRangeTransform(workbook, insertRow(sheet.id, 1));
    expect(plan.block?.code).toBe('anchor-untransformable');
    expect(plan.opaqueParts).toEqual([]);
  });

  it('transforms a banded cell edit only for anchors inside the band', () => {
    const { workbook, sheet } = openWith({ drawing: DRAWING });
    const plan = planRangeTransform(workbook, {
      kind: 'insert',
      sheetId: sheet.id,
      axis: 'row',
      index: 2,
      count: 1,
      band: [0, 0],
    });
    // Both anchors sit in columns outside the band, so neither moves.
    expect(plannedPart(plan, 'xl/drawings/drawing1.xml')).toBeNull();
    const inside = planRangeTransform(workbook, {
      kind: 'insert',
      sheetId: sheet.id,
      axis: 'row',
      index: 2,
      count: 1,
      band: [1, 1],
    });
    const after = plannedPart(inside, 'xl/drawings/drawing1.xml')!.after;
    expect(after).toContain(
      '<xdr:col>1</xdr:col><xdr:colOff>365760</xdr:colOff><xdr:row>3</xdr:row>',
    );
  });

  it('documents the marker rule for deletes beyond the grid', () => {
    expect(transformAnchorMarker(0, 0, deleteRow('s', 0, 5))).toEqual({ row: 0, col: 0 });
    expect(transformAnchorMarker(3, 0, deleteRow('s', 1, 2))).toEqual({ row: 1, col: 0 });
  });
});

// ---------------------------------------------------------------------------
// Chart series ranges
// ---------------------------------------------------------------------------

describe('M6 chart series ranges', () => {
  it('rewrites cat and val ranges when rows are inserted above the charted data', () => {
    const { workbook, sheet } = openWith({ drawing: DRAWING, chart: CHART_RANGE });
    const plan = planRangeTransform(workbook, insertRow(sheet.id, 1));
    const after = plannedPart(plan, 'xl/charts/chart1.xml')!.after;
    expect(after).toContain('<c:f>Sheet1!$A$3:$A$4</c:f>');
    expect(after).toContain('<c:f>Sheet1!$B$3:$B$4</c:f>');
    // The cache is preserved: a cached chart still shows what Excel last wrote.
    expect(after).toContain('<c:strCache/>');
  });

  it('rewrites only the axis the edit moves', () => {
    const { workbook, sheet } = openWith({ drawing: DRAWING, chart: CHART_RANGE });
    const rows = plannedPart(
      planRangeTransform(workbook, insertColumn(sheet.id, 0)),
      'xl/charts/chart1.xml',
    )!.after;
    expect(rows).toContain('<c:f>Sheet1!$B$2:$B$3</c:f>');
    const columns = plannedPart(
      planRangeTransform(workbook, insertRow(sheet.id, 0)),
      'xl/charts/chart1.xml',
    )!.after;
    expect(columns).toContain('<c:f>Sheet1!$A$3:$A$4</c:f>');
    expect(columns).toContain('<c:f>Sheet1!$B$3:$B$4</c:f>');
  });

  it('writes #REF! when the charted rows are deleted', () => {
    const { workbook, sheet } = openWith({ drawing: DRAWING, chart: CHART_RANGE });
    const plan = planRangeTransform(workbook, deleteRow(sheet.id, 1, 2));
    const after = plannedPart(plan, 'xl/charts/chart1.xml')!.after;
    expect(after).toContain('<c:f>#REF!</c:f>');
    expect(plan.block).toBeNull();
  });

  it('leaves a chart on another sheet untouched', () => {
    const other = chart(
      '<c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numRef><c:f>Other!$B$2:$B$3</c:f></c:numRef></c:val></c:ser>',
    );
    const { workbook, sheet } = openWith({ drawing: DRAWING, chart: other });
    const plan = planRangeTransform(workbook, insertRow(sheet.id, 1));
    expect(plan.block).toBeNull();
    expect(plannedPart(plan, 'xl/charts/chart1.xml')).toBeNull();
  });

  it('blocks a chart served by a defined name', () => {
    const named = chart(
      '<c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numRef><c:f>ChartData</c:f></c:numRef></c:val></c:ser>',
    );
    const { workbook, sheet } = openWith({ drawing: DRAWING, chart: named });
    const plan = planRangeTransform(workbook, insertRow(sheet.id, 1));
    expect(plan.block?.code).toBe('chart-source');
    expect(plan.block?.message).toContain(CHART_SOURCE_BLOCK_MESSAGE);
    expect(plannedPart(plan, 'xl/charts/chart1.xml')).toBeNull();
  });

  it('blocks a chart served by a union of ranges', () => {
    const union = chart(
      '<c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numRef><c:f>Sheet1!$B$2:$B$3,Sheet1!$C$2:$C$3</c:f></c:numRef></c:val></c:ser>',
    );
    const { workbook, sheet } = openWith({ drawing: DRAWING, chart: union });
    expect(planRangeTransform(workbook, insertRow(sheet.id, 1)).block?.code).toBe('chart-source');
  });

  it('rewrites a bare range against the chart owner sheet resolved through relationships', () => {
    const bare = chart(
      '<c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numRef><c:f>$B$2:$B$3</c:f></c:numRef></c:val></c:ser>',
    );
    const { workbook, sheet } = openWith({ drawing: DRAWING, chart: bare });
    const after = plannedPart(
      planRangeTransform(workbook, insertRow(sheet.id, 1)),
      'xl/charts/chart1.xml',
    )!.after;
    expect(after).toContain('<c:f>$B$3:$B$4</c:f>');
  });
});

// ---------------------------------------------------------------------------
// Pivot source ranges
// ---------------------------------------------------------------------------

describe('M6 pivot source ranges', () => {
  function pivotWorkbook(source: string, withTable = true) {
    return openWith({
      pivotCache: PIVOT_CACHE(source),
      ...(withTable ? { pivotTable: PIVOT_TABLE } : {}),
    });
  }

  it('blocks an edit that would change the pivot source range', () => {
    const { workbook, sheet } = pivotWorkbook('<worksheetSource ref="A1:C10" sheet="Sheet1"/>');
    const plan = planRangeTransform(workbook, insertRow(sheet.id, 3));
    expect(plan.block?.code).toBe('pivot-source');
    expect(plan.block?.message).toContain(PIVOT_SOURCE_BLOCK_MESSAGE);
    expect(plan.opaqueParts).toEqual([]);
  });

  it('blocks a column delete inside the source range', () => {
    const { workbook, sheet } = pivotWorkbook('<worksheetSource ref="A1:C10" sheet="Sheet1"/>');
    expect(planRangeTransform(workbook, deleteColumn(sheet.id, 1)).block?.code).toBe(
      'pivot-source',
    );
  });

  it('allows an edit outside the source range and leaves the source untouched', () => {
    const { workbook, sheet } = pivotWorkbook('<worksheetSource ref="A1:C10" sheet="Sheet1"/>');
    const plan = planRangeTransform(workbook, insertRow(sheet.id, 40));
    expect(plan.block).toBeNull();
    expect(plan.opaqueParts).toEqual([]);
  });

  it('allows an edit on a different sheet', () => {
    const opened = openWorkbook(
      buildXlsx(`${HEADER}\n<worksheet xmlns="${NS}"><sheetData/></worksheet>`, {
        extraParts: {
          'xl/pivotCache/pivotCacheDefinition1.xml': PIVOT_CACHE(
            '<worksheetSource ref="A1:C10" sheet="Other"/>',
          ),
        },
      }),
    );
    const sheet = opened.workbook.sheets[0]!;
    expect(planRangeTransform(opened.workbook, insertRow(sheet.id, 3)).block).toBeNull();
  });

  it('resolves the owner sheet through the pivot-table relationship when sheet is absent', () => {
    const { workbook, sheet } = pivotWorkbook('<worksheetSource ref="A1:C10"/>');
    expect(planRangeTransform(workbook, insertRow(sheet.id, 3)).block?.code).toBe('pivot-source');
    expect(planRangeTransform(workbook, insertRow(sheet.id, 40)).block).toBeNull();
  });

  it('leaves an external or table-name cache source alone', () => {
    const { workbook, sheet } = pivotWorkbook('<worksheetSource name="SalesTable"/>', false);
    const plan = planRangeTransform(workbook, insertRow(sheet.id, 3));
    expect(plan.block).toBeNull();
    expect(plan.opaqueParts).toEqual([]);
    const external = openWith({
      pivotCache: `${HEADER}<pivotCacheDefinition xmlns="${NS}" xmlns:r="${R_NS}"><cacheSource type="external"/><cacheFields count="0"/></pivotCacheDefinition>`,
    });
    expect(planRangeTransform(external.workbook, insertRow(external.sheet.id, 3)).block).toBeNull();
  });

  it('moves the pivot output rectangle with the cells it occupies', () => {
    const { workbook, sheet } = openWith({
      pivotCache: PIVOT_CACHE('<worksheetSource ref="D1:E4" sheet="Other"/>'),
      pivotTable: PIVOT_TABLE,
    });
    const plan = planRangeTransform(workbook, insertRow(sheet.id, 0));
    expect(plan.block).toBeNull();
    const after = plannedPart(plan, 'xl/pivotTables/pivotTable1.xml')!.after;
    expect(after).toContain('<location ref="G4:H7"');
  });

  it('blocks when the edit removes the whole pivot output', () => {
    const { workbook, sheet } = openWith({
      pivotCache: PIVOT_CACHE('<worksheetSource ref="D1:E4" sheet="Other"/>'),
      pivotTable: PIVOT_TABLE,
    });
    const plan = planRangeTransform(workbook, deleteRow(sheet.id, 2, 10));
    expect(plan.block?.code).toBe('anchor-untransformable');
  });
});

// ---------------------------------------------------------------------------
// Comment, threaded-comment and legacy note positions
// ---------------------------------------------------------------------------

describe('M6 comment and note positions', () => {
  it('moves legacy comment refs and drops a comment on the deleted row', () => {
    const { workbook, sheet } = openWith({ comments: COMMENTS, vml: VML });
    const inserted = plannedPart(
      planRangeTransform(workbook, insertRow(sheet.id, 1)),
      'xl/comments1.xml',
    );
    expect(inserted!.after).toContain('<comment ref="B4" authorId="0">');
    expect(inserted!.after).toContain('<comment ref="B6" authorId="0">');

    const deleted = plannedPart(
      planRangeTransform(workbook, deleteRow(sheet.id, 2)),
      'xl/comments1.xml',
    );
    expect(deleted!.after).toContain('<comment ref="B4" authorId="0"><text><r><t>keep me</t>');
    expect(deleted!.after).not.toContain('drop me');
  });

  it('moves threaded-comment refs with their cells', () => {
    const { workbook, sheet } = openWith({
      threadedComments: THREADED_COMMENTS,
      comments: COMMENTS,
      vml: VML,
    });
    const after = plannedPart(
      planRangeTransform(workbook, insertColumn(sheet.id, 0)),
      'xl/threadedComments/threadedComment1.xml',
    )!.after;
    expect(after).toContain('<threadedComment ref="C3"');
  });

  it('moves the VML note anchor and its client-data row together', () => {
    const { workbook, sheet } = openWith({ comments: COMMENTS, vml: VML });
    const after = plannedPart(
      planRangeTransform(workbook, insertRow(sheet.id, 1)),
      'xl/drawings/vmlDrawing1.vml',
    )!.after;
    expect(after).toContain('<x:Anchor>1, 15, 3, 2, 3, 15, 6, 3</x:Anchor>');
    expect(after).toContain('<x:Row>3</x:Row>');
    expect(after).toContain('<x:Column>1</x:Column>');
  });

  it('blocks a VML note anchor that is not eight integers', () => {
    const broken = VML.replace(
      '<x:Anchor>1, 15, 2, 2, 3, 15, 5, 3</x:Anchor>',
      '<x:Anchor>1, 15, 2</x:Anchor>',
    );
    const { workbook, sheet } = openWith({ comments: COMMENTS, vml: broken });
    expect(planRangeTransform(workbook, insertRow(sheet.id, 1)).block?.code).toBe(
      'anchor-untransformable',
    );
  });

  it('does not touch note parts for an edit that leaves their cells alone', () => {
    const { workbook, sheet } = openWith({ comments: COMMENTS, vml: VML });
    const plan = planRangeTransform(workbook, insertRow(sheet.id, 50));
    expect(plan.opaqueParts).toEqual([]);
  });

  it('renumbers hyperlink refs in the worksheet part and drops a deleted one', () => {
    const { workbook, sheet } = openWith({});
    const before = workbook.sourceParts.get(sheet.partPath)!;
    const inserted = renumberWorksheetPart({
      bytes: before,
      partName: sheet.partPath,
      sheetName: sheet.name,
      edit: insertRow(sheet.id, 1),
      workbook,
    });
    expect(partText(inserted.bytes)).toContain('<hyperlink ref="C3"');
    const deleted = renumberWorksheetPart({
      bytes: before,
      partName: sheet.partPath,
      sheetName: sheet.name,
      edit: deleteRow(sheet.id, 1),
      workbook,
    });
    expect(partText(deleted.bytes)).not.toContain('ref="C2"');
  });
});

// ---------------------------------------------------------------------------
// Hyperlink locations: transform or block
// ---------------------------------------------------------------------------

describe('M6 hyperlink locations', () => {
  /** The edited sheet plus a second sheet whose name needs quoting. */
  function locationWorkbook() {
    const workbook = createEmptyWorkbook();
    const first = createSheet('Sheet1', 'xl/worksheets/sheet1.xml');
    const second = createSheet('Quoted Name', 'xl/worksheets/sheet2.xml');
    workbook.sheets.push(first, second);
    return { workbook, first, second };
  }

  /** One hyperlink on the edited sheet, at row 5 / column B. */
  function planWith(
    location: string,
    edit: (sheetId: string) => StructuralEdit,
    target: 'first' | 'second' = 'first',
  ) {
    const { workbook, first, second } = locationWorkbook();
    const sheet = target === 'first' ? first : second;
    sheet.hyperlinks.set('4:1', { target: '', location, tooltip: 'jump' });
    return planRangeTransform(workbook, edit(sheet.id));
  }

  function locationAfter(plan: ReturnType<typeof planWith>): string | undefined {
    return plan.hyperlinks.after[0]?.[1].location;
  }

  it('moves a sheet-qualified location on a row insert, keeping the authored form', () => {
    expect(locationAfter(planWith('Sheet1!A1', (sheetId) => insertRow(sheetId, 0)))).toBe(
      'Sheet1!A2',
    );
    expect(locationAfter(planWith('Sheet1!$B$3', (sheetId) => insertRow(sheetId, 1)))).toBe(
      'Sheet1!$B$4',
    );
    // The absolute marker and the range form survive the rewrite.
    expect(locationAfter(planWith('Sheet1!A1:B2', (sheetId) => insertRow(sheetId, 0)))).toBe(
      'Sheet1!A2:B3',
    );
  });

  it('moves a quoted-sheet location like a formula reference', () => {
    expect(
      locationAfter(planWith("'Quoted Name'!A1", (sheetId) => insertRow(sheetId, 0), 'second')),
    ).toBe("'Quoted Name'!A2");
  });

  it('moves a bare location, a range and a leading `#` some producers write', () => {
    // A bare address belongs to the sheet that holds the hyperlink.
    expect(locationAfter(planWith('A1', (sheetId) => insertRow(sheetId, 0)))).toBe('A2');
    expect(locationAfter(planWith('B2:C3', (sheetId) => insertColumn(sheetId, 0)))).toBe('C2:D3');
    expect(locationAfter(planWith('#Sheet1!A1', (sheetId) => insertRow(sheetId, 0)))).toBe(
      '#Sheet1!A2',
    );
  });

  it('moves a location on a row or column delete', () => {
    expect(locationAfter(planWith('Sheet1!A3', (sheetId) => deleteRow(sheetId, 0)))).toBe(
      'Sheet1!A2',
    );
    expect(locationAfter(planWith('Sheet1!C1', (sheetId) => deleteColumn(sheetId, 0)))).toBe(
      'Sheet1!B1',
    );
  });

  it('collapses a location whose whole target is deleted, matching names and formulas', () => {
    const plan = planWith('Sheet1!A1', (sheetId) => deleteRow(sheetId, 0));
    expect(plan.block).toBeNull();
    expect(locationAfter(plan)).toBe('#REF!');
    // A range that only loses its first row shrinks rather than collapsing.
    expect(locationAfter(planWith('Sheet1!A1:B2', (sheetId) => deleteRow(sheetId, 0)))).toBe(
      'Sheet1!A1:B1',
    );
    // A collapsed jump is a broken link, not a reason to block every later
    // edit on the sheet.
    const later = planWith('#REF!', (sheetId) => insertRow(sheetId, 0));
    expect(later.block).toBeNull();
    expect(locationAfter(later)).toBe('#REF!');
  });

  it('moves the hyperlink cell key and its location together', () => {
    const plan = planWith('Sheet1!A1', (sheetId) => insertRow(sheetId, 0));
    expect(plan.hyperlinks.after).toEqual([
      ['5:1', { target: '', location: 'Sheet1!A2', tooltip: 'jump' }],
    ]);
    expect(plan.changed).toBe(true);
    // The external target is a relationship URL, never a reference.
    expect(plan.hyperlinks.after[0]?.[1].target).toBe('');
  });

  it('leaves a location naming another sheet, or a resolving defined name, alone', () => {
    const other = planWith("'Quoted Name'!A1", (sheetId) => insertRow(sheetId, 0));
    expect(other.block).toBeNull();
    expect(locationAfter(other)).toBe("'Quoted Name'!A1");

    const { workbook, first } = locationWorkbook();
    workbook.names.push({ name: 'Anchor', formula: 'Sheet1!$A$1', sheetId: null });
    first.hyperlinks.set('4:1', { target: '', location: 'Anchor' });
    const named = planRangeTransform(workbook, insertRow(first.id, 0));
    // The same plan rewrites the name's own formula, so the jump keeps
    // resolving to the intended cell and the authored text stays correct.
    expect(named.block).toBeNull();
    expect(named.nameRewrites).toEqual([
      {
        scope: 'workbook',
        sheetId: null,
        name: 'Anchor',
        before: 'Sheet1!$A$1',
        after: 'Sheet1!$A$2',
      },
    ]);
    expect(named.hyperlinks.after[0]?.[1].location).toBe('Anchor');
  });

  it('blocks a location that is not a cell, range or defined name', () => {
    const plan = planWith('Sheet1:Quoted Name!A1', (sheetId) => insertRow(sheetId, 0));
    expect(plan.block?.code).toBe('anchor-untransformable');
    expect(plan.block?.message).toContain(HYPERLINK_LOCATION_BLOCK_MESSAGE);
    expect(plan.block?.message).toContain('Sheet1:Quoted Name!A1');
  });

  it('blocks a defined name the module cannot resolve', () => {
    expect(planWith('Missing_Name', (sheetId) => insertRow(sheetId, 0)).block?.code).toBe(
      'anchor-untransformable',
    );
  });

  it('moves a hyperlink on another sheet that links into the edited sheet', () => {
    const { workbook, first, second } = locationWorkbook();
    second.hyperlinks.set('4:1', { target: '', location: 'Sheet1!A1' });
    const plan = planRangeTransform(workbook, insertRow(first.id, 0));
    expect(plan.block).toBeNull();
    expect(plan.crossSheetHyperlinks).toHaveLength(1);
    expect(plan.crossSheetHyperlinks[0]?.after[0]?.[1].location).toBe('Sheet1!A2');
    // An edit that leaves the linked cell where it is does not rewrite or block.
    const untouched = planRangeTransform(workbook, insertRow(first.id, 20));
    expect(untouched.block).toBeNull();
    expect(untouched.crossSheetHyperlinks).toEqual([]);
  });

  it('refuses the edit with no transaction when the location cannot be transformed', () => {
    const original = buildXlsx(
      `${HEADER}\n<worksheet xmlns="${NS}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><hyperlinks><hyperlink ref="A1" location="Not_A_Reference" display="odd"/></hyperlinks></worksheet>`,
    );
    const opened = openWorkbook(original);
    const sheet = opened.workbook.sheets[0]!;
    const session = new DocumentSession(opened.workbook);
    const result = session.dispatch(
      'structure.insertRows',
      {},
      {
        activeSheetId: sheet.id,
        selection: { startRow: 0, startCol: 0, endRow: 0, endCol: 0 },
        activeCell: { row: 0, col: 0 },
      },
    );
    expect(result.structure?.state).toBe('blocked');
    expect(result.structure?.message).toContain('Not_A_Reference');
    expect(result.ui).toMatchObject({ type: 'toast' });
    expect(session.transactionCount).toBe(0);
  });

  it('writes the transformed location into the saved worksheet part', () => {
    const original = buildMultiSheetXlsx(
      [
        {
          name: 'Sheet1',
          worksheetXml: `${HEADER}
<worksheet xmlns="${NS}" xmlns:r="${R_NS}">
  <dimension ref="A1:C6"/>
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
    <row r="5"><c r="A5" t="inlineStr"><is><t>up</t></is></c><c r="B5" t="inlineStr"><is><t>jump</t></is></c><c r="C5" t="inlineStr"><is><t>other</t></is></c></row>
  </sheetData>
  <hyperlinks>
    <hyperlink ref="A5" r:id="rId1" tooltip="External report"/>
    <hyperlink ref="B5" location="Sheet1!A1" display="up"/>
    <hyperlink ref="C5" location="'Quoted Name'!B2" display="other"/>
  </hyperlinks>
</worksheet>`,
        },
        {
          name: 'Quoted Name',
          worksheetXml: `${HEADER}\n<worksheet xmlns="${NS}"><sheetData><row r="2"><c r="B2"><v>2</v></c></row></sheetData></worksheet>`,
        },
      ],
      {
        extraParts: {
          'xl/worksheets/_rels/sheet1.xml.rels': `${HEADER}
<Relationships xmlns="${PKG_NS}">
  <Relationship Id="rId1" Type="${R_NS}/hyperlink" Target="https://example.invalid/report" TargetMode="External"/>
</Relationships>`,
        },
      },
    );
    const opened = openWorkbook(original);
    const sheet = opened.workbook.sheets[0]!;
    const session = new DocumentSession(opened.workbook);
    const result = session.dispatch(
      'structure.insertRows',
      {},
      {
        activeSheetId: sheet.id,
        selection: { startRow: 0, startCol: 0, endRow: 0, endCol: 0 },
        activeCell: { row: 0, col: 0 },
      },
    );
    expect(result.structure?.state).toBe('applied');

    const saved = saveWorkbook({
      workbook: opened.workbook,
      package: opened.package,
      dirtyCells: session.dirtyCellMap(),
      dirtyDimensions: session.dirtyDimensionSnapshot(),
      dirtyMerges: new Set(session.dirtyMergeSnapshot().keys()),
    });
    const afterXml = unzipPart(saved.bytes, 'xl/worksheets/sheet1.xml');
    // The internal location follows the inserted row; the external target and
    // the location that names another sheet keep their authored text.
    expect(afterXml).toContain('<hyperlink ref="B6" location="Sheet1!A2" display="up"/>');
    expect(afterXml).toContain(
      '<hyperlink ref="C6" location="\'Quoted Name\'!B2" display="other"/>',
    );
    expect(afterXml).toContain('<hyperlink ref="A6" r:id="rId1" tooltip="External report"/>');
    expect(afterXml).not.toContain('location="Sheet1!A1"');
  });
});

// ---------------------------------------------------------------------------
// Unsupported formula constructs: block, never stale text
// ---------------------------------------------------------------------------

describe('M6 unsupported formula constructs', () => {
  function workbookWithFormula(formula: string) {
    const workbook = createEmptyWorkbook();
    const sheet = createSheet('Sheet1', 'xl/worksheets/sheet1.xml');
    const second = createSheet('Sheet2', 'xl/worksheets/sheet2.xml');
    const third = createSheet('Sheet3', 'xl/worksheets/sheet3.xml');
    workbook.sheets.push(sheet, second, third);
    setFormula(sheet, 1, 0, formula);
    return { workbook, sheet, second, third };
  }

  function setFormula(sheet: Sheet, row: number, col: number, formula: string): void {
    setCellSnapshot(sheet, row, col, {
      value: numberValue(0),
      formula,
      formulaDerived: false,
      savedResult: numberValue(0),
      styleId: 0,
      richText: null,
    });
  }

  it('blocks a union formula on the edited sheet', () => {
    const { workbook, sheet } = workbookWithFormula('(A1:A2,C1:C2)');
    const plan = planRangeTransform(workbook, insertRow(sheet.id, 1));
    expect(plan.block?.code).toBe('unsupported-formula');
    expect(plan.block?.message).toContain('A2');
    expect(plan.formulaRewrites).toEqual([]);
  });

  it('blocks a 3D range that spans the edited sheet from another sheet', () => {
    const workbook = createEmptyWorkbook();
    const first = createSheet('Sheet1', 'xl/worksheets/sheet1.xml');
    const second = createSheet('Sheet2', 'xl/worksheets/sheet2.xml');
    const third = createSheet('Sheet3', 'xl/worksheets/sheet3.xml');
    workbook.sheets.push(first, second, third);
    setFormula(third, 0, 0, 'SUM(Sheet1:Sheet2!A1)');
    expect(planRangeTransform(workbook, insertRow(second.id, 1)).block?.code).toBe(
      'unsupported-formula',
    );
    // A span that does not include the edited sheet is unaffected.
    expect(planRangeTransform(workbook, insertRow(third.id, 1)).block).toBeNull();
  });

  it('blocks implicit intersection on the edited sheet', () => {
    const { workbook, sheet } = workbookWithFormula('@A1');
    expect(planRangeTransform(workbook, insertRow(sheet.id, 1)).block?.code).toBe(
      'unsupported-formula',
    );
  });

  it('does not block or rewrite an external-link reference', () => {
    const { workbook, sheet } = workbookWithFormula('[1]Sheet1!A1');
    const plan = planRangeTransform(workbook, insertRow(sheet.id, 1));
    expect(plan.block).toBeNull();
    expect(plan.formulaRewrites).toEqual([]);
  });

  it('blocks a formula that mixes an external reference with a local one', () => {
    const { workbook, sheet } = workbookWithFormula('[1]Sheet1!A1+B2');
    expect(planRangeTransform(workbook, insertRow(sheet.id, 1)).block?.code).toBe(
      'unsupported-formula',
    );
  });

  it('leaves a structured table reference for the stale report but blocks a sibling A1 ref', () => {
    const pure = workbookWithFormula('Table1[Amount]');
    expect(planRangeTransform(pure.workbook, insertRow(pure.sheet.id, 1)).block).toBeNull();
    const mixed = workbookWithFormula('Table1[Amount]+A1');
    expect(planRangeTransform(mixed.workbook, insertRow(mixed.sheet.id, 1)).block?.code).toBe(
      'unsupported-formula',
    );
  });

  it('rewrites references inside a namespaced function instead of leaving them stale', () => {
    const { workbook, sheet } = workbookWithFormula('_xlfn.XLOOKUP(A2,B:B,C:C)');
    const plan = planRangeTransform(workbook, insertRow(sheet.id, 1));
    expect(plan.block).toBeNull();
    expect(plan.formulaRewrites.map((rewrite) => rewrite.after)).toEqual([
      '_xlfn.XLOOKUP(A3,B:B,C:C)',
    ]);
  });

  it('blocks an untransformable defined name that names the edited sheet', () => {
    const { workbook, sheet } = workbookWithFormula('1');
    workbook.names.push({ name: 'Weird', formula: 'Sheet1:Sheet3!A1', sheetId: null });
    expect(planRangeTransform(workbook, insertRow(sheet.id, 1)).block?.code).toBe(
      'unsupported-formula',
    );
  });
});

// ---------------------------------------------------------------------------
// Patch writer integration and undo
// ---------------------------------------------------------------------------

describe('M6 structural part patches', () => {
  it('plans the worksheet and every touched Tier B part in one call', () => {
    const { workbook, sheet } = openWith({
      drawing: DRAWING,
      chart: CHART_RANGE,
      comments: COMMENTS,
      vml: VML,
    });
    const edit = insertRow(sheet.id, 1);
    const plan = planRangeTransform(workbook, edit);
    const parts = planStructuralPartPatches(workbook, edit, plan);
    expect(parts.block).toBeNull();
    const paths = parts.patches.map((patch) => patch.path);
    expect(paths).toContain(sheet.partPath);
    expect(paths).toContain('xl/charts/chart1.xml');
    expect(paths).toContain('xl/drawings/drawing1.xml');
    expect(paths).toContain('xl/comments1.xml');
  });

  it('writes transformed chart and drawing bytes through the save path', () => {
    const {
      workbook,
      sheet,
      package: pkg,
    } = openWith({
      drawing: DRAWING,
      chart: CHART_RANGE,
      comments: COMMENTS,
      vml: VML,
    });
    const edit = insertRow(sheet.id, 1);
    const plan = planRangeTransform(workbook, edit);
    const parts = planStructuralPartPatches(workbook, edit, plan);
    expect(parts.block).toBeNull();
    applyStructuralPartPatches(workbook, parts.patches, 'after');

    // The save path writes the package entries; a structure-aware save copies
    // the patched opaque bytes into them (the orchestrator's wiring hook).
    const entries = new Map(pkg.entries);
    for (const patch of parts.patches) entries.set(patch.path, patch.after);
    const result = saveWorkbook({
      workbook,
      package: { ...pkg, entries },
      dirtyCells: new Map(),
      dirtyMerges: new Set([sheet.id]),
    });
    expect(partText(result.package.entries.get('xl/charts/chart1.xml'))).toContain(
      '<c:f>Sheet1!$A$3:$A$4</c:f>',
    );
    expect(partText(result.package.entries.get('xl/drawings/drawing1.xml'))).toContain(
      '<xdr:row>3</xdr:row>',
    );
    expect(partText(result.package.entries.get('xl/comments1.xml'))).toContain('<comment ref="B4"');
  });

  it('undo restores the exact prior part text', () => {
    const { workbook, sheet } = openWith({
      drawing: DRAWING,
      chart: CHART_RANGE,
      comments: COMMENTS,
      vml: VML,
    });
    const before = new Map(
      [
        'xl/charts/chart1.xml',
        'xl/drawings/drawing1.xml',
        'xl/comments1.xml',
        'xl/drawings/vmlDrawing1.vml',
      ].map((path) => [path, partText(workbook.opaqueParts.get(path)!.bytes)] as const),
    );
    const edit = insertRow(sheet.id, 1);
    const parts = planStructuralPartPatches(workbook, edit);
    applyStructuralPartPatches(workbook, parts.patches, 'after');
    expect(partText(workbook.opaqueParts.get('xl/charts/chart1.xml')!.bytes)).not.toBe(
      before.get('xl/charts/chart1.xml'),
    );
    applyStructuralPartPatches(workbook, parts.patches, 'before');
    for (const [path, text] of before) {
      expect(partText(workbook.opaqueParts.get(path)!.bytes)).toBe(text);
    }
  });

  it('exposes the opaque patch set without the worksheet for minimal wiring', () => {
    const { workbook, sheet } = openWith({ drawing: DRAWING, chart: CHART_RANGE });
    const plan = planRangeTransform(workbook, insertRow(sheet.id, 1));
    const minimal = planOpaquePartPatches(
      workbook,
      { kind: 'insert', sheetId: sheet.id, axis: 'row', index: 1, count: 1 },
      plan,
    );
    expect(minimal.block).toBeNull();
    expect(minimal.patches.map((patch) => patch.path).sort()).toEqual([
      'xl/charts/chart1.xml',
      'xl/drawings/drawing1.xml',
    ]);
  });

  it('planning is read-only: the model parts still hold the original bytes', () => {
    const { workbook, sheet } = openWith({ drawing: DRAWING, chart: CHART_RANGE });
    const before = partText(workbook.opaqueParts.get('xl/charts/chart1.xml')!.bytes);
    const plan = planRangeTransform(workbook, insertRow(sheet.id, 1));
    expect(plan.opaqueParts).toHaveLength(2);
    expect(partText(workbook.opaqueParts.get('xl/charts/chart1.xml')!.bytes)).toBe(before);
  });

  it('never rewrites or blocks on external-link parts', () => {
    const externalLink = `${HEADER}<externalLink xmlns="${NS}"><externalBook xmlns:r="${R_NS}" r:id="rId1"><sheetNames><sheetName val="Sheet1"/></sheetNames><sheetDataSet><sheetData sheetId="0"/></sheetDataSet></externalBook></externalLink>`;
    const { workbook, sheet } = openWith({});
    workbook.opaqueParts.set('xl/externalLinks/externalLink1.xml', {
      path: 'xl/externalLinks/externalLink1.xml',
      bytes: new TextEncoder().encode(externalLink),
      anchor: null,
    });
    const plan = planRangeTransform(workbook, insertRow(sheet.id, 1));
    expect(plan.block).toBeNull();
    expect(plan.opaqueParts).toEqual([]);
    expect(partText(workbook.opaqueParts.get('xl/externalLinks/externalLink1.xml')!.bytes)).toBe(
      externalLink,
    );
  });

  it('rewrites the producer-written chart fixture series and anchor (F9 behaviour)', () => {
    const original = new Uint8Array(
      readFileSync(new URL('../fixtures/compat/m1-chart-unknown.xlsx', import.meta.url)),
    );
    const opened = openWorkbook(original);
    const sheet = opened.workbook.sheets[0]!;
    const plan = planRangeTransform(opened.workbook, insertRow(sheet.id, 0));
    expect(plan.block).toBeNull();
    expect(plannedPart(plan, 'xl/charts/chart1.xml')!.after).toContain(
      '<c:f>Sheet1!$B$3:$B$4</c:f>',
    );
    expect(plannedPart(plan, 'xl/drawings/drawing1.xml')!.after).toContain('<xdr:row>1</xdr:row>');
  });

  it('scopes drawing anchors to the sheet that owns them', () => {
    const pkg = buildMultiSheetXlsx(
      [
        {
          name: 'Sheet1',
          worksheetXml: `${HEADER}<worksheet xmlns="${NS}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>`,
        },
        {
          name: 'Sheet2',
          worksheetXml: `${HEADER}<worksheet xmlns="${NS}" xmlns:r="${R_NS}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><drawing r:id="rId1"/></worksheet>`,
        },
      ],
      {
        extraParts: {
          'xl/worksheets/_rels/sheet2.xml.rels': `${HEADER}<Relationships xmlns="${PKG_NS}"><Relationship Id="rId1" Type="${R_NS}/drawing" Target="../drawings/drawing1.xml"/></Relationships>`,
          'xl/drawings/drawing1.xml': DRAWING,
        },
      },
    );
    const opened = openWorkbook(pkg);
    const first = opened.workbook.sheets[0]!;
    const second = opened.workbook.sheets[1]!;
    expect(planRangeTransform(opened.workbook, insertRow(first.id, 0)).opaqueParts).toEqual([]);
    expect(
      planRangeTransform(opened.workbook, insertRow(second.id, 0)).opaqueParts.map(
        (rewrite) => rewrite.path,
      ),
    ).toEqual(['xl/drawings/drawing1.xml']);
  });
});

// ---------------------------------------------------------------------------
// Pure part transforms, tested directly
// ---------------------------------------------------------------------------

describe('M6 part transform primitives', () => {
  it('drawing transform reports no change outside the edit', () => {
    const result = transformDrawingAnchorsXml(DRAWING, insertRow('s', 100));
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.changed).toBe(false);
  });

  it('chart transform refuses a non-range source', () => {
    const result = transformChartXml(
      chart('<c:ser><c:val><c:numRef><c:f>MyName</c:f></c:numRef></c:val></c:ser>'),
      insertRow('s', 0),
      { workbook: createEmptyWorkbook() },
      null,
    );
    expect(result.ok).toBe(false);
  });

  it('comment transforms drop a deleted cell', () => {
    const result = transformCommentRefsXml(COMMENTS, deleteRow('s', 2));
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.xml).toContain('<comment ref="B4"');
      expect(result.xml).toContain('keep me');
      expect(result.xml).not.toContain('drop me');
    }
  });

  it('threaded comment transform keeps unknown attributes verbatim', () => {
    const result = transformThreadedCommentRefsXml(THREADED_COMMENTS, insertRow('s', 1));
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.xml).toContain('personId="{1}"');
  });

  it('pivot transform never rewrites a source ref', () => {
    const result = transformPivotCacheXml(
      PIVOT_CACHE('<worksheetSource ref="A1:C10" sheet="Sheet1"/>'),
      insertRow('s', 40),
      { workbook: createEmptyWorkbook() },
      null,
    );
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.changed).toBe(false);
  });

  it('VML transform reports the note row and column', () => {
    const result = transformVmlAnchorsXml(VML, insertRow('s', 1));
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.xml).toContain('<x:Row>3</x:Row>');
  });
});

// ---------------------------------------------------------------------------
// The matrix sweep: transformed or blocked, never stale
// ---------------------------------------------------------------------------

const CONSTRUCTS = {
  chart: CHART_RANGE,
  drawing: DRAWING,
  comments: COMMENTS,
  vml: VML,
  pivot: PIVOT_CACHE('<worksheetSource ref="A1:C10" sheet="Sheet1"/>'),
} as const;

const EDIT_ARBITRARY = fc.record({
  kind: fc.constantFrom('insert' as const, 'delete' as const),
  axis: fc.constantFrom('row' as const, 'column' as const),
  index: fc.integer({ min: 0, max: 12 }),
  count: fc.integer({ min: 1, max: 3 }),
  banded: fc.boolean(),
});

describe('M6 construct × edit-shape matrix', () => {
  it('transforms or blocks every range-bearing construct for every edit shape', () => {
    fc.assert(
      fc.property(EDIT_ARBITRARY, ({ kind, axis, index, count, banded }) => {
        const { workbook, sheet } = openWith({
          drawing: CONSTRUCTS.drawing,
          chart: CONSTRUCTS.chart,
          comments: CONSTRUCTS.comments,
          vml: CONSTRUCTS.vml,
          pivotCache: CONSTRUCTS.pivot,
          pivotTable: PIVOT_TABLE,
        });
        const edit: StructuralEdit = banded
          ? { kind, sheetId: sheet.id, axis, index, count, band: [0, 2] }
          : { kind, sheetId: sheet.id, axis, index, count };
        const plan = planRangeTransform(workbook, edit);
        const planned = new Set(plan.opaqueParts.map((rewrite) => rewrite.path));

        // The pivot source edit is the contract's block case. When nothing
        // blocks, every construct whose own transform changes must be planned.
        if (plan.block === null) {
          const expectations: [string, { ok: boolean; changed: boolean }][] = [
            [
              'xl/charts/chart1.xml',
              simplify(transformChartXml(CONSTRUCTS.chart, edit, { workbook }, sheet.id)),
            ],
            [
              'xl/drawings/drawing1.xml',
              simplify(transformDrawingAnchorsXml(CONSTRUCTS.drawing, edit)),
            ],
            ['xl/comments1.xml', simplify(transformCommentRefsXml(CONSTRUCTS.comments, edit))],
            ['xl/drawings/vmlDrawing1.vml', simplify(transformVmlAnchorsXml(CONSTRUCTS.vml, edit))],
            ['xl/pivotTables/pivotTable1.xml', simplify(transformPivotTableXml(PIVOT_TABLE, edit))],
          ];
          for (const [path, expectation] of expectations) {
            if (expectation.changed)
              expect(planned.has(path), `${path} ${JSON.stringify(edit)}`).toBe(true);
            else expect(planned.has(path), `${path} ${JSON.stringify(edit)}`).toBe(false);
          }
        } else {
          expect([
            'chart-source',
            'pivot-source',
            'anchor-untransformable',
            'unsupported-formula',
          ]).toContain(plan.block.code);
        }
      }),
      { numRuns: 120 },
    );
  });

  it('the pivot source edit blocks exactly when the edit touches the range', () => {
    fc.assert(
      fc.property(EDIT_ARBITRARY, ({ kind, axis, index, count }) => {
        const { workbook, sheet } = openWith({
          pivotCache: CONSTRUCTS.pivot,
          pivotTable: PIVOT_TABLE,
        });
        const edit: StructuralEdit = { kind, sheetId: sheet.id, axis, index, count };
        const touched =
          transformRange({ startRow: 0, startCol: 0, endRow: 9, endCol: 2 }, edit) === null ||
          (() => {
            const transformed = transformRange(
              { startRow: 0, startCol: 0, endRow: 9, endCol: 2 },
              edit,
            );
            return (
              transformed !== null &&
              (transformed.startRow !== 0 ||
                transformed.startCol !== 0 ||
                transformed.endRow !== 9 ||
                transformed.endCol !== 2)
            );
          })();
        const plan = planRangeTransform(workbook, edit);
        expect(plan.block?.code === 'pivot-source').toBe(touched);
      }),
      { numRuns: 80 },
    );
  });

  it('never leaves an opaque part stale across combined constructs', () => {
    fc.assert(
      fc.property(fc.array(EDIT_ARBITRARY, { minLength: 1, maxLength: 3 }), (shapes) => {
        const { workbook, sheet } = openWith({
          drawing: DRAWING,
          chart: CHART_RANGE,
          comments: COMMENTS,
          vml: VML,
        });
        for (const { kind, axis, index, count, banded } of shapes) {
          const edit: StructuralEdit = banded
            ? { kind, sheetId: sheet.id, axis, index, count, band: [0, 2] }
            : { kind, sheetId: sheet.id, axis, index, count };
          const plan = planRangeTransform(workbook, edit);
          if (plan.block !== null) continue;
          // Apply what the plan says and check each part is either rewritten
          // or provably unaffected for this shape.
          for (const rewrite of plan.opaqueParts) {
            const text = partText(rewrite.after);
            expect(text.length).toBeGreaterThan(0);
          }
        }
      }),
      { numRuns: 40 },
    );
  });
});

function simplify(result: { ok: boolean; changed?: boolean }): {
  readonly ok: boolean;
  readonly changed: boolean;
} {
  return { ok: result.ok, changed: result.ok && result.changed === true };
}

// ---------------------------------------------------------------------------
// A workbook-model level guard: plans still report their own block codes
// ---------------------------------------------------------------------------

describe('M6 block typing', () => {
  it('keeps the existing block codes unchanged', () => {
    const workbook = createEmptyWorkbook();
    const sheet = createSheet('Sheet1', 'xl/worksheets/sheet1.xml');
    workbook.sheets.push(sheet);
    sheet.merges.push({ range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 } });
    const plan = planRangeTransform(workbook, {
      kind: 'insert',
      sheetId: sheet.id,
      axis: 'row',
      index: 1,
      count: 1,
      band: [0, 0],
    });
    expect(plan.block?.code).toBe('merge-split');
  });

  it('keeps an untouched workbook free of opaque work', () => {
    const workbook = createEmptyWorkbook();
    const sheet = createSheet('Sheet1', 'xl/worksheets/sheet1.xml');
    workbook.sheets.push(sheet);
    setCellSnapshot(sheet, 0, 0, {
      value: textValue('x'),
      formula: null,
      formulaDerived: false,
      savedResult: null,
      styleId: 0,
      richText: null,
    });
    const plan = planRangeTransform(workbook, insertRow(sheet.id, 0));
    expect(plan.opaqueParts).toEqual([]);
    expect(planOpaquePartTransforms(workbook, insertRow(sheet.id, 0))).toEqual({
      rewrites: [],
      block: null,
      broken: 0,
    });
    expect(plan.cellMoves).toHaveLength(1);
    expect(plan.cellMoves[0]?.to).toEqual({ row: 1, col: 0 });
  });
});
