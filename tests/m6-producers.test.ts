import { execFileSync } from 'node:child_process';
import { existsSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { openWorkbook } from '@ooxml/open';
import { verifyPackage } from '@ooxml/verify-package';
import { parseCsv, serialiseCsv } from '@io/csv';

/**
 * M6 lane 6 — F22 producer matrix and F24 CSV fixture (file-compat §9).
 *
 * Every producer fixture that exists must open through `openWorkbook` and pass
 * the local repair-free verifier. Excel 365, Excel 2019 and Google Sheets are
 * maintainer captures; they are never synthesised here, and the group's
 * `--status` output is the machine-readable record of what this machine could
 * and could not generate.
 */

const root = fileURLToPath(new URL('..', import.meta.url));
const producersDir = join(root, 'fixtures', 'compat', 'producers');
const manifestPath = join(producersDir, 'manifest.json');
const groupScript = join(root, 'scripts', 'fixtures', 'groups', 'producers.mjs');

interface ManifestProducer {
  readonly id: string;
  readonly file: string;
  readonly kind: 'script' | 'maintainer-capture';
  readonly status: string;
  readonly tool?: string;
  readonly note: string;
}

interface Manifest {
  readonly clause: string;
  readonly producers: readonly ManifestProducer[];
}

function readManifest(): Manifest {
  return JSON.parse(readFileSync(manifestPath, 'utf8')) as Manifest;
}

interface StatusEntry {
  readonly id: string;
  readonly available: boolean;
  readonly reason: string;
}

function producerStatus(): StatusEntry[] {
  const output = execFileSync(process.execPath, [groupScript, '--status'], {
    encoding: 'utf8',
    timeout: 120_000,
  });
  return (JSON.parse(output) as { producers: StatusEntry[] }).producers;
}

describe('F22 producer matrix', () => {
  it('declares every producer and marks the manual captures as excluded from generation', () => {
    const manifest = readManifest();
    expect(manifest.clause).toBe('file-compat §9 F22 producer matrix');
    const ids = manifest.producers.map((producer) => producer.id);
    expect(ids).toEqual([
      'excel-365',
      'excel-2019',
      'google-sheets',
      'libreoffice',
      'openpyxl',
      'sheetjs',
      'xlsx-populate',
    ]);
    for (const id of ['excel-365', 'excel-2019', 'google-sheets']) {
      const producer = manifest.producers.find((candidate) => candidate.id === id)!;
      expect(producer.kind).toBe('maintainer-capture');
      // Either the maintainer has committed the capture, or the manifest says
      // it is absent. The generator never writes it.
      if (!existsSync(join(producersDir, producer.file))) {
        expect(producer.status).toBe('absent');
      }
    }
  });

  it('reports availability per producer without fabricating a missing capture', () => {
    const status = producerStatus();
    const manifest = readManifest();
    for (const entry of status) {
      const producer = manifest.producers.find((candidate) => candidate.id === entry.id);
      expect(producer, `producer ${entry.id} is not in the manifest`).toBeDefined();
      if (entry.available) {
        expect(
          existsSync(join(producersDir, producer!.file)),
          `${entry.id} is available but ${producer!.file} is missing`,
        ).toBe(true);
      }
    }
    for (const id of ['excel-365', 'excel-2019', 'google-sheets']) {
      const entry = status.find((candidate) => candidate.id === id)!;
      expect(entry.available).toBe(false);
      expect(entry.reason).toMatch(/maintainer|manual/i);
    }
  });

  it('opens every committed fixture and passes the local repair-free verifier', () => {
    const manifest = readManifest();
    let checked = 0;
    for (const producer of manifest.producers) {
      const path = join(producersDir, producer.file);
      if (!existsSync(path)) continue;
      const bytes = new Uint8Array(readFileSync(path));
      const result = verifyPackage(bytes, producer.file);
      expect(result.problems, `${producer.file}: ${JSON.stringify(result.problems)}`).toEqual([]);
      const opened = openWorkbook(bytes, producer.file);
      expect(opened.workbook.sheets.length).toBeGreaterThan(0);
      checked += 1;
    }
    expect(checked).toBeGreaterThan(0);
  });
});

describe('F24 csv-roundtrip.csv', () => {
  it('is committed, parses into the authored rows and round-trips byte-for-byte', () => {
    const path = join(root, 'fixtures', 'compat', 'csv-roundtrip.csv');
    const text = readFileSync(path, 'utf8');
    const rows = parseCsv(text);
    expect(rows).toHaveLength(5);
    expect(rows[0]).toEqual(['Product', 'Region', 'Units', 'Unit price', 'Total', 'Note']);
    // Displayed values with thousands separators stay quoted.
    expect(rows[2]![4]).toBe('2,469.00');
    // Embedded quotes double, embedded newline and comma stay inside one field.
    expect(rows[2]![5]).toBe('Quoted "as-is"');
    expect(rows[3]![5]).toBe('Multi-line\r\nnote with a comma, and "quotes"');
    expect(rows[4]![0]).toBe('Café');
    expect(rows[4]![5]).toBe('Non-ASCII snowman ☃');
    expect(serialiseCsv(rows)).toBe(text);
  });
});

describe('repair-free verifier', () => {
  function packageOf(parts: Record<string, string>): Uint8Array {
    return zipSync(
      Object.fromEntries(
        Object.entries(parts).map(([name, text]) => [name, new TextEncoder().encode(text)]),
      ),
      { level: 6, mtime: new Date('1980-01-01T00:00:00Z') },
    );
  }

  const contentTypes = `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`;

  const rootRels = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;

  const workbook = `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`;

  const workbookRels = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`;

  const sheet = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`;

  function goodPackage(): Record<string, string> {
    return {
      '[Content_Types].xml': contentTypes,
      '_rels/.rels': rootRels,
      'xl/workbook.xml': workbook,
      'xl/_rels/workbook.xml.rels': workbookRels,
      'xl/worksheets/sheet1.xml': sheet,
    };
  }

  it('accepts a well-formed minimal package', () => {
    const result = verifyPackage(packageOf(goodPackage()));
    expect(result.problems).toEqual([]);
    expect(result.ok).toBe(true);
  });

  it('rejects an unreadable ZIP', () => {
    const result = verifyPackage(new Uint8Array([1, 2, 3, 4]));
    expect(result.ok).toBe(false);
    expect(result.problems[0]!.code).toBe('zip-unreadable');
  });

  it('rejects a missing content type and a dangling override', () => {
    const parts = goodPackage();
    delete parts['[Content_Types].xml'];
    const missing = verifyPackage(packageOf(parts));
    expect(missing.problems.some((problem) => problem.code === 'content-types-missing')).toBe(true);

    const dangling = goodPackage();
    dangling['[Content_Types].xml'] = contentTypes.replace(
      '</Types>',
      '<Override PartName="/xl/missing.xml" ContentType="application/xml"/></Types>',
    );
    const result = verifyPackage(packageOf(dangling));
    expect(result.problems.some((problem) => problem.code === 'content-type-dangling')).toBe(true);
  });

  it('rejects a relationship that points at a missing part', () => {
    const parts = goodPackage();
    parts['xl/_rels/workbook.xml.rels'] = workbookRels.replace(
      'worksheets/sheet1.xml',
      'worksheets/sheet2.xml',
    );
    const result = verifyPackage(packageOf(parts));
    expect(result.problems.some((problem) => problem.code === 'rels-missing-target')).toBe(true);
    expect(result.problems.some((problem) => problem.code === 'sheet-part-missing')).toBe(true);
  });

  it('rejects a traversal target and malformed XML', () => {
    const traversal = goodPackage();
    traversal['xl/_rels/workbook.xml.rels'] = workbookRels.replace(
      'worksheets/sheet1.xml',
      '../../outside.xml',
    );
    const unsafe = verifyPackage(packageOf(traversal));
    expect(unsafe.problems.some((problem) => problem.code === 'rels-unsafe-target')).toBe(true);

    const malformed = goodPackage();
    malformed['xl/worksheets/sheet1.xml'] = '<worksheet><sheetData>';
    const result = verifyPackage(packageOf(malformed));
    expect(result.problems.some((problem) => problem.code === 'xml-malformed')).toBe(true);
  });

  it('rejects a DTD', () => {
    const parts = goodPackage();
    parts['xl/worksheets/sheet1.xml'] =
      '<?xml version="1.0"?><!DOCTYPE worksheet [<!ENTITY x SYSTEM "file:///etc/passwd">]><worksheet>&x;</worksheet>';
    const result = verifyPackage(packageOf(parts));
    expect(result.problems.some((problem) => problem.code === 'xml-doctype')).toBe(true);
  });
});
