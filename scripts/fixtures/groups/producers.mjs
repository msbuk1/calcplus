import { spawnSync } from 'node:child_process';
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { unzipSync, zipSync } from 'fflate';
import {
  XML_HEADER,
  SPREADSHEET_NS,
  packageEntries,
  themeXml,
  stylesXml,
  sharedStringsXml,
  workbookXml,
  workbookRels,
} from '../parts.mjs';

/**
 * M6 lane 6 — the F22 producer matrix and the F24 CSV fixture.
 *
 * file-compat §9 F22 asks that Excel 365, Excel 2019, LibreOffice and Google
 * Sheets exports all open without repair, and testing-conformance §5 widens
 * the matrix to openpyxl, SheetJS and xlsx-populate. Only three producers can
 * be generated on a build machine: openpyxl (Python), LibreOffice headless
 * and the two npm writers. Excel 365/2019 and Google Sheets are captured by
 * the maintainer; this module never fabricates their output.
 *
 * Producer output is normalised before it is written: the ZIP container gets
 * fixed timestamps and stable ordering, and the document timestamps producers
 * stamp with "now" are pinned. The workbook parts themselves are exactly what
 * the producer wrote, which is what the fixture proves. Regeneration is
 * byte-identical on a machine with the same tool versions.
 */

const FIXED_TIMESTAMP = '2026-11-14T00:00:00Z';
const FIXED_ZIP_MTIME = new Date('1980-01-01T00:00:00Z');
const FIXED_PROPERTIES = `${XML_HEADER}
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>calcplus producer fixture</Application>
</Properties>`;

/**
 * Pin the parts a producer stamps with the clock, then repack with fixed
 * timestamps. Keeping the producer's own parts (relationships, content types,
 * worksheet XML) is the point of the fixture; only packaging metadata varies.
 */
export function normaliseProducerPackage(bytes) {
  const entries = unzipSync(bytes);
  const core = entries['docProps/core.xml'];
  if (core !== undefined) {
    let text = new TextDecoder().decode(core);
    text = text
      .replace(
        /(<dcterms:created[^>]*>)[^<]*(<\/dcterms:created>)/,
        `$1${FIXED_TIMESTAMP}$2`,
      )
      .replace(
        /(<dcterms:modified[^>]*>)[^<]*(<\/dcterms:modified>)/,
        `$1${FIXED_TIMESTAMP}$2`,
      );
    entries['docProps/core.xml'] = new TextEncoder().encode(text);
  }
  // The extended properties carry a producer name and version that vary by
  // build; the manifest names the tool instead.
  if (entries['docProps/app.xml'] !== undefined) {
    entries['docProps/app.xml'] = new TextEncoder().encode(FIXED_PROPERTIES);
  }
  const sorted = {};
  for (const name of Object.keys(entries).sort()) sorted[name] = entries[name];
  return zipSync(sorted, { level: 6, mtime: FIXED_ZIP_MTIME });
}

function producerAvailable(command, args) {
  const result = spawnSync(command, args, { stdio: 'ignore' });
  return result.error === undefined && result.status === 0;
}

/** Per-producer availability, the machine-readable form of the run's skips. */
export async function producerAvailability() {
  const probes = [
    {
      id: 'openpyxl',
      available: producerAvailable('python3', ['-c', 'import openpyxl']),
      reason: 'python3 with openpyxl',
    },
    {
      id: 'libreoffice',
      available: producerAvailable('soffice', ['--version']),
      reason: 'soffice (LibreOffice headless)',
    },
    {
      id: 'sheetjs',
      available: await moduleAvailable('xlsx'),
      reason:
        'the xlsx npm package (the fixture is committed; install xlsx@0.18.5 temporarily to regenerate)',
    },
    {
      id: 'xlsx-populate',
      available: await moduleAvailable('xlsx-populate'),
      reason: 'the xlsx-populate npm package',
    },
    {
      id: 'excel-365',
      available: false,
      reason: 'maintainer-captured on the Windows PC (file-compat §9)',
    },
    {
      id: 'excel-2019',
      available: false,
      reason: 'maintainer-captured on a Windows PC with Excel 2019 (file-compat §9)',
    },
    {
      id: 'google-sheets',
      available: false,
      reason: 'manual Google Sheets export (file-compat §9)',
    },
  ];
  return probes;
}

async function moduleAvailable(name) {
  try {
    await import(name);
    return true;
  } catch {
    return false;
  }
}

// ---------------------------------------------------------------------------
// openpyxl
// ---------------------------------------------------------------------------

const OPENPYXL_SCRIPT = join(
  dirname(fileURLToPath(import.meta.url)),
  '..',
  'producers',
  'make-openpyxl-workbook.py',
);

function buildOpenpyxl(outputPath) {
  if (!producerAvailable('python3', ['-c', 'import openpyxl'])) {
    return 'skipped: python3 with openpyxl is not installed.';
  }
  const result = spawnSync('python3', [OPENPYXL_SCRIPT, outputPath], { stdio: 'pipe' });
  if (result.status !== 0) {
    return `skipped: the openpyxl writer failed (${(result.stderr?.toString() ?? '').trim() || result.status}).`;
  }
  writeFileSync(outputPath, normaliseProducerPackage(readFileSync(outputPath)));
  return null;
}

// ---------------------------------------------------------------------------
// LibreOffice
// ---------------------------------------------------------------------------

const libreofficeBaseSheet = `${XML_HEADER}
<worksheet xmlns="${SPREADSHEET_NS}">
  <dimension ref="A1:C4"/>
  <sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetData>
    <row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>LibreOffice</t></is></c><c r="B2"><v>1234.5</v></c></row>
    <row r="3"><c r="A3" t="inlineStr"><is><t>Doubled</t></is></c><c r="B3"><f>B2*2</f><v>2469</v></c></row>
    <row r="4"><c r="A4" t="inlineStr"><is><t>Accent</t></is></c><c r="B4" t="inlineStr"><is><t>café ☃</t></is></c></row>
  </sheetData>
</worksheet>`;

function libreofficeSourcePackage() {
  return packageEntries({
    workbook: workbookXml,
    workbookRels,
    sheets: [['xl/worksheets/sheet1.xml', libreofficeBaseSheet]],
    sharedStrings: sharedStringsXml,
    extra: { 'xl/styles.xml': stylesXml, 'xl/theme/theme1.xml': themeXml },
  });
}

function buildLibreOffice(outputPath) {
  if (!producerAvailable('soffice', ['--version'])) {
    return 'skipped: soffice (LibreOffice headless) is not installed.';
  }
  const workDirectory = mkdtempSync(join(tmpdir(), 'calcplus-lo-'));
  const profile = join(workDirectory, 'profile');
  const sourceDirectory = join(workDirectory, 'source');
  const outputDirectory = join(workDirectory, 'out');
  mkdirSync(sourceDirectory, { recursive: true });
  mkdirSync(outputDirectory, { recursive: true });
  const sourcePath = join(sourceDirectory, 'libreoffice-base.xlsx');
  writeFileSync(sourcePath, libreofficeSourcePackage());
  try {
    const result = spawnSync(
      'soffice',
      [
        `-env:UserInstallation=${pathToFileURL(profile).href}`,
        '--headless',
        '--norestore',
        '--convert-to',
        'xlsx',
        '--outdir',
        outputDirectory,
        sourcePath,
      ],
      { stdio: 'pipe', timeout: 180_000 },
    );
    if (result.status !== 0) {
      return `skipped: the LibreOffice conversion failed (${(result.stderr?.toString() ?? '').trim() || result.status}).`;
    }
    const producedPath = join(outputDirectory, 'libreoffice-base.xlsx');
    writeFileSync(outputPath, normaliseProducerPackage(readFileSync(producedPath)));
    return null;
  } finally {
    rmSync(workDirectory, { recursive: true, force: true });
  }
}

// ---------------------------------------------------------------------------
// SheetJS and xlsx-populate
// ---------------------------------------------------------------------------

async function buildSheetJs(outputPath) {
  let XLSX;
  try {
    XLSX = await import('xlsx');
  } catch {
    return 'skipped: the xlsx npm package is not installed.';
  }
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet([
    ['Producer', 'xlsx'],
    ['Header', 'Value'],
    ['Amount', 1234.5],
    ['Doubled', null],
    ['Accent', 'café ☃'],
  ]);
  sheet.B3 = { t: 'n', v: 1234.5, z: '#,##0.00' };
  sheet.B4 = { t: 'n', f: 'B3*2', v: 2469 };
  XLSX.utils.book_append_sheet(workbook, sheet, 'Producer');
  const written = XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer', compression: true });
  writeFileSync(outputPath, normaliseProducerPackage(new Uint8Array(written)));
  return null;
}

async function buildXlsxPopulate(outputPath) {
  let module;
  try {
    module = await import('xlsx-populate');
  } catch {
    return 'skipped: the xlsx-populate npm package is not installed.';
  }
  const XlsxPopulate = module.default ?? module;
  const workbook = await XlsxPopulate.fromBlankAsync();
  const sheet = workbook.sheet(0);
  sheet.name('Producer');
  sheet.cell('A1').value('Producer');
  sheet.cell('B1').value('xlsx-populate');
  sheet.cell('A3').value('Amount');
  // No number-format style: xlsx-populate 1.21.0 writes an empty `<fill/>`
  // when a number format is applied, which strict readers reject. The fixture
  // proves the producer's plain output opens; the finding is noted in the
  // lane report rather than hidden behind a post-processing fix.
  sheet.cell('B3').value(1234.5);
  sheet.cell('A4').value('Doubled');
  sheet.cell('B4').formula('B3*2');
  sheet.cell('A5').value('Accent');
  sheet.cell('B5').value('café ☃');
  const buffer = await workbook.outputAsync();
  writeFileSync(outputPath, normaliseProducerPackage(new Uint8Array(buffer)));
  return null;
}

// ---------------------------------------------------------------------------
// F24 — csv-roundtrip.csv
// ---------------------------------------------------------------------------

/**
 * A word processor for the CSV fixture. Authored as rows and written by the
 * RFC 4180 rules the module implements, so the committed bytes are reviewable
 * as text and the test can prove an exact round trip. The `Total` column holds
 * displayed values (37.50, 2,469.00) exactly as a formula cell exports them.
 */
const CSV_ROWS = [
  ['Product', 'Region', 'Units', 'Unit price', 'Total', 'Note'],
  ['Widget', 'North', '3', '12.5', '37.50', 'Gift wrap, blue'],
  ['Gadget', 'South', '2', '1234.5', '2,469.00', 'Quoted "as-is"'],
  ['Thing', 'East', '0', '0.5', '0.00', 'Multi-line\r\nnote with a comma, and "quotes"'],
  ['Café', 'West', '1', '€2.50', '€2.50', 'Non-ASCII snowman ☃'],
];

function csvField(field, delimiter) {
  return field.includes(delimiter) || field.includes('"') || /[\r\n]/.test(field)
    ? `"${field.replaceAll('"', '""')}"`
    : field;
}

export function csvRoundTripFixture() {
  const records = CSV_ROWS.map((row) => row.map((field) => csvField(field, ',')).join(','));
  return `${records.join('\r\n')}\r\n`;
}

// ---------------------------------------------------------------------------
// Manifest
// ---------------------------------------------------------------------------

const MANIFEST = {
  clause: 'file-compat §9 F22 producer matrix',
  note: 'Generated producers are normalised so regeneration is byte-identical. Excel 365, Excel 2019 and Google Sheets files are captured by the maintainer and are never generated here.',
  producers: [
    {
      id: 'excel-365',
      file: 'excel-365.xlsm',
      kind: 'maintainer-capture',
      status: 'captured',
      note: 'Captured 2026-11-19 from the maintainer’s demo run (created and saved in Excel 365; macro-enabled export with a chart, hyperlinks and formulas). The manual repair check passed on the same run.',
    },
    {
      id: 'excel-2019',
      file: 'excel-2019.xlsx',
      kind: 'maintainer-capture',
      status: 'absent',
      note: 'Capture an exported workbook on a Windows PC with Excel 2019.',
    },
    {
      id: 'google-sheets',
      file: 'google-sheets.xlsx',
      kind: 'maintainer-capture',
      status: 'absent',
      note: 'Export a small sheet from Google Sheets and commit the download.',
    },
    {
      id: 'libreoffice',
      file: 'libreoffice.xlsx',
      kind: 'script',
      status: 'generated',
      tool: 'LibreOffice headless',
      note: 'Generated by this module; skipped when soffice is not installed.',
    },
    {
      id: 'openpyxl',
      file: 'openpyxl.xlsx',
      kind: 'script',
      status: 'generated',
      tool: 'openpyxl',
      note: 'Generated by scripts/fixtures/producers/make-openpyxl-workbook.py; skipped when openpyxl is not installed.',
    },
    {
      id: 'sheetjs',
      file: 'sheetjs.xlsx',
      kind: 'script',
      status: 'generated',
      tool: 'SheetJS CE (xlsx)',
      note: 'Generated by this module; skipped when the xlsx package is not installed.',
    },
    {
      id: 'xlsx-populate',
      file: 'xlsx-populate.xlsx',
      kind: 'script',
      status: 'generated',
      tool: 'xlsx-populate',
      note: 'Generated by this module; skipped when the xlsx-populate package is not installed.',
    },
  ],
};

// ---------------------------------------------------------------------------
// build
// ---------------------------------------------------------------------------

export async function build({ compatDir }) {
  const directory = join(compatDir, 'producers');
  mkdirSync(directory, { recursive: true });
  const skips = [];

  const producers = [
    ['openpyxl.xlsx', buildOpenpyxl],
    ['libreoffice.xlsx', buildLibreOffice],
    ['sheetjs.xlsx', buildSheetJs],
    ['xlsx-populate.xlsx', buildXlsxPopulate],
  ];
  for (const [name, buildOne] of producers) {
    const path = join(directory, name);
    let skip;
    try {
      skip = await buildOne(path);
    } catch (error) {
      skip = `skipped: ${error instanceof Error ? error.message : String(error)}`;
    }
    if (skip !== undefined && skip !== null) {
      skips.push(`${name}: ${skip}`);
      continue;
    }
    console.log(`Wrote ${path}`);
  }

  const manifestPath = join(directory, 'manifest.json');
  writeFileSync(manifestPath, `${JSON.stringify(MANIFEST, null, 2)}\n`);
  console.log(`Wrote ${manifestPath}`);

  const csvPath = join(compatDir, 'csv-roundtrip.csv');
  writeFileSync(csvPath, csvRoundTripFixture());
  console.log(`Wrote ${csvPath}`);

  for (const skip of skips) console.log(`Producer fixture not generated: ${skip}`);
  console.log(
    'Producers not generated on this machine: excel-365.xlsx, excel-2019.xlsx, google-sheets.xlsx (maintainer captures; no fake fixtures are written).',
  );
}

const invokedDirectly =
  process.argv[1] !== undefined && resolve(process.argv[1]) === fileURLToPath(import.meta.url);

if (invokedDirectly && process.argv.includes('--status')) {
  const availability = await producerAvailability();
  console.log(JSON.stringify({ producers: availability }, null, 2));
}
