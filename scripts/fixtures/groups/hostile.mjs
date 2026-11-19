import { writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { unzipSync, zipSync } from 'fflate';
import {
  PACKAGE_REL_NS,
  REL_NS,
  SPREADSHEET_NS,
  XML_HEADER,
  contentTypes,
  packageEntries,
} from '../parts.mjs';

/**
 * F18-F21 (file-compat.md §9): the Tier C fixtures. Each is small, synthetic
 * and deterministic, so `npm run fixtures` regenerates it byte-identically.
 *
 *   F18 encrypted.xlsx     CFB container, rejected on its magic bytes
 *   F19 corrupt.xlsx       a truncated package with no readable directory
 *   F20 zip-bomb.xlsx      a highly compressible entry above the ratio cap
 *   F21 strict-ooxml.xlsx  ISO Strict namespaces throughout
 *
 * Every fixture is inert: nothing here executes, fetches or expands.
 */

const FIXED_MTIME = new Date('1980-01-01T00:00:00Z');
const STRICT_SPREADSHEET_NS = 'http://purl.oclc.org/ooxml/spreadsheetml/main';
const STRICT_REL_NS = 'http://purl.oclc.org/ooxml/officeDocument/relationships';

const hostileWorkbookXml = `${XML_HEADER}
<workbook xmlns="${SPREADSHEET_NS}" xmlns:r="${REL_NS}">
  <workbookPr date1904="0"/>
  <bookViews><workbookView activeTab="0"/></bookViews>
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`;

const hostileWorkbookRels = `${XML_HEADER}
<Relationships xmlns="${PACKAGE_REL_NS}">
  <Relationship Id="rId1" Type="${REL_NS}/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`;

const hostileSheetXml = `${XML_HEADER}
<worksheet xmlns="${SPREADSHEET_NS}">
  <dimension ref="A1"/>
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
  </sheetData>
</worksheet>`;

const strictWorkbookXml = `${XML_HEADER}
<workbook xmlns="${STRICT_SPREADSHEET_NS}" xmlns:r="${STRICT_REL_NS}">
  <workbookPr date1904="false"/>
  <bookViews><workbookView activeTab="0"/></bookViews>
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`;

const strictWorkbookRels = `${XML_HEADER}
<Relationships xmlns="${PACKAGE_REL_NS}">
  <Relationship Id="rId1" Type="${STRICT_REL_NS}/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`;

const strictRootRels = `${XML_HEADER}
<Relationships xmlns="${PACKAGE_REL_NS}">
  <Relationship Id="rId1" Type="${STRICT_REL_NS}/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;

const strictSheetXml = `${XML_HEADER}
<worksheet xmlns="${STRICT_SPREADSHEET_NS}">
  <dimension ref="A1"/>
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
  </sheetData>
</worksheet>`;

function zipFiles(files) {
  const encoder = new TextEncoder();
  return zipSync(
    Object.fromEntries(
      Object.entries(files).map(([name, content]) => [
        name,
        typeof content === 'string' ? encoder.encode(content) : content,
      ]),
    ),
    { level: 6, mtime: FIXED_MTIME },
  );
}

/**
 * A structurally valid, empty Compound File Binary container: a 512-byte
 * header, one FAT sector and one directory sector holding `Root Entry`.
 * An Excel-encrypted workbook uses the same container layout plus the
 * `EncryptionInfo` and `EncryptedPackage` streams, so the eight magic bytes
 * are what a reader sees first.
 */
function compoundFile() {
  const sectorSize = 512;
  const bytes = new Uint8Array(sectorSize * 3);
  const view = new DataView(bytes.buffer);
  const ENDOFCHAIN = 0xfffffffe;
  const FREESECT = 0xffffffff;
  const FATSECT = 0xfffffffd;

  bytes.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1], 0);
  view.setUint16(24, 0x003e, true); // minor version
  view.setUint16(26, 0x0003, true); // major version 3: 512-byte sectors
  view.setUint16(28, 0xfffe, true); // little-endian byte order mark
  view.setUint16(30, 0x0009, true); // sector shift: 2^9 = 512
  view.setUint16(32, 0x0006, true); // mini sector shift: 2^6 = 64
  view.setUint32(40, 0, true); // directory sector count (v3 uses 0)
  view.setUint32(44, 1, true); // one FAT sector
  view.setUint32(48, 1, true); // first directory sector
  view.setUint32(56, 4096, true); // mini stream cutoff
  view.setUint32(60, ENDOFCHAIN, true); // no mini FAT
  view.setUint32(64, 0, true);
  view.setUint32(68, ENDOFCHAIN, true); // no DIFAT extension
  view.setUint32(72, 0, true);
  view.setUint32(76, 0, true); // DIFAT[0]: sector 0 is the FAT
  for (let index = 1; index < 109; index += 1) view.setUint32(76 + index * 4, FREESECT, true);

  for (let sector = 0; sector < 128; sector += 1) {
    view.setUint32(sectorSize + sector * 4, FREESECT, true);
  }
  view.setUint32(sectorSize, FATSECT, true); // FAT sector 0
  view.setUint32(sectorSize + 4, ENDOFCHAIN, true); // directory chain: one sector

  const directoryEntry = sectorSize * 2;
  const rootName = 'Root Entry';
  for (let index = 0; index < rootName.length; index += 1) {
    view.setUint16(directoryEntry + index * 2, rootName.charCodeAt(index), true);
  }
  view.setUint16(directoryEntry + 64, (rootName.length + 1) * 2, true);
  bytes[directoryEntry + 66] = 0x05; // root storage
  bytes[directoryEntry + 67] = 0x01; // black
  view.setUint32(directoryEntry + 68, FREESECT, true);
  view.setUint32(directoryEntry + 72, FREESECT, true);
  view.setUint32(directoryEntry + 76, FREESECT, true);
  view.setUint32(directoryEntry + 116, ENDOFCHAIN, true);
  return bytes;
}

/** A valid package truncated mid-stream, as an interrupted copy would be. */
function truncatedPackage(base, keepBytes) {
  return base.slice(0, Math.max(0, base.length - keepBytes));
}

/**
 * An honest zip bomb: 16 MiB of zeros deflated to a few kilobytes, a ratio
 * far above the 200:1 cap and a total below it, so only the ratio cap can
 * reject it. It is appended to a real package so the shape is realistic.
 */
function zipBomb(base) {
  const files = unzipSync(base);
  files['xl/media/bomb.bin'] = new Uint8Array(16 * 1024 * 1024);
  return zipSync(files, { level: 6, mtime: FIXED_MTIME });
}

function strictPackage() {
  return zipFiles({
    '[Content_Types].xml': contentTypes([
      [
        'xl/workbook.xml',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
      ],
      [
        'xl/worksheets/sheet1.xml',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
      ],
    ]),
    '_rels/.rels': strictRootRels,
    'xl/workbook.xml': strictWorkbookXml,
    'xl/_rels/workbook.xml.rels': strictWorkbookRels,
    'xl/worksheets/sheet1.xml': strictSheetXml,
  });
}

export function build({ compatDir }) {
  const base = packageEntries({
    workbook: hostileWorkbookXml,
    workbookRels: hostileWorkbookRels,
    sheets: [['xl/worksheets/sheet1.xml', hostileSheetXml]],
  });

  const fixtures = [
    ['encrypted.xlsx', compoundFile()],
    ['corrupt.xlsx', truncatedPackage(base, 256)],
    ['zip-bomb.xlsx', zipBomb(base)],
    ['strict-ooxml.xlsx', strictPackage()],
  ];

  for (const [name, bytes] of fixtures) {
    const path = join(compatDir, name);
    writeFileSync(path, bytes);
    console.log(`Wrote ${path}`);
  }
}
