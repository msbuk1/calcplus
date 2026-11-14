import { unzipSync, zipSync } from 'fflate';
import { WorkbookOpenError } from './errors';

export interface ReadPackageResult {
  readonly entries: Map<string, Uint8Array>;
  readonly compressedSize: number;
  readonly uncompressedSize: number;
}

const ZIP_LOCAL_FILE_HEADER = 0x04034b50;
const CFB_MAGIC = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1] as const;

const MAX_PACKAGE_BYTES = 100 * 1024 * 1024;
const MAX_TOTAL_UNCOMPRESSED = 1.5 * 1024 * 1024 * 1024;
const MAX_ENTRIES = 20_000;
const MAX_SINGLE_XML_PART = 512 * 1024 * 1024;
const MAX_COMPRESSION_RATIO = 200;

export function isCompoundFile(bytes: Uint8Array): boolean {
  if (bytes.byteLength < CFB_MAGIC.length) return false;
  return CFB_MAGIC.every((value, index) => bytes[index] === value);
}

export function isZip(bytes: Uint8Array): boolean {
  if (bytes.byteLength < 4) return false;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  return view.getUint32(0, true) === ZIP_LOCAL_FILE_HEADER;
}

export function detectUnsupportedContainer(bytes: Uint8Array): void {
  if (isCompoundFile(bytes)) {
    throw new WorkbookOpenError(
      'encrypted',
      'This workbook is encrypted or password-protected. calcplus cannot open it.',
    );
  }
  if (!isZip(bytes)) {
    throw new WorkbookOpenError(
      'unsupported-format',
      'This file format is not supported, or the file is damaged. Open it in Excel and save as .xlsx.',
    );
  }
}

function assertSafeEntryName(name: string): void {
  if (
    name === '' ||
    name.startsWith('/') ||
    name.startsWith('\\') ||
    name.includes('\\') ||
    name.split('/').includes('..')
  ) {
    throw new WorkbookOpenError('corrupt', `Unsafe package entry name: ${name}`);
  }
}

export function readPackage(bytes: Uint8Array): ReadPackageResult {
  if (bytes.byteLength === 0) {
    throw new WorkbookOpenError('corrupt', 'This file is empty or damaged.');
  }
  if (bytes.byteLength > MAX_PACKAGE_BYTES) {
    throw new WorkbookOpenError(
      'too-large',
      'This workbook is too large or unusually compressed to open safely.',
    );
  }
  detectUnsupportedContainer(bytes);
  inspectCentralDirectory(bytes);

  let unzipped: Record<string, Uint8Array>;
  try {
    unzipped = unzipSync(bytes);
  } catch {
    throw new WorkbookOpenError('corrupt', 'This file is damaged or is not a valid workbook.');
  }

  const names = Object.keys(unzipped);
  if (names.length === 0) {
    throw new WorkbookOpenError('corrupt', 'This file is damaged or is not a valid workbook.');
  }
  if (names.length > MAX_ENTRIES) {
    throw new WorkbookOpenError(
      'too-large',
      'This workbook is too large or unusually compressed to open safely.',
    );
  }

  const entries = new Map<string, Uint8Array>();
  let total = 0;
  for (const name of names) {
    assertSafeEntryName(name);
    const entry = unzipped[name]!;
    total += entry.byteLength;
    if (total > MAX_TOTAL_UNCOMPRESSED) {
      throw new WorkbookOpenError(
        'too-large',
        'This workbook is too large or unusually compressed to open safely.',
      );
    }
    if (
      (name.endsWith('.xml') || name.endsWith('.rels')) &&
      entry.byteLength > MAX_SINGLE_XML_PART
    ) {
      throw new WorkbookOpenError(
        'too-large',
        'This workbook is too large or unusually compressed to open safely.',
      );
    }
    entries.set(name, entry);
  }

  return {
    entries,
    compressedSize: bytes.byteLength,
    uncompressedSize: total,
  };
}

interface CentralDirectoryEntry {
  readonly name: string;
  readonly compressedSize: number;
  readonly uncompressedSize: number;
}

/**
 * Reads the ZIP central directory before decompression so a zip bomb is
 * rejected on its declared sizes rather than after fflate allocates the
 * output. ZIP64 is rejected: an .xlsx at the guaranteed envelope is far below
 * the 4 GB ZIP32 limit.
 */
function inspectCentralDirectory(bytes: Uint8Array): CentralDirectoryEntry[] {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const minimumEocd = Math.max(0, bytes.byteLength - 22 - 0xffff);
  let eocd = -1;
  for (let index = bytes.byteLength - 22; index >= minimumEocd; index -= 1) {
    if (view.getUint32(index, true) === 0x06054b50) {
      eocd = index;
      break;
    }
  }
  if (eocd < 0) {
    throw new WorkbookOpenError('corrupt', 'This file is damaged or is not a valid workbook.');
  }
  const entryCount = view.getUint16(eocd + 10, true);
  const directoryOffset = view.getUint32(eocd + 16, true);
  if (entryCount > MAX_ENTRIES) {
    throw new WorkbookOpenError(
      'too-large',
      'This workbook is too large or unusually compressed to open safely.',
    );
  }
  if (directoryOffset >= bytes.byteLength) {
    throw new WorkbookOpenError('corrupt', 'This file is damaged or is not a valid workbook.');
  }

  const entries: CentralDirectoryEntry[] = [];
  let totalUncompressed = 0;
  let offset = directoryOffset;
  for (let index = 0; index < entryCount; index += 1) {
    if (offset + 46 > bytes.byteLength || view.getUint32(offset, true) !== 0x02014b50) {
      throw new WorkbookOpenError('corrupt', 'This file is damaged or is not a valid workbook.');
    }
    const compressedSize = view.getUint32(offset + 20, true);
    const uncompressedSize = view.getUint32(offset + 24, true);
    if (compressedSize === 0xffffffff || uncompressedSize === 0xffffffff) {
      throw new WorkbookOpenError(
        'too-large',
        'This workbook is too large or unusually compressed to open safely.',
      );
    }
    const nameLength = view.getUint16(offset + 28, true);
    const extraLength = view.getUint16(offset + 30, true);
    const commentLength = view.getUint16(offset + 32, true);
    const nameStart = offset + 46;
    const nameEnd = nameStart + nameLength;
    if (nameEnd > bytes.byteLength) {
      throw new WorkbookOpenError('corrupt', 'This file is damaged or is not a valid workbook.');
    }
    const name = new TextDecoder('utf-8').decode(bytes.subarray(nameStart, nameEnd));
    if (!name.endsWith('/') && compressedSize > 0) {
      if (uncompressedSize / compressedSize > MAX_COMPRESSION_RATIO) {
        throw new WorkbookOpenError(
          'too-large',
          'This workbook is too large or unusually compressed to open safely.',
        );
      }
    }
    totalUncompressed += uncompressedSize;
    if (totalUncompressed > MAX_TOTAL_UNCOMPRESSED) {
      throw new WorkbookOpenError(
        'too-large',
        'This workbook is too large or unusually compressed to open safely.',
      );
    }
    if (
      (name.endsWith('.xml') || name.endsWith('.rels')) &&
      uncompressedSize > MAX_SINGLE_XML_PART
    ) {
      throw new WorkbookOpenError(
        'too-large',
        'This workbook is too large or unusually compressed to open safely.',
      );
    }
    entries.push({ name, compressedSize, uncompressedSize });
    offset = nameEnd + extraLength + commentLength;
  }
  return entries;
}

export function writePackage(entries: ReadonlyMap<string, Uint8Array>): Uint8Array {
  const zippable: Record<string, Uint8Array> = {};
  for (const [name, bytes] of entries) {
    assertSafeEntryName(name);
    zippable[name] = bytes;
  }
  return zipSync(zippable, {
    level: 6,
    mtime: new Date('1980-01-01T00:00:00Z'),
  });
}

export const PACKAGE_LIMITS = Object.freeze({
  maxPackageBytes: MAX_PACKAGE_BYTES,
  maxTotalUncompressed: MAX_TOTAL_UNCOMPRESSED,
  maxEntries: MAX_ENTRIES,
  maxSingleXmlPart: MAX_SINGLE_XML_PART,
});
