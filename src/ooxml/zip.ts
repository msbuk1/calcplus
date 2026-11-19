import { unzipSync, zipSync } from 'fflate';
import { WorkbookOpenError } from './errors';

export interface ReadPackageResult {
  readonly entries: Map<string, Uint8Array>;
  readonly compressedSize: number;
  readonly uncompressedSize: number;
}

const ZIP_LOCAL_FILE_HEADER = 0x04034b50;
const ZIP_CENTRAL_DIRECTORY_HEADER = 0x02014b50;
const ZIP_END_OF_CENTRAL_DIRECTORY = 0x06054b50;
const CFB_MAGIC = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1] as const;

const MAX_PACKAGE_BYTES = 100 * 1024 * 1024;
const MAX_TOTAL_UNCOMPRESSED = 1.5 * 1024 * 1024 * 1024;
const MAX_ENTRIES = 20_000;
const MAX_SINGLE_XML_PART = 512 * 1024 * 1024;
const MAX_COMPRESSION_RATIO = 200;

/**
 * §5 messages, verbatim. Every rejection path in this module throws one of
 * these, so the surface never has to guess what went wrong.
 */
const ENCRYPTED_MESSAGE =
  'This workbook is encrypted or password-protected. calcplus cannot open it.';
const UNSUPPORTED_CONTAINER_MESSAGE =
  'This file format is not supported, or the file is damaged. Open it in Excel and save as .xlsx.';
const CORRUPT_MESSAGE = 'This file is damaged or is not a valid workbook.';
const TOO_LARGE_MESSAGE = 'This workbook is too large or unusually compressed to open safely.';

function corrupt(message: string = CORRUPT_MESSAGE): never {
  throw new WorkbookOpenError('corrupt', message);
}

function tooLarge(): never {
  throw new WorkbookOpenError('too-large', TOO_LARGE_MESSAGE);
}

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
    throw new WorkbookOpenError('encrypted', ENCRYPTED_MESSAGE);
  }
  if (!isZip(bytes)) {
    throw new WorkbookOpenError('unsupported-format', UNSUPPORTED_CONTAINER_MESSAGE);
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

interface PackageBudget {
  entries: number;
  uncompressed: number;
}

/**
 * The caps in file-compat.md §8. Declared sizes are checked before the entry
 * is inflated and again on the inflated bytes, so a package that lies in its
 * central directory is still bounded by the same budget.
 */
function spendEntryBudget(
  budget: PackageBudget,
  name: string,
  compressedSize: number,
  uncompressedSize: number,
): void {
  budget.entries += 1;
  if (budget.entries > MAX_ENTRIES) tooLarge();
  if (!name.endsWith('/') && uncompressedSize > 0) {
    const ratio = compressedSize > 0 ? uncompressedSize / compressedSize : Infinity;
    if (ratio > MAX_COMPRESSION_RATIO) tooLarge();
  }
  budget.uncompressed += uncompressedSize;
  if (budget.uncompressed > MAX_TOTAL_UNCOMPRESSED) tooLarge();
  if ((name.endsWith('.xml') || name.endsWith('.rels')) && uncompressedSize > MAX_SINGLE_XML_PART) {
    tooLarge();
  }
}

/**
 * Read every entry or throw. Nothing is returned until the whole package has
 * passed the caps, so a rejected package can never leave a partial import.
 */
export function readPackage(bytes: Uint8Array): ReadPackageResult {
  if (bytes.byteLength === 0) {
    throw new WorkbookOpenError('corrupt', 'This file is empty or damaged.');
  }
  if (bytes.byteLength > MAX_PACKAGE_BYTES) tooLarge();
  detectUnsupportedContainer(bytes);
  inspectCentralDirectory(bytes);

  let unzipped: Record<string, Uint8Array>;
  try {
    // The filter runs for each central-directory entry before that entry is
    // decompressed, so a bomb is stopped on its declared sizes and the
    // offending data is never inflated. It re-reads the directory on its own,
    // which keeps the cap true even if the two parses ever disagree.
    const declaredBudget: PackageBudget = { entries: 0, uncompressed: 0 };
    unzipped = unzipSync(bytes, {
      filter: (file) => {
        assertSafeEntryName(file.name);
        spendEntryBudget(declaredBudget, file.name, file.size, file.originalSize);
        return true;
      },
    });
  } catch (error) {
    if (error instanceof WorkbookOpenError) throw error;
    corrupt();
  }

  const names = Object.keys(unzipped);
  if (names.length === 0) corrupt();

  const entries = new Map<string, Uint8Array>();
  const actualBudget: PackageBudget = { entries: 0, uncompressed: 0 };
  for (const name of names) {
    assertSafeEntryName(name);
    const entry = unzipped[name]!;
    spendEntryBudget(actualBudget, name, entry.byteLength, entry.byteLength);
    entries.set(name, entry);
  }

  return {
    entries,
    compressedSize: bytes.byteLength,
    uncompressedSize: actualBudget.uncompressed,
  };
}

/**
 * Reads the ZIP central directory before decompression so a zip bomb is
 * rejected on its declared sizes rather than after fflate allocates the
 * output. ZIP64 is rejected: an .xlsx at the guaranteed envelope is far below
 * the 4 GB ZIP32 limit.
 */
function inspectCentralDirectory(bytes: Uint8Array): void {
  // A ZIP needs at least an end-of-central-directory record.
  if (bytes.byteLength < 22) corrupt();
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const minimumEocd = Math.max(0, bytes.byteLength - 22 - 0xffff);
  let eocd = -1;
  for (let index = bytes.byteLength - 22; index >= minimumEocd; index -= 1) {
    if (view.getUint32(index, true) === ZIP_END_OF_CENTRAL_DIRECTORY) {
      eocd = index;
      break;
    }
  }
  if (eocd < 0) corrupt();

  const diskNumber = view.getUint16(eocd + 4, true);
  const directoryDisk = view.getUint16(eocd + 6, true);
  const totalEntries = view.getUint16(eocd + 8, true);
  const entryCount = view.getUint16(eocd + 10, true);
  const directorySize = view.getUint32(eocd + 12, true);
  const directoryOffset = view.getUint32(eocd + 16, true);
  // A spanned, split or multi-disk archive is never a workbook.
  if (diskNumber !== 0 || directoryDisk !== 0 || totalEntries !== entryCount) corrupt();
  if (entryCount === 0) corrupt();
  if (entryCount > MAX_ENTRIES) tooLarge();
  if (directoryOffset >= bytes.byteLength) corrupt();
  if (directoryOffset + directorySize > bytes.byteLength) corrupt();

  const budget: PackageBudget = { entries: 0, uncompressed: 0 };
  let offset = directoryOffset;
  for (let index = 0; index < entryCount; index += 1) {
    if (
      offset + 46 > bytes.byteLength ||
      view.getUint32(offset, true) !== ZIP_CENTRAL_DIRECTORY_HEADER
    ) {
      corrupt();
    }
    const compressedSize = view.getUint32(offset + 20, true);
    const uncompressedSize = view.getUint32(offset + 24, true);
    if (compressedSize === 0xffffffff || uncompressedSize === 0xffffffff) tooLarge();
    const nameLength = view.getUint16(offset + 28, true);
    const extraLength = view.getUint16(offset + 30, true);
    const commentLength = view.getUint16(offset + 32, true);
    const nameStart = offset + 46;
    const nameEnd = nameStart + nameLength;
    if (nameEnd > bytes.byteLength) corrupt();
    const name = new TextDecoder('utf-8').decode(bytes.subarray(nameStart, nameEnd));
    assertSafeEntryName(name);
    spendEntryBudget(budget, name, compressedSize, uncompressedSize);
    offset = nameEnd + extraLength + commentLength;
  }
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
  maxCompressionRatio: MAX_COMPRESSION_RATIO,
});
