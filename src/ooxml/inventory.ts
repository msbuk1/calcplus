import type { PartInventoryEntry } from './types';

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let index = 0; index < 256; index += 1) {
    let value = index;
    for (let bit = 0; bit < 8; bit += 1) {
      value = (value & 1) === 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
    }
    table[index] = value >>> 0;
  }
  return table;
})();

export function crc32(bytes: Uint8Array): number {
  let crc = 0xffffffff;
  for (let index = 0; index < bytes.length; index += 1) {
    crc = CRC_TABLE[(crc ^ bytes[index]!) & 0xff]! ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

export async function sha256Hex(bytes: Uint8Array): Promise<string> {
  const buffer = bytes.buffer.slice(
    bytes.byteOffset,
    bytes.byteOffset + bytes.byteLength,
  ) as ArrayBuffer;
  const digest = await globalThis.crypto.subtle.digest('SHA-256', buffer);
  return [...new Uint8Array(digest)].map((value) => value.toString(16).padStart(2, '0')).join('');
}

export async function createPartInventory(
  parts: ReadonlyMap<string, Uint8Array>,
): Promise<PartInventoryEntry[]> {
  const entries: PartInventoryEntry[] = [];
  for (const name of [...parts.keys()].sort()) {
    const bytes = parts.get(name)!;
    entries.push({
      name,
      uncompressedSize: bytes.byteLength,
      crc32: crc32(bytes),
      sha256: await sha256Hex(bytes),
    });
  }
  return entries;
}

export function inventoryEntriesEqual(
  left: readonly PartInventoryEntry[],
  right: readonly PartInventoryEntry[],
): boolean {
  if (left.length !== right.length) return false;
  return left.every((entry, index) => {
    const other = right[index];
    return (
      other !== undefined &&
      entry.name === other.name &&
      entry.uncompressedSize === other.uncompressedSize &&
      entry.crc32 === other.crc32 &&
      entry.sha256 === other.sha256
    );
  });
}

export function inventoryByName(
  inventory: readonly PartInventoryEntry[],
): Map<string, PartInventoryEntry> {
  return new Map(inventory.map((entry) => [entry.name, entry]));
}
