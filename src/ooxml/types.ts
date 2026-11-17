import type { SheetId } from '@model/types';

export interface Relationship {
  readonly id: string;
  readonly type: string;
  readonly target: string;
  readonly targetMode: 'Internal' | 'External';
  /** Absolute part name for internal targets; the raw target for external links. */
  readonly resolvedTarget: string;
}

export interface PartInventoryEntry {
  readonly name: string;
  readonly uncompressedSize: number;
  readonly crc32: number;
  readonly sha256: string;
}

export interface SharedStringEntry {
  /** Plain text (runs concatenated). */
  readonly text: string;
  /** Whether the `<si>` contains rich-text runs rather than one plain `<t>`. */
  readonly richText: boolean;
  /** Original `<si>` XML. */
  readonly rawXml: string;
}

export interface SharedStringsTable {
  readonly partPath: string;
  readonly entries: SharedStringEntry[];
  /** Last-known `count` attribute, if present. */
  count: number;
  /** Last-known `uniqueCount` attribute, if present. */
  uniqueCount: number;
}

export interface WorkbookPackage {
  /** Original package bytes, exactly as opened. `null` after a snapshot restore. */
  readonly originalBytes: Uint8Array | null;
  /** Every part in the package, keyed by its OPC part name. */
  readonly entries: Map<string, Uint8Array>;
  readonly contentTypesPath: string;
  readonly rootRelsPath: string;
  readonly workbookPartPath: string;
  readonly workbookRelsPath: string;
  readonly sharedStringsPath: string | null;
  readonly stylesPath: string | null;
  readonly themePath: string | null;
  /** Sheet id -> worksheet part path. */
  readonly sheetPartPaths: Map<SheetId, string>;
  /**
   * Sheet id -> workbook relationship id. Kept separately from the parsed
   * relationship list so a restored snapshot can still patch the sheet list.
   */
  readonly sheetRelationshipIds: Map<SheetId, string>;
  readonly relationships: ReadonlyMap<string, Relationship[]>;
  readonly sharedStrings: SharedStringsTable | null;
  /** External links are never refreshed; this list is report material only. */
  readonly externalLinkParts: readonly string[];
}

export type PatchMode = 'none' | 'position' | 'tree';

export interface PackagePatchReport {
  readonly patchedParts: string[];
  readonly modeByPart: ReadonlyMap<string, PatchMode>;
  readonly appendedSharedStrings: number;
  /** Parts removed by this save: a stale calcChain or a deleted sheet part. */
  readonly removedParts: readonly string[];
}
