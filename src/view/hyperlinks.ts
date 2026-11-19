/**
 * Hyperlink follow-through rules (feature matrix §8 "Hyperlinks: rendered and
 * followable (Ctrl+click)").
 *
 * The model stores an external target as the relationship's URL and an internal
 * jump as `location` text (file-compat §3, §9 F13). This module is the pure
 * edge between those records and the surface: it decides which external
 * addresses may be opened and splits an internal location into a sheet name and
 * an A1 reference. It never opens a window or touches the workbook.
 */

/**
 * The only schemes a followed hyperlink may open. Everything else — notably
 * `javascript:`, `data:`, `vbscript:` and `file:` — is refused with a message.
 * The app is a local workbook viewer, so there is no host to navigate to.
 */
export const ALLOWED_LINK_SCHEMES: ReadonlySet<string> = new Set(['http:', 'https:', 'mailto:']);

export type ExternalLinkCheck =
  | { readonly kind: 'open'; readonly url: string }
  | { readonly kind: 'refused'; readonly message: string };

/**
 * Check an external target against the scheme allowlist.
 *
 * The address is normalised through the platform URL parser, which strips the
 * control characters a hostile workbook can hide inside a scheme (`java\nscript:`)
 * and lower-cases the protocol, so the comparison is against the scheme the
 * browser would actually use. A target with no scheme (a bare host name or a
 * relative path) cannot be opened and is refused, as is an unparseable target.
 */
export function checkExternalLink(target: string): ExternalLinkCheck {
  const trimmed = target.trim();
  if (trimmed === '') return refused('the link has no address');
  let url: URL;
  try {
    url = new URL(trimmed);
  } catch {
    return refused('only http, https and mailto addresses can be opened');
  }
  const scheme = url.protocol.toLowerCase();
  if (!ALLOWED_LINK_SCHEMES.has(scheme)) {
    return refused(
      `the ${scheme} scheme is not allowed; only http, https and mailto links can be opened`,
    );
  }
  return { kind: 'open', url: url.href };
}

function refused(reason: string): ExternalLinkCheck {
  return { kind: 'refused', message: `Hyperlink not followed: ${reason}.` };
}

/** A `location` attribute split into its sheet and reference parts. */
export interface InternalLocation {
  /** The target sheet, or `null` when the location names only a reference. */
  readonly sheetName: string | null;
  /** The A1 address or range text, as authored (absolute markers kept). */
  readonly reference: string;
}

/**
 * Split an internal `location` such as `Sheet2!A1`, `'Quoted Name'!A1` or
 * `'Quoted Name'!A1:B2`. Excel writes the sheet name quoted only when it needs
 * quoting; a doubled `''` inside the quotes is one literal apostrophe.
 *
 * `null` means the text cannot be a location: an empty string, an empty
 * sheet name or reference, or an unterminated quoted name. The caller reports
 * that as a link that cannot be resolved rather than guessing a cell.
 */
export function parseInternalLocation(location: string): InternalLocation | null {
  let text = location.trim();
  if (text.startsWith('#')) text = text.slice(1).trim();
  if (text === '') return null;
  // The sheet separator is the last `!`, so a quoted name containing `!` still
  // splits at the separator Excel wrote.
  const separator = text.lastIndexOf('!');
  if (separator < 0) {
    return { sheetName: null, reference: text };
  }
  const rawSheet = text.slice(0, separator).trim();
  const reference = text.slice(separator + 1).trim();
  if (rawSheet === '' || reference === '') return null;
  const sheetName = unquoteSheetName(rawSheet);
  if (sheetName === null || sheetName === '') return null;
  return { sheetName, reference };
}

/**
 * The reference text to resolve through the workbook's A1 parser. The sheet
 * name is always quoted, which is valid for every name and is what a name with
 * a space or an apostrophe needs.
 */
export function internalLocationText(location: InternalLocation): string {
  if (location.sheetName === null) return location.reference;
  return `${quoteSheetName(location.sheetName)}!${location.reference}`;
}

/** A sheet name in Excel's quoted-reference form, including escaped `'`. */
export function quoteSheetName(name: string): string {
  return `'${name.replaceAll("'", "''")}'`;
}

/** Decode a quoted sheet name, or return an unquoted name unchanged. */
function unquoteSheetName(raw: string): string | null {
  if (!raw.startsWith("'")) return raw;
  if (!raw.endsWith("'") || raw.length < 2) return null;
  return raw.slice(1, -1).replaceAll("''", "'");
}
