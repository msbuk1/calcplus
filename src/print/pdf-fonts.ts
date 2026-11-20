/**
 * The bundled fonts for the PDF export (print-pdf §4, research/26 §6).
 *
 * Every face is Open Font License 1.1 and ships unmodified in `assets/fonts`
 * with its licence file. PDFKit subsets what it embeds; the renderer renames
 * every subset `CALP+<family>` through the embed seam, so no subset carries a
 * Reserved Font Name (licensing §3.6: Selawik and Liberation both reserve
 * theirs).
 *
 * The workbook's font family is mapped to a face the metric-compatible OFL
 * clone covers:
 *
 * - Segoe UI -> Selawik (Latin only);
 * - Calibri and the pinned Aptos deviation -> Carlito (Latin, Greek, Cyrillic);
 * - Arial -> Liberation Sans, Times New Roman -> Liberation Serif,
 *   Courier New -> Liberation Mono.
 *
 * Anything else has no metric clone in v1 and is routed to the 300 DPI canvas
 * fallback (research/26 §7.2) instead of being silently substituted.
 */

export type PdfFamily =
  'selawik' | 'carlito' | 'liberation-sans' | 'liberation-serif' | 'liberation-mono';

/** One bundled font file. */
export interface PdfFaceAsset {
  /** Stable PDFKit registry key, e.g. `selawik-bold`. */
  readonly id: string;
  readonly family: PdfFamily;
  readonly bold: boolean;
  readonly italic: boolean;
  /** The family name written after `CALP+` in every PDF font descriptor. */
  readonly subsetFamily: string;
  /** `family`, `normal`/`bold` and `normal`/`italic` for the canvas fallback. */
  readonly cssFamily: string;
  readonly fontWeight: 400 | 700;
  readonly fontStyle: 'normal' | 'italic';
  readonly url: string;
}

/** The families the export can embed, in the order the notices list them. */
export const PDF_FAMILIES: readonly PdfFamily[] = [
  'selawik',
  'carlito',
  'liberation-sans',
  'liberation-serif',
  'liberation-mono',
];

const CSS_FAMILIES: Readonly<Record<PdfFamily, string>> = {
  selawik: 'Selawik',
  carlito: 'Carlito',
  'liberation-sans': 'Liberation Sans',
  'liberation-serif': 'Liberation Serif',
  'liberation-mono': 'Liberation Mono',
};

const SUBSET_FAMILIES: Readonly<Record<PdfFamily, string>> = {
  selawik: 'Selawik',
  carlito: 'Carlito',
  'liberation-sans': 'LiberationSans',
  'liberation-serif': 'LiberationSerif',
  'liberation-mono': 'LiberationMono',
};

/** Every bundled face. The URLs are statically analysable for Vite's asset graph. */
export const PDF_FACE_ASSETS: readonly PdfFaceAsset[] = [
  {
    id: 'selawik-regular',
    family: 'selawik',
    bold: false,
    italic: false,
    subsetFamily: SUBSET_FAMILIES.selawik,
    cssFamily: CSS_FAMILIES.selawik,
    fontWeight: 400,
    fontStyle: 'normal',
    url: new URL('../../assets/fonts/Selawik-Regular.ttf', import.meta.url).href,
  },
  {
    id: 'selawik-bold',
    family: 'selawik',
    bold: true,
    italic: false,
    subsetFamily: SUBSET_FAMILIES.selawik,
    cssFamily: CSS_FAMILIES.selawik,
    fontWeight: 700,
    fontStyle: 'normal',
    url: new URL('../../assets/fonts/Selawik-Bold.ttf', import.meta.url).href,
  },
  {
    id: 'carlito-regular',
    family: 'carlito',
    bold: false,
    italic: false,
    subsetFamily: SUBSET_FAMILIES.carlito,
    cssFamily: CSS_FAMILIES.carlito,
    fontWeight: 400,
    fontStyle: 'normal',
    url: new URL('../../assets/fonts/Carlito-Regular.ttf', import.meta.url).href,
  },
  {
    id: 'carlito-bold',
    family: 'carlito',
    bold: true,
    italic: false,
    subsetFamily: SUBSET_FAMILIES.carlito,
    cssFamily: CSS_FAMILIES.carlito,
    fontWeight: 700,
    fontStyle: 'normal',
    url: new URL('../../assets/fonts/Carlito-Bold.ttf', import.meta.url).href,
  },
  {
    id: 'carlito-italic',
    family: 'carlito',
    bold: false,
    italic: true,
    subsetFamily: SUBSET_FAMILIES.carlito,
    cssFamily: CSS_FAMILIES.carlito,
    fontWeight: 400,
    fontStyle: 'italic',
    url: new URL('../../assets/fonts/Carlito-Italic.ttf', import.meta.url).href,
  },
  {
    id: 'carlito-bold-italic',
    family: 'carlito',
    bold: true,
    italic: true,
    subsetFamily: SUBSET_FAMILIES.carlito,
    cssFamily: CSS_FAMILIES.carlito,
    fontWeight: 700,
    fontStyle: 'italic',
    url: new URL('../../assets/fonts/Carlito-BoldItalic.ttf', import.meta.url).href,
  },
  {
    id: 'liberation-sans-regular',
    family: 'liberation-sans',
    bold: false,
    italic: false,
    subsetFamily: SUBSET_FAMILIES['liberation-sans'],
    cssFamily: CSS_FAMILIES['liberation-sans'],
    fontWeight: 400,
    fontStyle: 'normal',
    url: new URL('../../assets/fonts/LiberationSans-Regular.ttf', import.meta.url).href,
  },
  {
    id: 'liberation-sans-bold',
    family: 'liberation-sans',
    bold: true,
    italic: false,
    subsetFamily: SUBSET_FAMILIES['liberation-sans'],
    cssFamily: CSS_FAMILIES['liberation-sans'],
    fontWeight: 700,
    fontStyle: 'normal',
    url: new URL('../../assets/fonts/LiberationSans-Bold.ttf', import.meta.url).href,
  },
  {
    id: 'liberation-sans-italic',
    family: 'liberation-sans',
    bold: false,
    italic: true,
    subsetFamily: SUBSET_FAMILIES['liberation-sans'],
    cssFamily: CSS_FAMILIES['liberation-sans'],
    fontWeight: 400,
    fontStyle: 'italic',
    url: new URL('../../assets/fonts/LiberationSans-Italic.ttf', import.meta.url).href,
  },
  {
    id: 'liberation-sans-bold-italic',
    family: 'liberation-sans',
    bold: true,
    italic: true,
    subsetFamily: SUBSET_FAMILIES['liberation-sans'],
    cssFamily: CSS_FAMILIES['liberation-sans'],
    fontWeight: 700,
    fontStyle: 'italic',
    url: new URL('../../assets/fonts/LiberationSans-BoldItalic.ttf', import.meta.url).href,
  },
  {
    id: 'liberation-serif-regular',
    family: 'liberation-serif',
    bold: false,
    italic: false,
    subsetFamily: SUBSET_FAMILIES['liberation-serif'],
    cssFamily: CSS_FAMILIES['liberation-serif'],
    fontWeight: 400,
    fontStyle: 'normal',
    url: new URL('../../assets/fonts/LiberationSerif-Regular.ttf', import.meta.url).href,
  },
  {
    id: 'liberation-serif-bold',
    family: 'liberation-serif',
    bold: true,
    italic: false,
    subsetFamily: SUBSET_FAMILIES['liberation-serif'],
    cssFamily: CSS_FAMILIES['liberation-serif'],
    fontWeight: 700,
    fontStyle: 'normal',
    url: new URL('../../assets/fonts/LiberationSerif-Bold.ttf', import.meta.url).href,
  },
  {
    id: 'liberation-serif-italic',
    family: 'liberation-serif',
    bold: false,
    italic: true,
    subsetFamily: SUBSET_FAMILIES['liberation-serif'],
    cssFamily: CSS_FAMILIES['liberation-serif'],
    fontWeight: 400,
    fontStyle: 'italic',
    url: new URL('../../assets/fonts/LiberationSerif-Italic.ttf', import.meta.url).href,
  },
  {
    id: 'liberation-serif-bold-italic',
    family: 'liberation-serif',
    bold: true,
    italic: true,
    subsetFamily: SUBSET_FAMILIES['liberation-serif'],
    cssFamily: CSS_FAMILIES['liberation-serif'],
    fontWeight: 700,
    fontStyle: 'italic',
    url: new URL('../../assets/fonts/LiberationSerif-BoldItalic.ttf', import.meta.url).href,
  },
  {
    id: 'liberation-mono-regular',
    family: 'liberation-mono',
    bold: false,
    italic: false,
    subsetFamily: SUBSET_FAMILIES['liberation-mono'],
    cssFamily: CSS_FAMILIES['liberation-mono'],
    fontWeight: 400,
    fontStyle: 'normal',
    url: new URL('../../assets/fonts/LiberationMono-Regular.ttf', import.meta.url).href,
  },
  {
    id: 'liberation-mono-bold',
    family: 'liberation-mono',
    bold: true,
    italic: false,
    subsetFamily: SUBSET_FAMILIES['liberation-mono'],
    cssFamily: CSS_FAMILIES['liberation-mono'],
    fontWeight: 700,
    fontStyle: 'normal',
    url: new URL('../../assets/fonts/LiberationMono-Bold.ttf', import.meta.url).href,
  },
  {
    id: 'liberation-mono-italic',
    family: 'liberation-mono',
    bold: false,
    italic: true,
    subsetFamily: SUBSET_FAMILIES['liberation-mono'],
    cssFamily: CSS_FAMILIES['liberation-mono'],
    fontWeight: 400,
    fontStyle: 'italic',
    url: new URL('../../assets/fonts/LiberationMono-Italic.ttf', import.meta.url).href,
  },
  {
    id: 'liberation-mono-bold-italic',
    family: 'liberation-mono',
    bold: true,
    italic: true,
    subsetFamily: SUBSET_FAMILIES['liberation-mono'],
    cssFamily: CSS_FAMILIES['liberation-mono'],
    fontWeight: 700,
    fontStyle: 'italic',
    url: new URL('../../assets/fonts/LiberationMono-BoldItalic.ttf', import.meta.url).href,
  },
];

const ASSETS_BY_ID: ReadonlyMap<string, PdfFaceAsset> = new Map(
  PDF_FACE_ASSETS.map((asset) => [asset.id, asset]),
);

/**
 * Workbook family names mapped to a bundled family, lower-cased and with
 * collapsed spaces. The list is deliberately finite and holds metric clones
 * only: a family with no metric-compatible clone goes to the 300 DPI fallback
 * rather than being silently replaced with an unrelated face (research/26
 * §7.2). Aptos -> Carlito is the contract's pinned, non-metric deviation
 * (print-pdf §5); nothing else here changes metrics.
 */
const FAMILY_ALIASES: Readonly<Record<string, PdfFamily>> = {
  selawik: 'selawik',
  'segoe ui': 'selawik',
  segoeui: 'selawik',
  carlito: 'carlito',
  calibri: 'carlito',
  aptos: 'carlito',
  arial: 'liberation-sans',
  helvetica: 'liberation-sans',
  'liberation sans': 'liberation-sans',
  'times new roman': 'liberation-serif',
  times: 'liberation-serif',
  'liberation serif': 'liberation-serif',
  'courier new': 'liberation-mono',
  courier: 'liberation-mono',
  'liberation mono': 'liberation-mono',
};

/** The face an unknown-family page falls back to for canvas rendering only. */
export const CANVAS_FALLBACK_STACK = '"Carlito", "Selawik", "Liberation Sans", sans-serif';

/** Resolve a workbook family to a bundled family, or `null` (-> raster). */
export function resolvePdfFamily(family: string): PdfFamily | null {
  const key = family.trim().toLowerCase().replace(/['"]/g, '').replace(/\s+/g, ' ');
  return FAMILY_ALIASES[key] ?? null;
}

/** The file for a family and weight style. */
export function faceAssetFor(
  family: PdfFamily,
  bold: boolean,
  italic: boolean,
): PdfFaceAsset | undefined {
  // Selawik ships no italic file; the renderer selects the upright face and
  // asks PDFKit to shear it, which keeps the Segoe UI metrics intact.
  const style = bold ? (italic ? 'bold-italic' : 'bold') : italic ? 'italic' : 'regular';
  const key =
    family === 'selawik' ? `${family}-${bold ? 'bold' : 'regular'}` : `${family}-${style}`;
  return ASSETS_BY_ID.get(key);
}

/** The Carlito face used when the primary face cannot draw a code point. */
export function supplementAssetFor(bold: boolean, italic: boolean): PdfFaceAsset | undefined {
  return faceAssetFor('carlito', bold, italic);
}

/**
 * The embedded subset's `/BaseFont` and `/FontName`, e.g. `CALP+Selawik`.
 *
 * The style suffix keeps two weights of one family apart: a regular and a bold
 * Carlito subset must not share a BaseFont name in one document.
 *
 * The `CALP+` prefix is the contract's four-letter example, not ISO 32000's
 * six. Poppler's `pdffonts` therefore reports `sub no` (it only recognises
 * six-letter tags); the embedded program is still subsetted — six tables, no
 * `name` table (research/26 §3.3).
 */
export function subsetName(asset: PdfFaceAsset): string {
  const style =
    asset.bold && asset.italic
      ? '-BoldItalic'
      : asset.bold
        ? '-Bold'
        : asset.italic
          ? '-Italic'
          : '';
  return `CALP+${asset.subsetFamily}${style}`;
}

/** The CSS font-family for the canvas fallback, preferring the mapped clone. */
export function canvasFamilyFor(family: string): string {
  const resolved = resolvePdfFamily(family);
  if (resolved === null) return CANVAS_FALLBACK_STACK;
  return `"${CSS_FAMILIES[resolved]}", ${CANVAS_FALLBACK_STACK}`;
}
