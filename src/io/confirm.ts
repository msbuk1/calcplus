import type { SavedFormat } from '@ooxml/part-policy';

/**
 * Persistence §4/§6: opening or creating another document replaces the single
 * M1 document session, so unsaved edits must be protected by a prompt.
 */
export function confirmReplaceDocument(
  hasUnsavedChanges: boolean,
  confirmFn: (message: string) => boolean = (message) =>
    typeof window === 'undefined' ? false : window.confirm(message),
): boolean {
  if (!hasUnsavedChanges) return true;
  return confirmFn(
    'This workbook has unsaved changes. Discard them and replace the current workbook?',
  );
}

export interface MacroSavePlan {
  /** The format the save writes; `xlsx` from a macro workbook drops the macros. */
  readonly format: SavedFormat;
  /** The file name to suggest to the Save As picker. */
  readonly fileName: string;
}

export function formatForFileName(fileName: string): SavedFormat {
  return /\.xlsm$/i.test(fileName) ? 'xlsm' : 'xlsx';
}

/**
 * File-compat §6.3: the default Save keeps `.xlsm`; Save As from a macro
 * workbook converts to `.xlsx` and warns that the macros are removed, only
 * after explicit confirmation. A declined confirmation returns `null` and the
 * caller leaves the file untouched.
 */
export function confirmMacroRemoval(
  confirmFn: (message: string) => boolean = (message) =>
    typeof window === 'undefined' ? false : window.confirm(message),
): boolean {
  return confirmFn(
    'This workbook contains macros. Saving it as .xlsx removes them and invalidates any digital signature. Continue?',
  );
}

export function planMacroSave(options: {
  readonly fileName: string;
  readonly hasMacros: boolean;
  readonly saveAs: boolean;
  readonly confirmFn?: (message: string) => boolean;
}): MacroSavePlan | null {
  if (!options.saveAs) {
    // Plain Save stays in the source format, so the macro parts survive.
    return {
      format: options.hasMacros ? 'xlsm' : formatForFileName(options.fileName),
      fileName: options.fileName,
    };
  }
  if (!options.hasMacros) {
    return { format: 'xlsx', fileName: toXlsxName(options.fileName) };
  }
  if (!confirmMacroRemoval(options.confirmFn)) return null;
  return { format: 'xlsx', fileName: toXlsxName(options.fileName) };
}

/** File-compat §6.5: any save rewrites the package, which invalidates a signature. */
export const SIGNATURE_INVALIDATION_MESSAGE =
  "Saving will invalidate the workbook's digital signature. Save anyway?";

/**
 * File-compat §6.5 and §7: signature parts are preserved, but writing the
 * package invalidates them, so the write is preceded by a confirmation. A
 * declined confirmation returns `false` and the caller must not write.
 */
export function confirmSignatureInvalidation(
  confirmFn: (message: string) => boolean = (message) =>
    typeof window === 'undefined' ? false : window.confirm(message),
): boolean {
  return confirmFn(SIGNATURE_INVALIDATION_MESSAGE);
}

export interface SignatureSaveInput {
  /** `CompatibilityInfo.hasDigitalSignatures`, as `open` surfaces it. */
  readonly hasDigitalSignatures: boolean;
  /**
   * Whether this save would write anything. A save with nothing to write
   * returns the package's original bytes, so the signature is still valid.
   */
  readonly changed: boolean;
  readonly confirmFn?: (message: string) => boolean;
}

/**
 * File-compat §6.5: warn before writing over a signed package. A no-op save
 * copies the original bytes and an unsigned workbook has nothing to lose, so
 * neither asks; returning `false` means the caller writes nothing.
 */
export function planSignatureSave(options: SignatureSaveInput): boolean {
  if (!options.hasDigitalSignatures || !options.changed) return true;
  return confirmSignatureInvalidation(options.confirmFn);
}

function toXlsxName(fileName: string): string {
  if (/\.xlsm$/i.test(fileName)) return fileName.replace(/\.xlsm$/i, '.xlsx');
  if (/\.xlsx$/i.test(fileName)) return fileName;
  return `${fileName}.xlsx`;
}
