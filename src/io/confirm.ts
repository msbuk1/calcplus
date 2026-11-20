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

/**
 * Persistence §4: closing the document session discards anything unsaved, so a
 * dirty workbook asks first. A clean workbook closes without a prompt.
 */
export function confirmCloseDocument(
  hasUnsavedChanges: boolean,
  confirmFn: (message: string) => boolean = (message) =>
    typeof window === 'undefined' ? false : window.confirm(message),
): boolean {
  if (!hasUnsavedChanges) return true;
  return confirmFn('This workbook has unsaved changes. Discard them and close the workbook?');
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

/**
 * Persistence §2: startup offers Restore or Discard for an unsaved recovery
 * copy. The OK button restores; Cancel discards the copy, which is why the
 * question names both outcomes.
 */
export function recoveryRestoreQuestion(displayName: string): string {
  return `Unsaved changes in "${displayName}" were found from an earlier session. Restore them? Choose Cancel to discard the recovery copy.`;
}

export function confirmRecoveryRestore(
  displayName: string,
  confirmFn: (message: string) => boolean = defaultConfirm,
): boolean {
  return confirmFn(recoveryRestoreQuestion(displayName));
}

/** Persistence §4: every external-change branch, plus the no-action choice. */
export type ExternalChangeChoice = 'reload' | 'overwrite' | 'saveAs' | 'cancel';

export function externalChangeReloadQuestion(displayName: string): string {
  return `"${displayName}" changed on disk since it was opened. Reload it and discard your unsaved changes?`;
}

export function externalChangeOverwriteQuestion(displayName: string): string {
  return `Keep your unsaved changes and overwrite the newer "${displayName}" on disk?`;
}

export const EXTERNAL_CHANGE_SAVE_AS_QUESTION = 'Save your unsaved changes to a new file instead?';

/**
 * Persistence §4: an external edit offers Reload (discard local changes),
 * Overwrite (keep local) or Save As. A browser confirm has two buttons, so the
 * three branches are asked in order; declining every question cancels and
 * writes nothing. The app never overwrites a newer external file silently.
 */
export function chooseExternalChange(
  displayName: string,
  confirmFn: (message: string) => boolean = defaultConfirm,
): ExternalChangeChoice {
  if (confirmFn(externalChangeReloadQuestion(displayName))) return 'reload';
  if (confirmFn(externalChangeOverwriteQuestion(displayName))) return 'overwrite';
  if (confirmFn(EXTERNAL_CHANGE_SAVE_AS_QUESTION)) return 'saveAs';
  return 'cancel';
}

/**
 * Persistence §4: leaving the tab with unsaved changes prompts. Browsers own
 * the wording and ignore a page's custom text on `beforeunload`; the app can
 * only trigger the browser's own prompt. This constant records that limit so
 * the UI never promises a message it cannot show.
 */
export const TAB_CLOSE_WARNING =
  'This workbook has unsaved changes. The browser will ask before leaving the page; browsers do not let a page choose the wording of that prompt.';

export function shouldWarnBeforeUnload(hasUnsavedChanges: boolean): boolean {
  return hasUnsavedChanges;
}

/** Persistence §4: browsers without handles cannot detect external edits. */
export const EXTERNAL_EDIT_UNSUPPORTED_WARNING =
  'This browser saves a downloaded copy. Edits made to the file elsewhere cannot be detected and may be overwritten.';

function defaultConfirm(message: string): boolean {
  return typeof window === 'undefined' ? false : window.confirm(message);
}
