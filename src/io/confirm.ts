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
 * File-compat §6.3: saving a macro workbook as `.xlsx` drops the macros and
 * its digital signature, so that save needs an explicit confirmation first.
 * A plain Save keeps `.xlsm` and never asks.
 */
export function macroRemovalRequired(options: {
  readonly saveAs: boolean;
  readonly hasMacros: boolean;
}): boolean {
  return options.saveAs && options.hasMacros;
}

export const MACRO_REMOVAL_MESSAGE =
  'This workbook contains macros. Saving it as .xlsx removes them and invalidates any digital signature.';

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
  return confirmFn(`${MACRO_REMOVAL_MESSAGE} Continue?`);
}

/** The format and name a save writes; the confirmation gate is `planMacroSave`. */
export function macroSavePlan(options: {
  readonly fileName: string;
  readonly hasMacros: boolean;
  readonly saveAs: boolean;
}): MacroSavePlan {
  if (!options.saveAs) {
    // Plain Save stays in the source format, so the macro parts survive.
    return {
      format: options.hasMacros ? 'xlsm' : formatForFileName(options.fileName),
      fileName: options.fileName,
    };
  }
  return { format: 'xlsx', fileName: toXlsxName(options.fileName) };
}

export function planMacroSave(options: {
  readonly fileName: string;
  readonly hasMacros: boolean;
  readonly saveAs: boolean;
  readonly confirmFn?: (message: string) => boolean;
}): MacroSavePlan | null {
  if (macroRemovalRequired(options) && !confirmMacroRemoval(options.confirmFn)) return null;
  return macroSavePlan(options);
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
export function signatureInvalidationRequired(options: {
  readonly hasDigitalSignatures: boolean;
  readonly changed: boolean;
}): boolean {
  return options.hasDigitalSignatures && options.changed;
}

export function planSignatureSave(options: SignatureSaveInput): boolean {
  if (!signatureInvalidationRequired(options)) return true;
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

export interface ExternalChangeOption {
  /** The outcome this button chooses. */
  readonly choice: Exclude<ExternalChangeChoice, 'cancel'>;
  readonly label: string;
  /** Plain statement of what this choice does and what it loses. */
  readonly statement: string;
}

export interface ExternalChangePrompt {
  readonly displayName: string;
  readonly summary: string;
  /** The outcome buttons in Persistence §4 order: Reload, Overwrite, Save As. */
  readonly outcomes: readonly ExternalChangeOption[];
  /** The write-nothing button, last in the contract's order. */
  readonly cancelLabel: string;
}

/**
 * The one prompt an external edit raises. It names both destructive outcomes
 * plainly: Reload loses the local changes, Overwrite loses the external edit.
 */
export function externalChangePrompt(displayName: string): ExternalChangePrompt {
  return {
    displayName,
    summary: `"${displayName}" changed on disk since it was opened. Choose what to keep.`,
    outcomes: [
      {
        choice: 'reload',
        label: 'Reload',
        statement: `Reload — discard your unsaved changes and take the newer "${displayName}" from disk. Your unsaved changes are lost.`,
      },
      {
        choice: 'overwrite',
        label: 'Overwrite',
        statement: `Overwrite — keep your unsaved changes and overwrite the newer "${displayName}" on disk. The external edit is lost.`,
      },
      {
        choice: 'saveAs',
        label: 'Save As',
        statement: 'Save As — write your unsaved changes to a new file. Nothing is lost.',
      },
    ],
    cancelLabel: 'Cancel',
  };
}

export type ExternalChangeAsker = (prompt: ExternalChangePrompt) => Promise<ExternalChangeChoice>;

/**
 * Persistence §4: an external edit offers Reload (discard local changes),
 * Overwrite (keep local) or Save As (new file), plus a write-nothing choice.
 * The answer comes from one in-page dialog, not a cascade of browser confirms:
 * a confirm cannot carry the user gesture a file picker needs, so the dialog's
 * Save As button opens the picker from its own click handler and the save flow
 * resumes with the chosen branch. Declining every branch resolves to `cancel`
 * and writes nothing; the app never overwrites a newer external file silently.
 */
export function chooseExternalChange(
  displayName: string,
  ask: ExternalChangeAsker = defaultExternalChangeAsker,
): Promise<ExternalChangeChoice> {
  return ask(externalChangePrompt(displayName));
}

/** No dialog is wired: decline. Persistence §4: declining writes nothing. */
function defaultExternalChangeAsker(): Promise<ExternalChangeChoice> {
  return Promise.resolve('cancel');
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
