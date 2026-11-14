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
