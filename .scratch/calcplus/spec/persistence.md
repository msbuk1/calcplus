# calcplus v1 — file handling and persistence

Resolved by [08 — File handling and persistence](../issues/08-file-persistence.md) on 2026-11-13. Distribution is in [distribution.md](distribution.md); the workbook model is in [workbook-model.md](workbook-model.md).

## 1. Opening and saving

**Chrome/Edge**

- Open with the File System Access API; accept `.xlsx`, `.xlsm` and `.csv`.
- Keep the file handle for the document session.
- `Ctrl+S` writes to the same handle; **Save As** picks a new file.
- Re-request permission if the handle loses read/write access.
- Drag-and-drop a file onto the app opens it.

**Safari/Firefox**

- Open through a file picker (or drag-and-drop) and read the file into memory.
- Save by downloading a copy (Blob + `<a download>`).
- No in-place save; explain the limitation once, clearly.
- Recent entries are name-only or omitted; no file handles.

**New workbook and CSV**

- A new workbook has no file handle until the first save; the first save uses Save As.
- Opening a CSV creates a new unsaved workbook unless the user saves it as `.xlsx`.
- Saving CSV exports the active sheet's values.

## 2. Autosave and recovery

- Autosave writes a **recovery copy** to IndexedDB: JSON metadata (workbook, sheets, cells, styles, names, side tables, view settings) plus every part payload held as a binary record. Payloads are never base64-encoded inside the JSON.
- Writes are incremental and keyed by part: an edit rewrites only the parts whose bytes or model records changed. A 50 MB workbook is never re-serialised in full every three seconds.
- A write is chunked so that no IndexedDB transaction blocks the main thread for more than 50 ms; a snapshot above 8 MB is written in part-sized batches with the manifest record written last, so a crash mid-write leaves the previous good copy.
- Cadence: debounced about 3 seconds after the last document edit, plus on `visibilitychange` (hidden) and `pagehide`.
- One latest recovery copy per open document, keyed by document ID, plus a small `pendingDocuments` record (document ID, display name, last write) that a restarted app reads to offer Restore. The record is created with the first recovery copy and deleted on save or discard.
- The user's file is written only on explicit Save or Save As. Autosave never writes the file.
- On startup, if an unsaved recovery copy exists, offer **Restore** or **Discard**.
- Delete the recovery copy after a successful save.
- Restore is complete and lossless: it rebuilds the model exactly as at open plus unsaved commands, and a later Save writes a full-fidelity `.xlsx` whose Tier B parts (charts, pivots, macros, external links) survive byte-identical. Recovery must never violate the no-silent-loss goal.
- Restore never re-reads the user’s file; a file changed or deleted while the tab was closed cannot alter or block recovery.
- Request persistent storage (`navigator.storage.persist()`) and handle quota errors by warning and keeping the last good snapshot.
- Recovery copies are local only; nothing is uploaded.
- **Origin stability:** the launcher must keep a stable scheme, host and port. Recovery copies, recent entries and file handles are origin-scoped; a random port would look like data loss. `--port N` creates a separate storage area and must say so.

## 3. Recent files

- A local recent list in IndexedDB: document name, file handle where available, last-opened time and stable document ID.
- Reopening a recent entry re-requests permission if needed; if permission is denied or the file is gone, remove or mark the entry.
- Safari/Firefox: name-only or omitted; no handle persistence.
- No absolute paths are stored or shown.
- Clearing site data clears the list.

## 4. Dirty state and conflicts

- Dirty state means the document has unsaved commands since the last save. Selection, scroll and other view changes do not count.
- Prompt before closing or reloading a tab with unsaved changes.
- On save, compare the file's `lastModified` time and size with the values captured at open or last save.
- If the file changed externally, offer **Reload** (discard local changes), **Overwrite** (keep local) or **Save As** (new file).
- Browsers without handles cannot detect external changes; warn once that external edits may be overwritten.
- No file locking.

## 5. Unsupported parts

- Preserve opaque parts byte-for-byte; never drop them.
- A read-only compatibility report lists what is preserved but not editable or rendered.
- Structural edits transform opaque anchors where possible; if an anchor cannot be transformed safely, warn and block that specific edit.
- The exact feature-by-feature matrix is ticket 12.

## 6. Multiple documents

- One workbook per browser tab.
- Opening another file opens a new tab, or prompts to replace the current document if the user chooses.
- Each tab has its own document session, recovery copy, dirty state and recent-list entry.
- Closing one tab does not affect the others.

## 7. Future cloud

- Local-first: no accounts, no backend, no sync in v1.
- Stable document IDs, the JSON snapshot and the command log are the sync surface.
- A future sync layer can upload snapshots and replay commands without changing the model schema.
- No cloud code in v1.

## 8. Acceptance criteria

1. Chrome/Edge: open `.xlsx`, edit, `Ctrl+S` writes back to the same file; Save As creates a new file.
2. Safari/Firefox: open via picker, save via download; the limitation is explained.
3. Kill the tab mid-edit; reopen; the recovery prompt restores the unsaved changes, and a subsequent save writes charts, pivots, macros and external links byte-identical.
4. Autosave never writes the user's file; only explicit Save does.
5. Recent entries reopen with permission; denied permission prompts and can be removed.
6. An external edit is detected on save; the user chooses reload, overwrite or Save As.
7. Opaque parts survive byte-identical; the compatibility report lists them; an unsafe structural edit warns or blocks.
8. Two tabs hold two workbooks; closing one leaves the other intact.
9. No workbook data leaves the machine.
10. Snapshot and command log use stable IDs; a future sync layer can replay them without a schema change.
11. Recovery copies, recent entries and file handles survive a launcher restart because the origin is stable; `--port N` is documented as separate storage.
12. A 50 MB workbook's recovery copy is written incrementally: an edit rewrites only the parts it changed, no write blocks the main thread for more than 50 ms, and a crash mid-write leaves the previous copy restorable.
