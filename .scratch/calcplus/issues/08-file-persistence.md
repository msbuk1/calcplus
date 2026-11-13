# 08 — File handling and persistence

Status: resolved
Type: grilling
Blocked by: 25

## Question

How are workbooks opened, saved, autosaved, and recovered with no backend?

Decide:

- Open/save path: File System Access API vs download/upload, plus fallback for browsers without it
- Autosave cadence and IndexedDB persistence
- Crash recovery and session restore
- Recent files and document identity
- Dirty state, save conflicts, and external-change detection
- What happens to unsupported `.xlsx` features on save
- How this maps to a future cloud backend without rework

## Answer

The persistence contract is in [persistence.md](../spec/persistence.md).

- **Open/save:** File System Access API in Chrome/Edge for in-place save; Safari/Firefox use a picker plus download.
- **Autosave:** a JSON recovery copy in IndexedDB about 3 seconds after the last edit and on hide/close; the real file is only written on explicit save.
- **Recovery:** on restart, offer to restore an unsaved copy; delete it after a successful save.
- **Recent files:** a local list with handles where available; no paths; Safari/Firefox name-only or omitted.
- **Conflicts:** prompt on close; compare `lastModified` and size on save; offer reload, overwrite or Save As.
- **Unsupported parts:** preserve byte-for-byte; show a compatibility report; warn or block only when an anchor cannot be transformed (ticket 12).
- **Multiple documents:** one workbook per tab.
- **Future cloud:** local-first; stable IDs, the snapshot and the command log are the sync surface.
