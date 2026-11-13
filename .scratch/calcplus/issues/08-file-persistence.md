# 08 — File handling and persistence

Status: open
Type: grilling
Blocked by: none

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
