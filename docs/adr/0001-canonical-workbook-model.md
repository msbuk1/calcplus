# The workbook model, not the `.xlsx` file, is the source of truth

calcplus edits a canonical in-memory workbook model and treats `.xlsx` as an import/export format. Opening a file fills the model; saving maps it back and patches only the parts that changed, preserving everything else byte-for-byte.

**Considered options:** the `.xlsx` package as the master (best raw round-trip, but every edit becomes an XML patch and undo/querying suffer), HyperFormula's model as the master (couples the app to the engine), and an event log as the master (strong co-authoring, too much complexity for v1).

**Consequences:** the model must own authored formula text and saved results; the OOXML layer must preserve unmodelled parts; structural edits must go through one range-transform module; the engine's dependency graph is derived, never persisted.
