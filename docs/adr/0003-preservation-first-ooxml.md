# Preservation-first in-house OOXML layer

calcplus reads and writes `.xlsx` with its own package layer: `fflate` for ZIP, `saxes` for SAX parsing, a part store for the OPC package, and a dirty-part writer. Modelled parts map into the canonical model; everything else is copied byte-for-byte and reported.

**Considered options:** the in-house preservation-first layer (chosen); SheetJS CE (a data extractor that drops styles, validation, conditional formatting, charts and external links); SheetJS Pro (closed, quote-only, and still not an in-house model); ExcelJS (drops charts, external links, pivots, VBA and custom properties); xlsx-populate (preserves parts but drops cached formula results); treating the `.xlsx` package as the master (every edit becomes an XML patch).

**Why:** the core promise is high-fidelity round-trip. A library that deserializes into its own rich model can only write back what that model understands. The experiment in ticket 04 confirmed the loss; the preservation-first design is the only one that can honestly say "no silent loss".

**Consequences:** no-op save returns the original bytes; untouched parts are copied; style and shared-string tables are append-only; structural edits go through one range-transform module; the I/O layer plus range transforms is the largest engineering item. See [file compatibility](../../.scratch/calcplus/spec/file-compat.md) and [ticket 04](../../.scratch/calcplus/issues/04-xlsx-io.md).
