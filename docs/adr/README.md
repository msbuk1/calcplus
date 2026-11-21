# Architecture decision records

| ADR | Decision | Status |
|---|---|---|
| [0001](0001-canonical-workbook-model.md) | The workbook model, not the `.xlsx` file, is the source of truth | Accepted |
| [0002](0002-custom-canvas-grid.md) | Build the custom canvas grid; reject Univer | Accepted |
| [0003](0003-preservation-first-ooxml.md) | Preservation-first in-house OOXML layer | Accepted |
| [0004](0004-hyperformula-engine.md) | HyperFormula behind a narrow formula-engine port | Accepted |
| [0005](0005-gplv3.md) | calcplus is GPL-3.0-only | Accepted |
| [0006](0006-patch-in-place-xml.md) | Save by patching the original XML; tree-edit only small parts | Accepted |
| [0007](0007-sheet-projection.md) | The grid reads a projection of the workbook, never the workbook | Accepted |
| [0008](0008-terminal-sheet-delete.md) | Deleting a sheet is terminal: no undo entry, no archive | Accepted |
| [0009](0009-browser-app-window.md) | Open the app in a Chromium app window, not a browser tab | Accepted |

Each ADR records the decision, the options considered and the consequences. The
narrative record is the [spec map](../../.scratch/calcplus/map.md). The
[technical design outline](../technical-design.md) links back here from the
components it describes.

Add the next ADR as the next number (`0009-...`) and add its row to the table
above. Do not rewrite a recorded ADR; supersede it with a new one and say so in
the status column.
