# PROTOTYPE — calcplus UI shell (throwaway)

This is a throwaway UI prototype for ticket 10. It is **not production code**.

It answers: what should calcplus's shell look like, and how is the chrome structured?
It shows three structurally different chrome variants on one page, switchable with
`?variant=A|B|C` or the floating bar at the bottom.

- **A — Ribbon**: full Excel-style ribbon with tabs and groups.
- **B — Compact**: a single command bar with a compact tab dropdown and overflow.
- **C — Classic**: a menu bar plus one toolbar row.

Shared in all variants: title bar, name box, formula bar, grid, sheet tabs, status bar,
context menu and a Format Cells dialog. Every control dispatches a fake command ID and
the state panel records it, demonstrating the command-registry architecture.

## Run

```bash
python3 -m http.server 8080 --directory prototype/ui-shell
```

Open <http://localhost:8080/?variant=A>.

Arrow keys navigate the grid. Click a cell to select it. Right-click for the context menu.
`Alt+←` / `Alt+→` cycle variants; the floating bar also works. The `state` button in the
status bar shows the current prototype state.

## Verdict

Recorded in [ticket 10](../../.scratch/calcplus/issues/10-ui-chrome.md) once a variant is chosen.
