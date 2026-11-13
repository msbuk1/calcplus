# 20 — Conditional-formatting rendering

Status: claimed
Type: grilling
Blocked by: none

## Question

Which conditional-formatting rules must v1 evaluate, and how?

Decide:

- The exact rule set: cell-value comparisons, text contains/begins/ends, dates, top/bottom N/%, duplicates/unique, data bars, colour scales, icon sets. Which are in and which are v2.
- Evaluation order, priority, stop-if-true, multiple rules on one cell, and interaction with manual formatting.
- Relative vs absolute formula references inside rules; cross-sheet references.
- Rendering under merged cells, freeze panes, zoom and canvas clipping.
- Performance on large ranges; incremental re-evaluation after edits.
- Round-trip: preserve rule XML; transform ranges on structural edits.
- What renders when a rule type is unsupported.

Output: the rendering contract and acceptance fixtures for tickets 07, 12 and 13.
