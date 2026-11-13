# 20 — Conditional-formatting rendering

Status: resolved
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

## Answer

The contract is in [conditional-formatting.md](../spec/conditional-formatting.md), with 20 fixtures.

**Rendered:** cell-value comparisons; text contains/does-not-contain/begins/ends; date comparisons and time-period presets; top/bottom N and N%; duplicate/unique; two- and three-colour scales; data bars; all built-in icon sets. **Preserved but unpainted:** formula-based rules, above-average, blank/error rules, formula thresholds and unknown operators. **V2:** authoring, rule manager and stop-if-true UI.

**Evaluation:** rules apply in ascending priority; true rules merge their `dxf` properties with the first writer winning per property; `stopIfTrue` stops lower-priority rules; CF overrides manual formatting only where it sets a property. Blanks coerce to 0 in numeric comparisons. Text rules use the formatted display text, are case-insensitive and support `*`, `?` and `~`. Dates compare as serials; time periods refresh on open, edits, focus and date rollover. Top/bottom includes ties and excludes blanks/errors; duplicate/unique ignores blanks.

**Rendering:** evaluate at the merged anchor and paint the merged rectangle; data bars scale to the merged width; icons draw at the left and shift text right. Each pane clips independently; zoom scales visuals but never thresholds; icons are suppressed below 25%. CF paints after manual fill and borders, bars behind text, icons in front, selection on top. Print/PDF and the accessibility display text follow the same layer.

**Performance:** per-sheet CF index; per-cell rules cached on demand; range statistics cached per contiguous range and invalidated lazily. Visible evaluation under 4 ms per frame; 100k-cell statistics under 200 ms; 250k-cell sheets at 60 fps.

**Round-trip:** read base and x14 `conditionalFormatting`, merge by priority; preserve unknown XML; `dxfs` append-only with stable IDs; structural edits transform `sqref` and references and remove collapsed rules; no-op save is byte-identical. Unsupported rules are listed in the compatibility and export reports.
