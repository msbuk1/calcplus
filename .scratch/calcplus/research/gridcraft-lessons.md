# GridCraft lessons for calcplus

Source repository: https://github.com/storytold/gridcraft
Scope: what transfers to calcplus (browser single-user Excel desktop clone, canvas grid with projection-only rendering, fixed chrome with scrolling grid, command system with single undo, workbook model with append-only styles). Contracts take priority over summaries. Conflicts flagged against `.scratch/calcplus/spec/feature-matrix.md` in versus version two fence and `workbook-model.md` append-only and projection-only rules.

Date read: 2026-11-18. GridCraft version observed: workspace `0.19.1`, edition 2024, Rust 1.95 (from `Cargo.toml`).

## Documents and code read

Primary sources (all sourced facts below come from these unless marked as measurement or inference):

- Overview and layering:
  - https://github.com/storytold/gridcraft (README)
  - https://raw.githubusercontent.com/storytold/gridcraft/main/Cargo.toml
  - https://raw.githubusercontent.com/storytold/gridcraft/main/AGENTS.md
  - https://raw.githubusercontent.com/storytold/gridcraft/main/docs/architecture.md
  - https://raw.githubusercontent.com/storytold/gridcraft/main/ROADMAP.md
  - https://raw.githubusercontent.com/storytold/gridcraft/main/docs/gaps.md
- Grid and interaction:
  - https://raw.githubusercontent.com/storytold/gridcraft/main/crates/ui-egui/src/grid.rs (header comment: virtualised painting of only visible cells, frozen panes, headers, selection, fill handle, resizing, overflow, borders, conditional formats, tables, comments, filters, sparklines, floating charts)
  - https://raw.githubusercontent.com/storytold/gridcraft/main/crates/engine/src/selection.rs
  - https://raw.githubusercontent.com/storytold/gridcraft/main/docs/ui-parity.md
  - https://raw.githubusercontent.com/storytold/gridcraft/main/crates/ui-egui/src/grid.rs lines for `DISPLAY_SCALE`, `WHEEL_ZOOM_LADDER`, `Drag`, `FillTag`, `Trace`, `GridState`
- Engine, commands, history:
  - https://raw.githubusercontent.com/storytold/gridcraft/main/crates/engine/src/lib.rs (Session, DocState, HistoryEntry, Clipboard, Trust, UiRequest, undo groups, journal caps)
  - https://github.com/storytold/gridcraft/tree/main/crates/engine/src (cmd, io, fill, display, selection, clip_html, record, recent)
  - https://raw.githubusercontent.com/storytold/gridcraft/main/docs/control-protocol.md
  - https://raw.githubusercontent.com/storytold/gridcraft/main/docs/mcp.md
  - https://raw.githubusercontent.com/storytold/gridcraft/main/docs/cli.md (`gridcraft-cli check`, `run`, `cat`, `eval`)
- Model:
  - https://raw.githubusercontent.com/storytold/gridcraft/main/crates/model/src/lib.rs
  - https://raw.githubusercontent.com/storytold/gridcraft/main/crates/model/src/store.rs (`CellStore`: copy-on-write sparse storage, 64-row bands behind `Arc`)
- Number formats and display:
  - https://raw.githubusercontent.com/storytold/gridcraft/main/crates/numfmt/src/lib.rs
  - https://raw.githubusercontent.com/storytold/gridcraft/main/crates/engine/src/display.rs (format cache, `cell_text`, `value_text`, `range_text`, status bar stats)
- Fill and clipboard:
  - https://raw.githubusercontent.com/storytold/gridcraft/main/crates/engine/src/fill.rs
  - https://raw.githubusercontent.com/storytold/gridcraft/main/crates/engine/src/clip_html.rs
  - https://raw.githubusercontent.com/storytold/gridcraft/main/docs/file-format-parity.md (clipboard-adjacent HTML export notes, CSV behaviour)
- File compatibility:
  - https://raw.githubusercontent.com/storytold/gridcraft/main/docs/file-format-parity.md
  - https://raw.githubusercontent.com/storytold/gridcraft/main/docs/architecture.md section 5 (files: `engine/io.rs` sniffs content, lenient reader with warnings)

Calcplus contracts compared:
- `.scratch/calcplus/spec/README.md`, `feature-matrix.md`, `workbook-model.md`, `grid-spike.md`, `command-system.md`, `clipboard.md`, `performance.md`, `testing-conformance.md`, `file-compat.md`
- Vocabulary from `CONTEXT.md` (workbook, sheet, cell, style, opaque part, part patch, sheet projection, document state versus view state, style table append-only, touched cell, guaranteed and degraded envelopes, conformance oracle, golden file, release gate)

## Key lessons by focus area

### 1. Canvas grid rendering and virtualization

Sourced fact: GridCraft grid header states virtualised painting of only visible cells (`crates/ui-egui/src/grid.rs` line 1 comment). Architecture document states virtualised grid that only lays out what is on screen (README Why GridCraft: Fast; `docs/architecture.md` known gaps still list no threading in calculation, whole workbook in memory for XLSX reader).

Sourced fact: Grid keeps cached line positions and outlines with weak reference to sheet (`GridState.lines: Option<(Weak<Sheet>, geo::SheetLines)>`), plus cached conditional formatting statistics (`GridState.cf`), cached page ranges, pivot fields, show-details cells, paper pagination (`GridState.pages`, `pivot_fields`, `paper`). Comment in code: while weak reference still upgrades to shown sheet, sheet has not changed (every change goes through `Arc::make_mut`), so cached lines still hold.

Sourced fact: Geometry helpers answer by index rather than iteration in calcplus contract; GridCraft has `Geo` struct for hit testing, `cell_text_y`, `wrapped_cell_text_y`, heading font size scales with zoom (`heading_font_size`), display scale `DISPLAY_SCALE = 0.8` mapping 96 dots per inch sheet units to screen points, wheel zoom ladder with 22 steps from 10 to 400 percent.

Sourced fact: Charts are cached between frames and resolved once per workbook, zoom or fonts change, not on every frame (ROADMAP progress 2026-10-13, #1139). Conditional format rule gathers range statistics only when read and borrows instead of copying per cell (ROADMAP 2026-10-13, #1098). Picture textures decoded at drawn size into byte-bounded cache dropped with workbook (ROADMAP 2026-10-16).

What transfers:
- Only paint visible window plus frozen panes; never build sheet-sized canvas or full-height spacer. This matches calcplus `grid-spike.md` section 6 (viewport-sized panes, no sheet-sized canvas, custom scrollbars).
- Cache line positions and conditional formatting statistics keyed by document version and revision, invalidated on model change via shared ownership check. Aligns with calcplus projection idea.
- Cache charts and heavy primitives across idle frames, not per frame.
- Bound text measurement: share column header dropdown allowance logic (ui-parity notes 12.5 points at 100 to 200 percent zoom, 25 points at 50 percent, from observed Excel via automation interface). Useful for autofit.

### 2. Selection and editing model

Sourced fact: `crates/engine/src/selection.rs` defines `Selection { active, anchor, ranges }`, never empty, last range contains active cell. Methods: `extend_to`, `expand_merges`, `extend_by` (Shift plus arrows with merged cell and hidden line handling), `cycle` (Enter and Tab inside multi-cell selection), `next_corner` (Control plus period), `jump` (Control plus arrow to edge of data region using used range to avoid walking million empty cells), `current_region` (Control plus A, AutoSum, sort, tables; grows one side at a time scanning only lines next to region, linear time; fixed past 10,000 row limit bug).

Sourced fact: Grid `Drag` enum covers Select, Rows, Cols, PointRows, PointCols, MoveCols, ResizeCol and ResizeRow (preview on dragged line, apply to all selected on release), Fill with target and menu flag, Freeze, Move (border drag with replace prompt via `range.move`), Object (floating object with eight handles, stepped flag to merge into one undo step), ChartData (drag corner of values outline to change chart data, #4389).

Sourced fact: Point mode details: Command and Control click adds separate argument with locale separator, Shift click extends from stable sheet-aware anchor; dragging over row or column headers points at whole rows or columns; Control plus arrow and Control plus Shift plus arrow while pointing formula reference; F4 cycles anchors. Reference boxes (RefEdit) in dialogs pick range on grid, collapse dialog while selecting, arrow keys move selection while collapsed, F4 cycles anchors, follows sheet by identity through renames and moves.

Sourced fact: Formula bar grows one to five lines then scrolls, drag bottom edge to set height in whole lines up to half window, chevron or Control plus Shift plus U expands back. In-cell editor is Document Object Model overlay conceptually similar to calcplus spike (GridCraft uses egui editor; calcplus uses Document Object Model editor).

What transfers:
- Selection as active plus anchor plus ranges with merge expansion that keeps whole columns and whole rows as-is (do not expand merges for full column or row selections). Directly reusable logic.
- `jump` using used range to avoid million-cell walk; `current_region` linear scan; `extend_by` hidden-line skipping bounded to 100,000 steps; fill target shape validation (`holds_source_in_corner`).
- Trace arrows model (`Trace { uid, revision, arrows }`) for precedents and dependents if calcplus ever does auditing in version two; note calcplus feature matrix puts formula auditing in version two, so do not build now.
- RefEdit pattern for dialogs that take ranges (Name Manager is version two, but Go To is in version one; Go To already lists names and recent references, multi-address comma support is a good fit).

Conflict flag: GridCraft supports discontiguous multi-area selection, whole column reorder by header drag, border drag with Control or Option copy and Shift insert partially wired. Calcplus feature matrix puts discontiguous Control multi-range selection in version two, reorder and move or copy or hide sheets in version two, drag-and-drop move and Shift-drag insert is in version one but only for cells, not whole-column header reorder. Do not adopt multi-area editing or column reorder for version one.

### 3. Formatting and number formats

Sourced fact: Engine display layer reads display values with number formats applied in `engine/display.rs` for visible window only; charts through `crates/chart` primitives (`docs/architecture.md` data flow step 4).

Sourced fact: `crates/numfmt` is standalone layer zero crate: Excel format codes (sections, conditions, colours, dates, elapsed time, fractions, scientific, fill), General, `TEXT()` function, English only. Key details:
- 255 character limit (`MAX_FORMAT_CODE_CHARS`), longer codes degrade to General and `TEXT` returns `#VALUE!`; linear parse in code length; fraction denominator with 65,536 zeros no longer crashes.
- Parsed formats cached by code (`display.rs` global cache capped at 4,096, per-thread cache capped at 64 in `parse_cache`); `TEXT()` over range parses once per code not per value (ROADMAP 2026-10-12, #806).
- One table of built-in formats; `format_general_fit` with 11-character General width; `general_full` for comma separated values export avoids exponent so reopen does not gain scientific format.
- Colour from `[Red]` or `[Color1]` to `[Color56]` mapped to Red Green Blue values; ribbon category guess via `FormatKind`; `round_as_displayed` for Set precision as displayed; `date_shape` for calendar versus time parts; elapsed handling.
- Conditional formatting number-format colours rendered in grid and Portable Document Format from shared Red Green Blue mappings (#1452).

Sourced fact: Styles are interned in model (`docs/architecture.md`: copy-on-write cell store 64-row bands behind `Arc`, interned styles). Display resolves conditional formatting into record, never mutates model (calcplus already requires this; GridCraft does same via `CfCache` and `CfLook`).

What transfers:
- Keep number format engine separate from model and renderer, with bounded caches keyed by code text only (date system and values applied at format time). Cap cache size and clear or bound, cap code length.
- Single General width constant and full versus fitted General distinction for export versus display. Useful for calcplus comma separated values contract (formulas export as displayed values).
- Parse once per code for bulk `TEXT()` and chart labels.
- Guard hostile codes (long runs, unclosed brackets, huge denominators) with linear caps and fallback to General.

Conflict flag: None with append-only styles. GridCraft interned styles plus append-only matches calcplus `workbook-model.md` section 2 (table is append-only, never renumber, built-in formats never materialised, preserve identifier). Adopt pattern, do not copy Rust code into TypeScript product; reimplement with HyperFormula port boundary intact.

### 4. Ribbon and chrome patterns

Sourced fact: Panels in `crates/ui-egui` share building blocks (`docs/architecture.md` user interface building blocks): theme tokens with light and dark values and contrast tests, ribbon widgets (`big_button`, `small_button`, `split_button`, `toggle_button`), function library module, code editor widget with Python and LaTeX colouring, formula bar layout records, baseline alignment for Complex text.

Sourced fact: Ribbon catalog generated by `cargo xtask parity` from `crates/engine/src/catalog.rs`; 281 of 296 standard ribbon commands in catalog (ROADMAP headline). Every ribbon button, menu item, dialog, keystroke and mouse gesture is a command that agents can run over JavaScript Object Notation control channel, Model Context Protocol server, or Command Line Interface (`docs/architecture.md`, `docs/mcp.md`, `docs/control-protocol.md`).

Sourced fact: Backstage File view (Home, New, Open, Info, Save As, Print, Export, Options), status bar statistics menu (Average, Count, Numerical Count, Minimum, Maximum, Sum via `view.statusBar`), sheet tab bar with 60 percent cap and drag handle, multiple windows per workbook with per-window selection and scroll, freeze panes by dragging freeze handles, Page Layout view with rulers and margin drag.

What transfers:
- Single command catalogue as seam for ribbon, menus, keyboard, formula bar, dialogs. Calcplus already requires this (`command-system.md` section 1: every surface dispatches by stable identifier, no surface mutates model directly). GridCraft is strong evidence that catalogue plus generated parity list works; adopt `commands --search` equivalent and command journal for replay.
- Status bar stats pattern (`display.rs Stats` with deduplication via seen set, hidden rows skipped, spill handling, used-range intersection) is directly reusable for calcplus status bar sum, average, count contract.
- Theme tokens plus contrast tests pattern is useful, but dark mode and theming is version two never in version one per calcplus feature matrix section 9 (dark mode and theming deferred). Do not build theming now; keep tokens minimal.
- Ribbon overflow handling (scroll arrows when tab wider than window, groups do not collapse when narrow is noted gap) and split-button accessible names are useful for compact bar below 900 pixels.

Conflict flag: GridCraft ribbon breadth (281 of 296, contextual tabs, Backstage, Page Layout, multiple native windows, Draw tab, Automate tab) far exceeds calcplus version one. Calcplus feature matrix puts split view, page-break preview, page-layout view, custom views, full page setup, Insert Function dialog, Name Manager, manual calculation, tables authoring, protection authoring in version two or never. Do not adopt those surfaces. Adopt only command dispatch, enablement and checked state, and visible focus return to grid.

### 5. Clipboard and fill

Sourced fact: `clip_html.rs` reads HyperText Markup Language flavour on paste: only table structure, cell text and `x:num` (Excel) and `sdval` (LibreOffice Calc) attributes; ignores everything else; limits: 16 megabyte HyperText Markup Language, 1,000,000 cells, spans clamped to sheet extent; entity decoding bounded to 12-byte window to avoid quadratic rescan; whitespace runs collapsed; `format_for` infers number format that renders shown text back character for character (e.g. `"Rp"#,##0.00`); colspan keeps next cell in column; rowspan keeps later rows in columns.

Sourced fact: `display.rs range_text` writes tab-separated text skipping hidden rows, quoting only text with tab or line break (quotes doubled), matching Excel (`#2225`). `value_text` versus `shown_text` distinction for comma separated values export (displayed versus raw values via `file.exportCsv {values: true}`).

Sourced fact: `fill.rs` (`MAX_FILL_CELLS = 5,000,000`, `holds_source_in_corner` validation, lane times steps bound to avoid 17 billion writes from whole rows into whole columns, refused shapes like `2:3` into `A:A`, `H4:H9` into `H:H`): FillMode Copy, Series, Alternate (Control drag swaps copy and extend), Formats, ValuesOnly, Extend, Dates with Day, Weekday, Month, Year units; pattern detection via least-squares step for numbers, month stepping, weekday skipping Saturday and Sunday, text with trailing or leading number with width preservation, lists (days, months, quarters with case handling), ordinal suffix (`1st Period` to `2nd Period`), quarter wrap, overflow-safe near 64-bit integer max; Series dialog with linear, growth, date, step, stop; Flash Fill learns token, concatenation, case, substring programs from examples and reads shown text (dates as formatted, not serial).

What transfers:
- HyperText Markup Language paste hardening: clamp spans, cap cells, bound entity window, ignore styles and scripts, prefer `x:num` or `sdval` for numeric fidelity, infer format only if it renders back exactly. This matches calcplus `clipboard.md` sections 2 to 5 (read `text/html` first, sanitise, Excel dialect legacy and drift). Adopt limits and `x:num` handling.
- Tab-separated text quoting rule (only quote cells with tab or line break) and hidden-row skipping.
- Fill shape validation and lane-times-steps product bound before allocating writes. Critical for calcplus bulk budget (paste or fill 100k cells within 1.0 second, 1M within 3.0 seconds; 100k-cell paste is one undo step).
- Fill pattern details for version one only: copy with relative-reference adjustment and double-click to last contiguous row. Series linear, growth, date, custom lists, Flash Fill, Fill Without Formatting versus Formats distinction beyond copy are version two per calcplus feature matrix (Paste Special and fill series are version two). Do not adopt series engine for version one except as reference for what to defer.

Conflict flag: GridCraft fill series and Flash Fill exceed version one. Calcplus version one fill handle is copy, relative-reference adjustment, double-click to adjacent data only. Adopt validation and bounds, defer pattern engine.

### 6. File round trip if any

Sourced fact: Extensible Stylesheet Language for Spreadsheets is native format (`docs/file-format-parity.md`): transitional read and write, strict read only; cells, shared and array and dynamic formulas, shared strings with Japanese phonetic runs preserved, styles, themes, merges, conditional formatting including Excel 2010 extension, validation, tables, comments and notes, hyperlinks, names, freeze panes, outline, print settings, pictures, shapes, equations, charts simplified, sparklines, PivotTables caches and tables, protection, calc settings; pass-through (`passthrough.rs`) keeps custom Extensible Markup Language, external links, connections, query tables, printer settings, controls, ActiveX, Object Linking and Embedding, backgrounds, chart sheets; drops slicers, timelines, data model, table query tables, unmodelled extension list entries with warning; never checked in Excel across corpus (repair prompts unknown); 100 megabyte files fail (#175).

Sourced fact: Legacy `.xls` (Excel 97 to 2003) partial read (values, formulas, number formats, merges, sizes, names, cell formatting) via shared `craft-xls` crate, never written; `.xlsb` and OpenDocument Spreadsheet values only; password-encrypted Agile and Standard open with password, never saved with password; comma separated values with delimiter detection and per-column day versus month date-order detection; HyperText Markup Language export with pictures and charts; Portable Document Format via own 1.7 writer.

Sourced fact: Dirty save strategy differs from calcplus: GridCraft rewrites modelled parts and passes through unmodelled parts; calcplus contract requires part patch (in-place edit of part's original Extensible Markup Language, only changed cells, rows and range-bearing elements regenerated) plus append-only style and shared-string tables, byte-identical untouched parts, fallback to tree rewrite under 16 megabyte else refuse (`file-compat.md` sections 3, 5, 9 fixtures F1 to F24).

What transfers:
- Pass-through list as checklist for opaque parts (calcplus calls them opaque parts or preserved features). GridCraft `passthrough.rs` scope helps define calcplus Tier B inventory.
- Shared-formula handling note: GridCraft expands shared formulas on import (`array_export.rs`, #1505) and keeps legacy `INDEX(range, MATCH())` as ordinary formulas; calcplus instead preserves `t="shared"`, `si` and `ref` for untouched groups and flags follower text as derived never written back as authored text (`file-compat.md` section 6.1). Prefer calcplus rule; do not adopt GridCraft expansion.
- Dynamic-array metadata (`cm` in `xl/metadata.xml`) kept on save to avoid implicit intersection change; calcplus version one has no dynamic arrays or spill (testing-conformance section 12 excludes dynamic arrays and spill). Flag as version two or never, do not implement now.
- Google Sheets `IFERROR(__xludf.DUMMYFUNCTION(...))` unwrapping when GridCraft calculates every function inside is clever but out of scope for 52-function version one; note as future robustness, not version one.
- Repair-free validation via `gridcraft-cli check` (package validation: content types, relationships, relationship identifiers unique; round-trip categories compared) is a good model for calcplus local relationship verifier plus manual Excel repair check.

Conflict flag: GridCraft models charts, PivotTables, images, shapes, sparklines, tables, validation enforcement, protection enforcement. Calcplus feature matrix: charts, images, shapes, sparklines are version two preserved but not rendered or authored; PivotTables and Power Query as features are never; tables ListObjects are version two preserved; data validation preserved not enforced; protection preserved not enforced (deliberate deviation). Do not adopt modelling or rendering of Tier B objects. Adopt only preservation plus anchor and series-range transform and block rules (`file-compat.md` section 4 and 7).

### 7. Performance approaches

Sourced fact: Copy-on-write workbooks (undo snapshots nearly free), incremental recalculation, virtualised grid (README). `CellStore` bands of 64 rows behind `Arc`; cloning copies only band pointers; editing copies just its band; `diff` skips shared bands; `retained_bytes_against` estimates heap for history bounding; `contiguous_value_run` walks band row map with range (one in-order pass per band) for long runs.

Sourced fact: History: at most 100 steps and 128 megabyte estimated retained cell storage, newest step retained even if exceeds budget (`docs/architecture.md` data flow step 2; `lib.rs UNDO_LIMIT`, `UNDO_BYTES_LIMIT`, `trim_to_budget`, `push_history`, `bound_tail`, `trim_history`). Journal capped at 10,000 entries and 64 megabyte serialized, oldest half dropped (`JOURNAL_MAX_ENTRIES`, `JOURNAL_MAX_BYTES`). Panic guard in `Session::execute` restores workbook, selection, undo and redo; open groups handling preserved.

Sourced fact: Recalculation marks dependents dirty through dependency graph with range nodes, topological order, spills, `#SPILL!` and cycles; volatile always dirty; single-threaded (no threading in calculation is known gap; large workbooks do not open, streaming read missing, 100k-dependent edit about 0.4 seconds, whole-sheet Fill out of memory issues #152, current region past 10,000 rows fixed).

Sourced fact: Text width callback `TextWidth = Arc<dyn Fn(&str, &Font, bool) -> f32>` for AutoFit; without it Command Line Interface and Model Context Protocol use approximate width; autofit column width includes dropdown space only for filter header cells with point allowances.

What transfers:
- 64-row band copy-on-write plus shared-band diff plus retained-bytes history bound is strong precedent for calcplus single-undo bounded log (calcplus `command-system.md` section 5: last 200 transactions or 64 megabyte, range-based changes keep large pastes compact, single large transaction kept, redo counts). Parameters differ (GridCraft 100 steps or 128 megabyte; calcplus 200 transactions or 64 megabyte). Keep calcplus numbers; adopt estimation method (count shared formulas, values, text, comments, hyperlinks) and always-keep-newest rule.
- Lane-times-steps pre-check plus `MAX_FILL_CELLS` style cap before writes; per-frame conditional formatting statistics borrow not copy; chart resolve once per revision; picture texture byte-bounded cache.
- Dedicated worker separation: calcplus performance contract requires engine in dedicated Web Worker, no recalculation blocks main thread over 50 milliseconds, with progress cue over 100 milliseconds. GridCraft single-threaded calculation is a caution, not a model. Do not adopt single-threaded recalculation; keep HyperFormula worker and incremental after edit, rebuild derived graph on structural edit within budget.

Conflict flag: GridCraft guaranteed envelope is implicit (250k touched cells mentioned only in grid-spike comparison context, 100 megabyte Extensible Stylesheet Language for Spreadsheets fails). Calcplus envelopes are contractual (guaranteed 250,000 touched cells and 50,000 formulas, degraded 1,000,000 and 200,000, refused above). Do not import GridCraft limits; enforce calcplus `performance.md` sections 2 to 4.

### 8. Testing strategy

Sourced fact: Quality gates before every commit: `cargo xtask ci` (format, clippy with deny warnings, tests, assets, layers, WebAssembly) (`AGENTS.md`). Never crash rule: no `unwrap`, `expect`, `panic`, `unreachable`, `todo`, `unimplemented` in non-test code, no `unsafe`, errors via `Result`, hostile input treated as hostile, bounded recursion and allocations, every crash fix lands with regression test that panicked before fix. Fuzzing: `tests::every_command_survives_empty_params` fuzzes every command, `crates/engine/tests/it/command_sequence_fuzz.rs` fuzzes random command sequences against undo, redo, Extensible Stylesheet Language for Spreadsheets round-trip and full recalculation invariants, soak with seed and steps environment variables, replay with seed.

Sourced fact: `gridcraft-cli check` is calc oracle plus round-trip fidelity harness: reads files keeping cached values, recalculates, compares numbers within relative tolerance (absolute when magnitude below 1), errors by kind, text and logicals exactly; skips no-oracle and volatile separately; rolls up mismatches by function and by expected-to-actual signature; round-trip validates package and compares categories (package, sheets, values, formulas, number formats per cell, merges, names, tables, conditional formats, validations, charts, comments, hyperlinks, pivots); `#SPILL!` and `#CALC!` compare equal to `#VALUE!` and `#CIRC!` to 0 as file format stores.

Sourced fact: Command-sequence fuzz found spill, AutoSum, Watch Window, sheet-copy bugs; formula hardening (400-digit numbers as text, Not a Number column index as `#VALUE!`, `REGEXREPLACE` result cap); hostile markup tests (empty, unclosed, huge spans, entities) with cell budget.

What transfers:
- Hostile-input test pattern for clipboard HyperText Markup Language, number format codes, Extensible Markup Language parts: huge spans, overlong codes, bad entities, truncated packages. Adopt fixed corpus and seed, new crashes become regression fixtures (matches calcplus testing-conformance section 10).
- Command-sequence fuzz against undo, redo, round-trip and full recalculation is excellent precedent for calcplus property invariants (range transforms inverse, sort permutation and stability, filter visibility, undo exact snapshot, no-op save byte-identical, dirty part untouched rows byte-identical, number format never throws and deterministic, comparator antisymmetry). Adapt to TypeScript fast-check plus committed generator.
- Oracle comparison with tolerance plus primary versus downstream mismatch attribution (GridCraft primary function rollup) helps calcplus 52-function golden suite and file-compatibility fixtures F1 to F24.
- `commands --search` plus `functions --search` plus `journal` replay supports scripted regression.

Conflict flag: GridCraft runs `cargo xtask ci` and full tests in Continuous Integration; calcplus testing-conformance sections 2 and 11 forbid tests on GitHub, require local runs on maintainer personal computer with stored goldens, GitHub only for source hosting and release downloads. Do not adopt GridCraft Continuous Integration test workflow. Adopt only local cadence idea (per-change under 15 minutes, full local under 60 minutes) and manual Excel oracle with manifest (Excel version, build, Operating System, locale, date system, date).

## Viable options

### Option A: Adopt patterns only, no dependency (recommended)
Reimplement proven patterns inside calcplus architecture: virtualised visible-window grid, selection and jump and current-region logic, format-code cache with caps, HyperText Markup Language paste hardening with span and cell caps and `x:num` handling, lane-times-steps fill validation, banded copy-on-write plus retained-bytes history bound, command catalogue plus journal, oracle plus round-trip harness locally.

Trade-offs:
- For: keeps HyperFormula port, append-only styles, part-patch preservation, projection-only rendering, version one scope fence. No licence or runtime risk. Smallest integration surface.
- Against: requires writing TypeScript equivalents; no free code.

### Option B: Port GridCraft modules directly
Translate Rust crates (`numfmt`, `selection`, `fill`, `clip_html`, `store`, `display`) to TypeScript.

Trade-offs:
- For: saves design time on edge cases (ordinals, quarters, weekdays, entity window, General width).
- Against: language and engine mismatch (GridCraft calculation engine is bespoke, not HyperFormula; copy-on-write `Arc` maps poorly to JavaScript garbage collection); clean-room risk if copied too closely; version two features would leak in unless carefully cut. High review cost.

### Option C: Adopt GridCraft as foundation or engine
Replace calcplus model or engine or grid with GridCraft or Univer-like framework.

Trade-offs:
- For: fastest route to a spreadsheet.
- Against: contradicts grid-spike verdict (custom canvas wins for Excel clone; Univer rejected because chrome, model, engine and interaction are Univer's and `.xlsx` exchange is commercial Pro). Same applies to GridCraft: Rust only, no handwritten JavaScript or TypeScript in native apps, egui not swappable to Document Object Model canvas without rewrite, WebAssembly build exists but calcplus is browser TypeScript with HyperFormula. Breaks append-only and projection-only invariants and version one fence. Not viable for version one.

## Recommendation

Adopt Option A.

Specifically for calcplus version one:

1. Grid: keep custom canvas with viewport-sized panes and custom scrollbars; add GridCraft-style caches keyed by document version (line positions, conditional formatting statistics) and resolve charts or heavy objects once per revision where applicable in version two only.
2. Selection: port `active`, `anchor`, `ranges` semantics, merge expansion exception for full rows and columns, `jump` via used range, linear `current_region`, hidden-line skipping with bound, `cycle` and Go To multi-address support.
3. Number formats: isolate formatter behind port, cache parsed codes with size cap and 255-character fallback to General, parse once per bulk operation, keep fitted versus full General distinction for display versus comma separated values.
4. Clipboard: enforce HyperText Markup Language caps (size, cells, spans), decode entities with bounded window, read `x:num` or `sdval`, infer format only on exact render-back, quote tab-separated values only for tab or line break, skip hidden rows, keep in-memory snapshot as only lossless path.
5. Fill: validate target holds source in corner and bound lane-times-steps before writing; keep version one to copy with relative-reference adjustment and double-click fill; defer series, Flash Fill, Paste Special.
6. History: keep calcplus 200 transactions or 64 megabyte bound with range-based changes; adopt retained-bytes estimation and always-keep-newest, plus panic-safe restore concept adapted to worker error boundary.
7. Commands: keep single registry with stable identifiers, enablement and checked state, transaction boundaries, journal for replay; add command and function search for debugging.
8. Files: keep part-patch plus append-only tables plus Tier A, B, C fence; use GridCraft pass-through inventory only as checklist; do not expand shared formulas or model Tier B objects; keep pivot source and non-range chart source block rules.
9. Testing: add hostile-input fixtures, command-sequence fuzz with fixed seeds against undo, redo, round-trip and recalculation invariants, oracle comparison with tolerance and primary attribution, all local only.

Avoid for version one:
- Multi-area selection editing, column reorder, border copy and insert wiring beyond specified move, split view, Page Layout, multiple windows, Draw, Automate and macros, Python cells, equations, Solver, Analysis ToolPak, Forecast Sheet, Get and Transform, Power Query, collaboration, localisation, dark mode, accessibility mirror beyond keyboard and visible focus, encrypted write, `.xls` and `.xlsb` and OpenDocument Spreadsheet support, Strict Open XML write, dynamic arrays and spill, R1C1 authoring, structured references evaluation, chart and PivotTable modelling or rendering, theme and named-style authoring, Name Manager, Insert Function dialog, formula auditing, manual calculation.

## Open questions

1. Autofit measurement: can calcplus reproduce GridCraft filter-dropdown allowance (12.5 points at 100 to 200 percent zoom, 25 points at 50 percent) with pinned fonts (Selawik, Carlito, Liberation Sans) and still meet 16.7 millisecond frame budget on 250,000-cell formatted sheet? Needs local harness, not decided here.
2. HyperText Markup Language `x:fmla` paste anchoring: GridCraft treats formulas via adjustment on fill; calcplus calls formulas best-effort because anchoring is undocumented. Should version one store `x:fmla` only in in-memory snapshot path and never trust it from Excel HyperText Markup Language? Needs ticket 06 decision.
3. Shared-formula preservation versus expansion: GridCraft expands on import; calcplus preserves `t="shared"`. No conflict if calcplus rule wins, but should conformance fixture F8 explicitly assert GridCraft-style expansion never happens? Needs file-compat confirmation.
4. History byte accounting in JavaScript: can retained-bytes estimation be cheap enough per transaction for 100,000-cell paste within 1.0 second and undo within 1.0 second? Needs benchmark corpus measurement.
5. `general_full` exponent avoidance for comma separated values: does Excel 365 on maintainer personal computer write `124412140710` or `1.24412E+11` for General numbers, and should calcplus golden pin Excel display or keep digits? Needs oracle capture.
6. Command journal size: GridCraft caps at 10,000 entries and 64 megabyte; calcplus log shape has `seq`, `actorId`, `commandId`, `args`, `changes`, `inverseChanges`, `timestamp` for future co-authoring seam but no co-authoring in version one. Should journal and co-authoring log be same structure or separate? Needs command-system clarification.
7. Fill refusal message: GridCraft refuses targets not holding source in corner. Calcplus fill-handle contract only promises drag copy and double-click. Should invalid drag be refused silently (keep selection) or with message? Needs grid milestone decision.

## Licence note

GridCraft is dual-licensed Massachusetts Institute of Technology or Apache 2.0 (README). Calcplus spec package notes General Public License version 3 only. No product code was copied here; patterns only. If any code is ever ported, record licence compatibility and attribution per `THIRD_PARTY_NOTICES.md` and do not copy Office artwork, icons, theme Extensible Markup Language, templates or sample files (GridCraft `AGENTS.md` clean-room and asset rules).
