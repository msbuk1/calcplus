# calcplus 0.1.0 release notes

Released 23 November 2026. This is the first public release of calcplus.

## What calcplus is

calcplus is a single-user spreadsheet that runs in the browser. It aims to be
faithful to Excel desktop (Microsoft 365) for a core set of work: a grid you can
edit, formatting, formulas, several sheets, and `.xlsx` files that open in Excel
without a repair prompt. It is free software under GPL-3.0-only.

You can run it two ways:

- **Hosted:** serve the static build from any host over HTTPS. Nothing to
  install; a network is needed for the first load.
- **Local:** download the launcher for your OS. It serves the bundled app from
  your own machine, works offline, and needs no Node, Python or other runtime.

## Downloads

| Platform | Artifact |
|---|---|
| Windows x64 | `calcplus-0.1.0-windows-x64/calcplus.exe` (portable) |
| Linux x64 | `calcplus-0.1.0-linux-x64.tar.gz` (static binary and bundled build) |

Each package carries a `SHA256SUMS` covering every file in it. The release
page carries one covering the artifacts, and
`calcplus-0.1.0-hosted-SHA256SUMS` covers every file of the hosted build. Verify
a download before running it:

```powershell
Get-FileHash .\calcplus.exe -Algorithm SHA256
```

```bash
sha256sum calcplus-0.1.0-linux-x64.tar.gz
```

The release also publishes the source tag and a source archive. The source is at
<https://github.com/msbuk1/calcplus>.

## What's in 0.1.0

This release covers the full v1 feature matrix. The highlights:

**Grid and editing**

- Range selection and keyboard navigation, including Shift-extend, Ctrl+arrow
  and Page Up/Down.
- Typing, F2 and double-click editing, and formula-bar editing. Enter and Tab
  commit; Esc cancels.
- Cut, copy and paste inside the app and through the OS clipboard, including
  values, formulas, formats and merges.
- Fill handle (copy and relative references; double-click to fill to the last
  row), drag-move and Shift-drag insert.
- Insert and delete rows, columns and cells, with reference rewriting and a
  single undo step.
- Row and column resize and autofit; merge and unmerge.
- Find and Replace, Go To, Clear, and drag-and-drop a workbook onto the app to
  open it.
- Undo and redo for every user change.

**Formulas**

- Automatic calculation with 52 functions, using Excel's semantics for errors,
  coercion, operators, dates, references, lookup and circular references.
- Function autocomplete, signature hints, the Insert Function dialog and
  AutoSum.
- Defined names that evaluate and survive a round-trip. Imported names work in
  formulas and Go To.

**Formatting**

- Fonts, fill, borders, alignment and wrap, with a Format Cells dialog and
  Format Painter.
- Number formats: every built-in category and the full custom format grammar.
- Conditional formatting renders imported cell-value, text, date, top/bottom,
  duplicate, data-bar, colour-scale and icon-set rules.
- Read-only fidelity for imported styles, theme colours and rich text.

**Workbook, files and printing**

- New, open, save, save as and close; multi-sheet insert, delete and rename;
  freeze panes; zoom; gridline and heading toggles; status-bar sums.
- `.xlsx` and `.xlsm` open and save through a preservation-first package layer:
  a no-op save is byte-identical, and untouched parts survive an edit
  byte-for-byte. Macro parts are preserved and never executed.
- CSV import and export. Encrypted workbooks are rejected with a clear message
  and are never modified.
- Recent files, a recovery copy, and conflict detection when a file changes on
  disk.
- Print and PDF export: print area, orientation, fit-to-pages, uniform scale
  (10–400%) and margins. The PDF matches the print preview.

**Distribution**

- A Go launcher that serves the embedded build from `http://127.0.0.1:47823/`,
  opens a Chromium app window when it can, and uses your own browser profile.
  It binds to loopback only, makes no outbound calls, and has no updater or
  telemetry.
- One stable address across restarts, so recent files, file handles and the
  recovery copy stay available. `--port N` uses a separate origin and says so.

## First run

**Windows:** the executable is not code-signed. SmartScreen may show "Windows
protected your PC" on the first launch until the download builds a reputation.
Choose **More info**, check that the app is `calcplus.exe`, then choose **Run
anyway**. This is the documented v1 state (limitation L-1), not a defect. Verify
the download with `SHA256SUMS` instead of trusting the prompt.

**Browsers:** Chrome or Edge can open and save a workbook in place. Firefox
opens through a file picker and saves by download. Both keep workbooks on your
machine.

**`file://` is not supported.** Open the app from the hosted URL or through the
launcher. The app shows a clear message if it is opened from a local file.

## Compatibility and known differences

Files saved by calcplus open in Excel 365 without a repair prompt. The M8 check
opened all 64 outputs of the repair corpus — one untouched save and one edited
save of every openable fixture — in Excel 365 (16.0, build 17932).

The v1 behaviour differs from Excel in a few deliberate places. They are
documented and pinned by tests in
[known deviations](known-deviations.md). The ones users notice most:

- Data validation and sheet protection are preserved, not enforced (D-1, D-2).
- Charts, images, shapes and PivotTables are preserved but not rendered (D-4).
  Tables, validation and pivot parts survive a save.
- Conditional formatting renders but cannot be authored (D-5).
- Comments, hyperlinks and defined names are read-only (D-6).
- A structural edit rewrites references and marks a removed target `#REF!`
  rather than refusing the edit (D-8). Internal hyperlink locations are
  rewritten to track their targets; Excel leaves them verbatim (D-9).
- Editing rich text flattens it to plain text (D-11).
- Headers and footers, print titles, manual page breaks, page order and
  independent X/Y scale are v2 (D-3).
- A call to a function outside the published 52 is not argument-count checked,
  because calcplus preserves it verbatim. Excel may refuse such a formula and
  remove it on open (D-13).

Other limitations: no macOS artifact (L-2), Firefox cannot save in place (L-3),
no assistive-technology claim (L-4), `--port N` uses a separate origin (L-5), and
there is no background updater (L-6). Updating means replacing the executable
with a newer release.

Out of scope for v1: charts, PivotTables, tables, data-validation and
conditional-formatting authoring, comments and hyperlink authoring, split view,
dark mode, localisation, co-authoring, macros and Solver. The full matrix is in
the [feature matrix](../.scratch/calcplus/spec/feature-matrix.md).

## Verification

The v1 release gate passes in full. All 14 items of testing and conformance §13
are green, with named evidence in the [release gate record](release-gate.md).
The suite on the release commit is 1527 unit tests and 211 browser tests, all
green, plus the licence gate and the production build.

## Licence and components

calcplus is free software under **GPL-3.0-only**. The licence choice is forced
by the HyperFormula GPLv3 edition, which is the formula engine.

- [LICENSE](../LICENSE) — the full GPLv3 text
- [NOTICE](../NOTICE) — copyright, warranty disclaimer and source link
- [THIRD_PARTY_NOTICES.md](../THIRD_PARTY_NOTICES.md) — every dependency, its
  licence and the obligations
- [CONTRIBUTING.md](../CONTRIBUTING.md) — DCO sign-off and contribution rules

This is the first release, so every component is new against the empty previous
release. The full component list is in `THIRD_PARTY_NOTICES.md`, and the release
attaches a machine-readable SPDX SBOM, `calcplus-0.1.0.spdx.json`. No dependency
carries an incompatible or unknown licence, and no Microsoft asset or
proprietary code is included.

calcplus is an independent project. It is not affiliated with, endorsed by or
sponsored by Microsoft Corporation. Excel is a trademark of Microsoft
Corporation. calcplus reads and writes `.xlsx` files but includes no Microsoft
code, logos or other assets.

## Privacy

No accounts, no telemetry, no workbook uploads. Workbooks never leave the
machine. The hosted build loads only its own static files; the launcher makes
no outbound calls, has no updater and does not check for updates.
