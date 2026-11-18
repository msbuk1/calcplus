# Excel clipboard interop spike (M5)

The M5 spike named in `spec/milestones.md`: "Excel clipboard interop for values,
formulas, formats and merges", which exists to retire the risk-register row
"Clipboard interop with Excel is unreliable".

`clipboard-spike.ps1` runs on the maintainer's Windows PC against real desktop
Excel 365 — the authoritative oracle (`spec/testing-conformance.md` §2) — and
produces **observed** evidence for the Excel-side open questions in
[research 06](../../.scratch/calcplus/research/06-clipboard-interop.md) §10:
questions 1 to 5, 9 and 10. Those answers decide what
[`spec/clipboard.md`](../../.scratch/calcplus/spec/clipboard.md) may promise.

Questions 6 to 8 are browser-side and are walked by hand (checklist below); this
script cannot answer them.

Nothing here runs on GitHub. The capture is manual and local, like the value and
error golden capture in `scripts/oracle/README.md`.

## What the script does

Two probe families, both driven through Excel COM with the same conventions as
`capture-goldens.ps1`: `$ErrorActionPreference = 'Stop'`, a retry helper for the
transient COM busy codes, explicit release of every COM reference,
`AutomationSecurity` forced off so macros never run, and a stable UTF-8 JSON
writer without a byte order mark.

| Family | Direction | Mechanism |
|---|---|---|
| `excel-source` | Excel → clipboard → script | Build cells in Excel, call `Range.Copy`, then read the registered `HTML Format` (CF_HTML) and `CF_UNICODETEXT` payloads through Win32 (`RegisterClipboardFormat`, `GetClipboardData`, `GlobalLock`, `GlobalSize`). The raw bytes are written to disk verbatim. |
| `excel-target` | script → clipboard → Excel | Publish a CF_HTML payload written in the calcplus dialect (and, where the question is about the fallback, a TSV payload), let Excel paste it into a clean sheet, then record a typed cell inventory of the result. |

Each probe carries the research §10 question it answers (or `context` when it
only supplies evidence for a section 3 or 5 claim), what to look for, and the
decision rule that applies to the result. All three are echoed into the evidence
JSON, so the report can be written from the JSON alone.

The script never executes macros, never fetches anything and never uploads
anything. It closes the probe workbook without saving and quits Excel.

## Before you run

1. A Windows PC with an activated desktop Excel 365 (the same machine used for
   the value and error goldens).
2. Windows PowerShell 5.1 (the script is 5.1-compatible; PowerShell 7 also runs
   it). It refuses to run on any non-Windows platform, because CF_HTML and
   CF_UNICODETEXT are Win32 clipboard formats.
3. Close clipboard managers, clipboard history tools and anything else that
   holds the clipboard open. The script retries `OpenClipboard` for a few
   seconds and then fails with a clear message, but a clipboard manager that
   rewrites formats will silently corrupt the target probes.
4. A clean working tree. The script writes only inside `-EvidencePath`, and it
   empties the Windows clipboard while it runs: whatever was on the clipboard
   before the run is gone and cannot be restored.

## Run it

Open Windows PowerShell, go to the repository root, and use a process-scoped
execution policy for this command only — do not change the machine policy:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\oracle\clipboard-spike.ps1 `
  -EvidencePath .\evidence\clipboard-spike-2026-11-18
```

If you prefer the policy scope, the equivalent is
`Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass` in that window,
then `.\scripts\oracle\clipboard-spike.ps1 -EvidencePath ...`.

Partial runs, when a single answer is being chased or re-checked after an Office
update:

```powershell
# one probe
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\oracle\clipboard-spike.ps1 `
  -EvidencePath .\evidence\clipboard-spike-s03 -Probe S03
# one family
... -Probe 'S*'
# the hyperlink and anchoring questions only
... -Probe T06,T07
```

`-Probe` matches with `-like` against the probe identifier, so `T0*` works. An
identifier that matches nothing fails the run and lists the available ones.

A full run takes a couple of minutes; the two 50,000-cell probes (S11 and T12)
dominate, and S11 also pastes its own payload back into a clean sheet.

The script prints one line per probe and then the path of the evidence file. A
failure names the call and the COM HRESULT, for example
`Range.Copy for S03 failed (COM HRESULT 0x800A03EC): ...`.

## Output

```text
<EvidencePath>/
  clipboard-evidence.json      one file: manifest plus one record per probe
  S03-html.txt                 raw CF_HTML bytes exactly as Excel published them
  S03-text.txt                 decoded CF_UNICODETEXT, written as UTF-8 without a BOM
  T07-sent-html.txt            the HTML document the harness published
  T07-sent-text.txt            the TSV the harness published (text-only probes)
```

Source probes use `-html` and `-text` for what was captured. Target probes use
`-sent-html` and `-sent-text` for what was published, so a report can diff
intent against result. `CF_UNICODETEXT` is UTF-16LE on the clipboard; the file is
the decoded text so it diffs and reads cleanly, and the original byte length is
recorded in the JSON beside it.

Do not commit the evidence directory: the raw payloads for the large probes are
megabytes each. Commit only the distilled fixtures the report names.

A failure stops the run. Payload files captured before the failure stay on disk,
but `clipboard-evidence.json` is written only after every selected probe has
finished, so a partial run cannot be mistaken for a complete one. Re-running into
the same directory overwrites the files whose names it produces.

### Manifest

`clipboard-evidence.json` starts with a manifest in the same spirit as the
golden manifests: `excelVersion`, `excelBuild`, `os`, `locale`, `date1904`,
`generatedAt`, plus `spike`, `script`, `research`, `contract`, `probeFilter`,
`probesRequested` and the registered clipboard format identifiers. The workbook
date system is forced to 1900 and recorded, so the date serial 46266 in the
fixtures always means 2026-09-01.

### Probe record

| Field | Meaning |
|---|---|
| `id`, `family`, `title` | identity, and which family produced the record |
| `question` | the research 06 §10 question, or `context` with the section it supports |
| `lookFor` | what the payload or inventory has to show |
| `decisionRule` | what the result changes in the contract |
| `expectations` | the candidate answers written out where arithmetic matters (T07) |
| `payloads` | file name, format, role and byte length of everything written for the probe |
| `timing` | copy call, clipboard read, payload build, clipboard write and paste times |
| `clipboard` | target probes only: the formats published, their byte lengths, the CF_HTML header the harness wrote, and `verified` for the read-back check |
| `observation` | the parsed facts (source probes) or the pasted cell inventory (target probes) |

### Source observations

`htmlFormatAvailable`, `htmlByteLength`, `htmlValidUtf8`,
`htmlSuspectedUtf16Le`, `fragmentSource` (`cfHtmlHeader`, `wholePayload` or
`none`), `cfHtmlHeaderPresent`, `cfHtml` (version, the four byte offsets,
`offsetsConsistent`, `markersPresent`, `fragmentByteLength`, `fragmentExtracted`),
`html` (the document facts below), `textAvailable`, `textByteLength`,
`textLineCount`, `textFirstLines`, and, for S01 only, `clipboardFormats` with the
identifier, hexadecimal identifier and name of every format Excel published, in
the order it published them.

CF_HTML is UTF-8 ([Doc] "HTML Clipboard Format" in research 06 §11).
`htmlValidUtf8` false is itself a
finding (the browser would read mojibake), and `htmlSuspectedUtf16Le` marks the
shape of a payload that looks UTF-16 rather than UTF-8.

The `html` block records: `characterLength`, `detailLimitedToCounts`,
`hasHtmlElement`, `hasProgIdMeta`, `hasExcelNamespace`, `hasOfficeNamespace`,
`hasStyleElement`, `hasTable`, `hasStartFragmentMarker`, `hasHeightSourceStyle`,
`hasSameCellDataPlacement`, `hasScriptElement`, `hasEventHandlerAttribute`,
`programId`, `generator`, `styleBlockCharacterLength`, `tableCellCount`,
`formulaAttributeCount`, `numericAttributeCount`, `textAttributeCount`,
`booleanAttributeCount`, `errorAttributeCount`, `formulaExamples`,
`numericExamples`, `errorValues`, `numberFormats`, `colSpanValues`,
`rowSpanValues`, `heightValues` and `anchorHrefs`. `hasScriptElement` and
`hasEventHandlerAttribute` are the baseline for the paste sanitiser: Excel writes
neither, so a payload that contains them came from somewhere else.

A payload over two million characters records counts only
(`detailLimitedToCounts`), so the 50,000-cell probe does not spend a minute in
regular expressions.

S11 also records `excelPasteMs`, `excelPasteReturned`, `pastedCellCount` and
`pastedRangeAddress` for Excel pasting its own payload back.

### Target observations

`pastedRangeAddress`, `pastedCellCount`, `pasteReturned` (Excel answers `True`
or `False`), `rowHeights` keyed by row number, and `inventory`, which is
`{ cells, truncated, rowCount, columnCount }` with one record per cell:

`ref`, `hasFormula`, `formula`, `kind` (`text`, `number`, `boolean`, `error`,
`blank`), `value`, `numberFormat`, `font` (`name`, `size`, `bold`, `italic`,
`underline`, `colorIndex`), `fill` (`pattern`, `colorIndex`,
`patternColorIndex`), `borders` (four edges with `lineStyle`, `weight`,
`colorIndex`), `alignment` (`horizontal`, `vertical`, `wrapText`,
`indentLevel`, `orientation`), `merge` (`merged`, `area`), `rowHeight` and
`hyperlink` (`count`, and `address`, `subAddress`, `textToDisplay` when one is
present).

The inventory stops at 200 cells and sets `truncated`. T12 records no inventory;
it records three sampled cells (`topLeft`, `middle`, `bottomRight`) plus the
pasted cell count instead.

Values are read from `Value2`, never from `Range.Text`, which is
width-dependent (`####`), exactly as `capture-goldens.ps1` does. Enum-valued
properties are translated to names so the JSON is readable on its own; an
unknown value keeps its number.

<details>
<summary>Excel constant key used by the inventory</summary>

| Field | Values |
|---|---|
| `alignment.horizontal` | `general`, `left`, `center`, `right`, `fill`, `justify`, `centerAcrossSelection`, `distributed` |
| `alignment.vertical` | `top`, `center`, `bottom`, `justify`, `distributed` |
| `borders[].lineStyle` | `none`, `continuous`, `dash`, `dot`, `dashDot`, `dashDotDot`, `double`, `slantDashDot` |
| `borders[].weight` | `hairline`, `thin`, `medium`, `thick` |
| `fill.pattern` | `none`, `solid`, `automatic`, `gray25`, `gray50`, `gray75`, `gray125` |
| `font.underline` | `none`, `single`, `double`, `singleAccounting`, `doubleAccounting` |
| `colorIndex` | `-4105` automatic, `-4142` none, otherwise a palette index: `1` black, `2` white, `3` red, `4` bright green, `5` blue, `6` yellow, `7` pink, `8` turquoise, and the rest of the 56-entry palette |
| `orientation` | degrees, `-90` to `90`; `255` is stacked vertical text |

</details>

### How to read a CF_HTML payload

`StartHTML` to `EndHTML` bracket the whole HTML document, and
`StartFragment` to `EndFragment` bracket the pasteable slice between the
`<!--StartFragment-->` markers. **Every offset is a byte offset** from the first
byte of the payload, `Version:` included, and the header lines are CRLF
terminated; the header is ASCII, which is also valid UTF-8, because the format
itself is UTF-8 ([Doc] "HTML Clipboard Format" in research 06 §11). The script
records `offsetsConsistent` so a payload whose offsets do not point inside
itself is called out instead of being sliced blindly.

A fragment carrying a full `<html>` document inside the envelope's own
`<html><body>` is expected, not a defect: Chromium wraps whatever a page wrote
with `text/html` in exactly this envelope, so calcplus payloads reach Excel
nested. T01 to T12 test that shape deliberately.

## Probe matrix

`Question` cites research 06 §10 unless it says `context`, which marks a probe
that only supplies the evidence a section 3 or 5 claim rests on.

| Probe | Family | Question | What to look for | Decision rule |
|---|---|---|---|---|
| S01 | source | context §3.1 — the format inventory Excel publishes | Names and order of every clipboard format for a four-cell copy: expect `Biff12`, `XML Spreadsheet`, `HTML Format`, `Rich Text Format`, `Csv`, `Hyperlink` and picture formats ([Obs] in research 06 §3.1) | Context only. Record the list. No browser can request the others, so the contract cannot change |
| S02 | source | context §3.3 and §5 — typed values | `x:num` with the full-precision value in the attribute, `x:str` on numeric-looking text and on a text constant starting with `=`, `x:bool`, `x:err`, and how a line break inside a cell is encoded (`<br>` with `mso-data-placement:same-cell`) | If a typed value arrives without its `x:` attribute, the parser stops trusting element text for typing and re-infers from the format code. The captured spellings become fixture strings |
| S03 | source | **Q1** — is `x:fmla` in multi-cell clipboard HTML? | One `x:fmla` per formula cell, English A1 text with a leading `=`, the cross-sheet reference intact, element text holding the computed value. C3 holds the `=B2*2` of question 2 | Any formula cell without `x:fmla` demotes Excel → calcplus formulas to not supported in `clipboard.md` §3 and in the paste dialog. Every cell carrying it means formulas stay best-effort with a conformance fixture built from this payload |
| S04 | source | context §3.3 and §5 — number formats | The exact `mso-number-format` strings for currency, date, percent, scientific, a custom code with a negative section, and text; note the `\#\,\#\#0` escaping and whether a named format (`Short Date`) is written instead of a code | Emit the escaping Excel writes; if a code arrives unescaped, emit it unescaped. Captured strings become the number-format fixture strings for the HTML writer |
| S05 | source | context §3.4 and §5 — fonts, fills, borders, alignment | CSS per cell: font family, size, weight, style, decoration and colour; `background` plus `mso-pattern`; per-edge border properties and whether diagonals use `mso-border-diagonal-up`; `text-align`, `vertical-align`, `white-space`, `mso-rotate`, indent | The captured CSS is the write contract. Anything Excel writes that `clipboard.md` §4 does not list is either added with its observed spelling or removed from the promise. Solid fills need both `background` and `mso-pattern` |
| S06 | source | context §5 — merges and row heights; feeds **Q4** | `colspan` and `rowspan` on the top-left cell of each merge, covered cells omitted, the merge value present once, `height` and `mso-height-source` on the sized rows | Spans present means the parser maps spans to merges exactly as Excel does and the writer emits the same shape. Spans absent means merges never travel through the clipboard and `clipboard.md` §3 loses the merge promise |
| S07 | source | **Q3** — the source half | `<a href="...">` around the cell text, the URL as the attribute value, and whether the display text is the link text or the URL | If there is no anchor at all, Excel → calcplus can never recover a hyperlink from HTML and the promise stays best-effort. T06 decides the other direction |
| S08 | source | **Q9** — single-cell text copy | Whether the payload has `<table>` and `<td>` at all, whether `<style>` and `ProgId` survive, and whether the fragment is only the value text. An absent or empty HTML flavour is an **observation, not a failure** | No table means the parser needs the no-table path (one cell from the body text) before paste ships, with a fixture from this payload |
| S09 | source | **Q9** — single-cell numeric copy | Whether `x:num` and `mso-number-format` survive a one-cell copy, and whether a table is present | If the format code disappears, single-cell pastes fall back to plain typing and the interface must not claim format fidelity for a one-cell copy |
| S10 | source | **Q9** and **Q1** — single-cell formula copy | Whether `x:fmla` appears for a one-cell copy | If `x:fmla` is present for a range but absent for one cell, warn on single-formula copies or offer paste-all only. Record both spellings as fixtures |
| S11 | source | **Q10** — large range | `htmlByteLength`, `textByteLength`, the `Range.Copy` call time, the clipboard read time, and the time for Excel to paste its own payload back into a clean sheet | The measured size is the payload budget input. If 50,000 cells exceed the roughly 4 MB Windows clipboard history limit (research 06 §6.2), the writer caps the payload at the measured size and cell count and warns instead of truncating silently |
| T01 | target | context §3.4 — typed cells | A1 number `1234.5`, B1 text `00123`, C1 boolean, D1 `#DIV/0!`, E1 text, F1 the **number** `46266` and not a date | A wrong kind means a type attribute is missing for that case. A number for the date-looking cell confirms the no-guessing claim in `clipboard.md` §4 |
| T02 | target | context §3.4 and §6.3 — number formats | `NumberFormat` read back per cell against the code sent: escaped separators, `Short Date`, scientific, the negative section, and text format | Emit the spelling this probe shows Excel accepts. Normalised escapes mean emit normalised and say so in `clipboard.md` §4. A rejected named format means never emit named formats |
| T03 | target | context §3.4 and §6.3 — styles | Font weight, italic, size and colour; fill colour plus pattern; per-edge border style, weight and colour; horizontal and vertical alignment, wrap, rotation | Every property that does not survive leaves the calcplus → Excel promise in `clipboard.md` §4, or its CSS spelling changes to one Excel accepts. Survivors become fixture expectations |
| T04 | target | context §3.4 and §6.3 — spans; the `rowspan` half of **Q4** | `MergeCells` and `MergeArea` for `A1:C1` (`colspan` 3), `A2:A3` (`rowspan` 2) and `A4:B5` (both), with neighbours left unmerged; used range `A1:C5` | All spans merging means `colspan`/`rowspan` stays guaranteed. A working `colspan` with a broken `rowspan` means only single-row merges are promised and multi-row merges are reported as unsupported |
| T05 | target | **Q4** — row heights | The `rowHeights` map for rows 1 to 4: `height=40` (pixels in HTML), `height:60pt` with `mso-height-source:userset`, `height=20`, and no attribute. Default for 11 point Calibri is about 15 points | Pixel heights landing means calcplus converts points to pixels (1 point = 96/72 pixels). Only the point form landing means emit points. Excel auto-fitting everything drops row heights from the calcplus → Excel promise |
| T06 | target | **Q3** — hyperlink import | Per cell `hyperlink.count` and `hyperlink.address`; a `mailto:` target; a cell whose text equals its URL | One hyperlink with the right address keeps `<a href>` in the write contract as best-effort, pinned by a fixture. Text only removes the `a` promise from `clipboard.md` §4 and calcplus emits cell text instead of a URL |
| T07 | target | **Q2** — `x:fmla` anchoring | Formulas read back in F10 to F14 after pasting at F10: verbatim keeps `=B2*2`; a spreadsheet-style rebase shifts by 5 columns and 9 rows to `=G11*2` and `=SUM(G11:G14)`; mixed and absolute references show which part moves | Verbatim means calcplus must write the formula already anchored for the destination or default to paste as values when formulas are present. Rebased means emit source text verbatim and document that pasting at an offset shifts references |
| T08 | target | **Q5** — TSV quoting | The used range shape: quoting honoured gives `A1:D2` with `quoted "inner" quotes`, a two-line cell and a cell containing a tab. Quoting ignored splits row 1 into extra columns or rows | Honoured keeps CSV-style quoting in `clipboard.md` §4 and in the TSV writer. Ignored means the TSV writer refuses embedded tabs and newlines and the HTML flavour carries multi-line cells |
| T09 | target | **Q5** — TSV line endings | `A1:B2` if a bare line feed splits rows, otherwise one row | CRLF is what calcplus writes. If only CRLF splits, add a unit test that asserts no bare line feed reaches the clipboard |
| T10 | target | context §6.1 rule 2 — the metadata control | Whether the number format and fill still arrive without `xmlns:x`, `ProgId` or `Generator`, with `x:num` as an undeclared prefix; compare with T02 and T03 | A control that loses formatting makes `ProgId` and the `x` namespace load-bearing, with a conformance test asserting both. A control that keeps everything leaves them as defence in depth |
| T11 | target | context §3.2 — paste priority | `A1` is `111` if HTML won or `999` if text won, and whether the number format ran at all. This is the real calcplus clipboard shape: HTML plus TSV | Text winning means Excel never sees formats or merges when both flavours are present, so the write path needs an HTML-only mode or a documented Paste Special hint for Excel targets |
| T12 | target | **Q10** — large payload written by calcplus | Generated HTML size, clipboard write time, Excel paste time, pasted cell count, sampled cells. Build time is PowerShell, not the browser, so only the clipboard and paste times compare with the browser measurement | S11 and T12 bracket the payload budget. Excel paste time is the floor a browser paste cannot beat. Needing materially more bytes than S11 for the same cell count means the writer emits too much per cell |

## Browser-side checklist (questions 6 to 8)

These cannot come from the script. Walk them in the Playwright or manual pass
with the built app, and paste into the grid, not into a text field. There is no
clipboard end-to-end spec yet, so the first one should record these checks.

- [ ] **Q6 — Safari paste target.** In current Safari, with focus on the grid
  canvas or its root element (not an input), press Command-V after copying from
  Excel. Record whether a document-level paste event fires. If it does not, the
  hidden-textarea shim from research 06 §7 point 5 is required and
  `clipboard.md` §5 needs correcting: it currently says no focusable element is
  needed. Record the Safari version.
- [ ] **Q6b — Safari with an offscreen textarea.** Repeat with a focused
  offscreen (not `display:none`) textarea and confirm the paste arrives.
- [ ] **Q7 — Firefox HTML fidelity.** Copy a range from Excel on Windows, paste
  into the app in Firefox, and check that `clipboardData.getData('text/html')`
  starts with a full document (`<html`, with `<style>` and the `ProgId` meta),
  not a bare fragment. Record `clipboardData.types` order. Note that Firefox
  writes CF_HTML itself (research 06 §4.1, [Src] bug 586587), so the reverse
  direction should be symmetric — verify, do not assume.
- [ ] **Q8 — Excel for the web as a source.** Copy a range in Excel for the web
  inside Chromium and capture the pasted `text/html`. Expect the same dialect
  plus `data-ccp-*` and `ccp_infra_*` attributes and possibly a second
  `<html><body>` envelope ([Obs] Univer issue 6087 in research 06 §3.3). Confirm
  the parser drops the unknown attributes and keeps values, formats and merges.
- [ ] Record browser versions for all of the above in the report.

Also worth a look while the browser pass is open, though it is not an open
question: paste a range copied from LibreOffice Calc on Linux and confirm the
TSV fallback behaves.

## Report template

Copy this into the ticket or the M5 notes and fill it in. Every verdict must
name the evidence file it rests on.

```markdown
# Excel clipboard spike — result

Manifest: Excel <version> (<build>), <os>, locale <locale>, date system
<1900|1904>, generated <date>. Evidence: <EvidencePath>.

## Answers

| Question | Answer | Evidence | Contract change |
|---|---|---|---|
| Q1 multi-cell `x:fmla` | yes / no / partial | S03, S10 | none / demote formulas in clipboard.md §3 |
| Q2 `x:fmla` anchoring | verbatim / rebased / broken | T07 | export rule / paste-as-values default |
| Q3 `a href` import | hyperlink / text only | S07, T06 | keep / remove the anchor promise |
| Q4 `height` and `rowspan` | honoured / pixels only / points only / auto-fit | T04, T05 | unit conversion / drop row heights |
| Q5 TSV quoting | honoured / ignored | T08, T09 | keep / refuse embedded tabs and newlines |
| Q9 single-cell copies | table / no table | S08, S09, S10 | no-table parser path |
| Q10 large range | <bytes HTML, <bytes text, <ms copy, <ms read, <ms paste for 50,000 cells | S11, T12 | payload cap and warning threshold |
| Q6 Safari paste target | fires / needs shim | manual | clipboard.md §5 |
| Q7 Firefox fragments | full document / fragment | manual | parser |
| Q8 Excel for the web | parsed / needs work | manual | parser |

## Observations worth recording beyond the questions

- Clipboard formats Excel publishes (S01):
- Number-format spellings to copy verbatim (S04, T02):
- CSS properties that survive into Excel and those that do not (S05, T03):
- Fill needs `background` plus `mso-pattern`: yes / no
- `ProgId` and the `x` namespace load-bearing: yes / no (T10)
- Which flavour Excel prefers when both are present: HTML / text (T11)
- Does Excel write `<script>` or event handlers? (baseline for the paste
  sanitiser)

## Fixtures to commit

| Probe payload | Fixture | Test that uses it |
|---|---|---|
| | | |

## Follow-ups

- Spec text to change:
- Conformance fixtures to add:
- Research 06 §10 to mark answered, with the date and Excel build:
- Anything to re-check after the next Office channel update:
```

## What this spike does not settle

- **macOS.** Research question 1 names Windows and macOS. This script is Win32
  only, so the macOS half needs a separate manual capture: copy in Excel for Mac
  with a small clipboard dumper (or paste into a browser page that logs
  `clipboardData`) and compare the `public.html` payload. Mark that half
  `verify` until captured.
- **Excel for the web as a source** is browser-side (Q8) and is only reached
  through a browser copy.
- **One Excel build.** The payload is a de-facto dialect that drifts
  (`clipboard.md` §7). Re-run after an Office channel update and compare the
  manifest build number before trusting an old verdict.
- **Browser timings.** The build times recorded by the script are PowerShell
  string building; only clipboard and paste times are comparable with the
  browser budgets.
- **Column widths.** Research 06 §5 records that Excel ignores widths on paste
  ([Obs]); no probe here re-tests it, and no promise is made.
