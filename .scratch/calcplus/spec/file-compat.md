# calcplus v1 — `.xlsx` compatibility contract

Resolved by [12 — `.xlsx` file-compatibility contract](../issues/12-file-compat-contract.md) on 2026-11-13. It applies the preservation-first approach from [04 — `.xlsx` import/export](../issues/04-xlsx-io.md) to the v1 scope in the [feature matrix](feature-matrix.md). File handling is in [persistence](persistence.md); formula storage is in [formula semantics](formula-semantics.md).

## 1. What the promise means

Three verbs are kept strictly separate:

- **Read** — parse a part into the calcplus model so it can be displayed or edited.
- **Write** — serialize the model into a valid OOXML part.
- **Preserve** — keep a part calcplus does not model, without changing its meaning.

Two consequences:

- **Byte preservation is not semantic preservation.** Copying a chart's XML keeps the chart, but inserting a row above its source range leaves the chart pointing at the wrong cells unless the anchor is transformed.
- **Preserved is not displayed.** A chart can survive open → save while never being drawn. That is a stated v1 cut, not an accident.

**No silent loss.** Any save that would drop or invalidate a preserved construct produces a visible export report. If a structural edit cannot be applied safely, calcplus blocks that specific edit and explains why.

## 2. Tiers

| Tier | Meaning | Guarantee |
|---|---|---|
| **A — modelled and editable** | calcplus reads the construct into its model and can change it. | Semantic round-trip: open → edit → save preserves the meaning, and an independent reader compares equal. |
| **B — preserved, not editable** | calcplus does not model the construct. Its parts are copied verbatim. | Byte-identical parts; relationships still resolve; not rendered or edited, with one exception — a structural edit transforms a chart's series ranges. Anchors transform where possible, otherwise warn or block. |
| **C — rejected** | calcplus cannot read the file safely. | Clear, actionable message. No partial import and no modification. |

## 3. Tier A — modelled and editable

| Feature | Open | Edit | Save |
|---|---|---|---|
| Cell values: number, text, boolean, error, blank | Read and type | Editable | Written with the correct cell type |
| Dates and times | Serial number plus number format | Editable as values | Serial and format preserved; 1900/1904 flag honoured |
| Formulas and cached results | Authored text plus saved result | Editable | Exact text preserved; `<v>` written only when the engine evaluated the cell |
| Number formats | Built-in ID or custom code | Editable | Format code and built-in ID preserved |
| Fonts, fills, borders, alignment, indent, wrap | Read into the style table | Editable | Existing style indices untouched; new styles appended |
| Shrink to fit, text rotation | Read and render | Authoring is v2 | Preserved |
| Merged ranges | Read | Merge and unmerge | Written; borders follow the merged range |
| Row heights, column widths, hidden rows/columns | Read | Resize; hide/unhide is v2 | Preserved; imported hidden state respected |
| Freeze panes, gridline/heading visibility, saved zoom | Read | Editable | Written |
| Defined names (workbook and sheet scope) | Read and evaluate | Read-only for user names in v1 | Preserved; range transforms apply. System names owned by an IN feature — `_xlnm._FilterDatabase` and `_xlnm.Print_Area` — are maintained by that feature |
| Sheet list, order, names, visibility, tab colours | Read | Insert, delete, rename; reorder/hide is v2 | Written; hidden and veryHidden preserved |
| Legacy notes and threaded comments | Read; indicator and popup | Read-only in v1 | Preserved; authors and shapes untouched unless edited |
| Hyperlinks (external and internal) | Read; follow with Ctrl+click | Read-only in v1 | Preserved |
| Data validation | Read and preserve | Not enforced in v1 | Preserved |
| Conditional formatting | Read; evaluate the supported rule types | Read-only in v1 | Rule XML preserved; ranges transformed |
| AutoFilter state | Read | Apply and clear | Written |
| Sheet and workbook protection | Read | Not enforced (deliberate) | Preserved unchanged |
| Print area, orientation, fit-to-width, margins | Read | Editable | Written |
| Theme colours and named styles | Resolve for display | Read-only in v1 | Preserved |
| Rich text | Plain text plus opaque runs | Editing warns and flattens | Untouched runs preserved |

**Style and shared-string tables are append-only.** Existing indices are never renumbered, because cells, charts and conditional formats reference them by integer index. Unused shared-string entries are permitted.

**Dirty parts are patched, not regenerated.** When a modelled part must change, calcplus writes into the original XML: changed rows and cells are rewritten, and every unmodelled child element, attribute, namespace declaration and its order is carried through unchanged. A dirty part is never re-serialised from the model, because the model does not describe `sheetPr`, `extLst`, `legacyDrawing`, `oleObjects`, `phoneticPr`, `ignoredErrors` and the rest. This is what makes "no silent loss" true for edited files, not only for no-op saves.

**When a part does not match the expected shape.** The position-based patch is attempted first. If the worksheet cannot be cut safely, calcplus falls back to a tree rewrite for that part when it is under 16 MB, and refuses the save above 16 MB with a clear message while keeping the recovery copy. The fallback preserves meaning — every unknown element survives — but not byte identity, and the export report names the part and the path taken (ADR 0006).

## 4. Tier B — preserved, not editable

| Construct | Parts | Behaviour |
|---|---|---|
| Charts and chart sheets | `xl/charts/*`, `xl/chartsheets/*`, drawings | Parts and anchors preserved; not rendered in v1; series references are rewritten where they are plain A1 ranges (§4) |
| Images, shapes, drawings, OLE objects, form controls | `xl/drawings/*`, `xl/media/*` | Preserved; anchors transformed on row/column insert/delete where the anchor is modelled |
| PivotTables and pivot caches | `xl/pivotTables/*`, `xl/pivotCache/*` | Definition and cache preserved; never refreshed; output cells render as stored values; source refs are not rewritten |
| External links | `xl/externalLinks/*` | Preserved; never resolved or refreshed; cached values shown |
| VBA projects | `xl/vbaProject.bin` | Preserved in `.xlsm`; never executed or edited; any save invalidates digital signatures |
| Digital signatures | signature parts | Preserved; save warns that the signature is now invalid |
| Slicers, timelines, sparklines | extension parts | Preserved; not displayed |
| Custom XML, third-party extensions, webextensions | `customXml/*`, extension parts | Preserved |
| Custom document properties | `docProps/custom.xml` | Preserved; not editable |
| Tables / ListObjects | `xl/tables/*` | Preserved; structured references are not evaluated in v1; table refs may go stale and are reported |
| `calcChain.xml` | `xl/calcChain.xml` | Not trusted. Deleted when formulas change, with `fullCalcOnLoad="1"` set |

**Structural edits and Tier B.** A single range-transform module rewrites merged ranges, validations, conditional formats, names, comment and hyperlink positions, dimension records and drawing anchors. Chart series references expressed as plain A1 ranges are rewritten by the same module, so inserting a row above a charted range leaves the chart pointing at the intended cells; matching Excel here is v1 scope, because PRD journey 4 depends on it. Three cases are blocked with an explanation rather than left silently wrong: a chart served by a defined name or another non-range source, a pivot source range (pivot caches are never refreshed in v1), and any anchor the module cannot transform. Edits elsewhere on the sheet are allowed.

## 5. Tier C — rejected

| Input | Detection | Message |
|---|---|---|
| Password-encrypted or IRM-protected workbook | CFB magic `D0 CF 11 E0 A1 B1 1A E1`, not a ZIP | "This workbook is encrypted or password-protected. calcplus cannot open it." |
| `.xls` (BIFF8), `.xlsb` (BIFF12), ODS | Extension and container check | "This file format is not supported. Open it in Excel and save as `.xlsx`." |
| Corrupt, truncated or non-ZIP file | ZIP central-directory parse fails | "This file is damaged or is not a valid workbook." |
| Zip bomb or oversized package | Caps in §8 | "This workbook is too large or unusually compressed to open safely." |
| ISO Strict OOXML | Strict namespace detected | "This workbook uses an OOXML variant calcplus does not support yet." |

No Tier C case imports anything partially, and no Tier C case modifies the file.

## 6. Specific behaviours

### 6.1 Formulas without cached values

- **Open:** parse the formula text. If it uses only the 52 published functions and resolvable references, recalculate and show the result.
- If the function is outside the set and there is no cached value, show `#NAME?` and mark the cell "not recalculated".
- If the function is outside the set but a cached value exists, show the cached value marked stale.
- **Save:** write `<v>` only when the engine evaluated the cell **from inputs it also evaluated**. A cell that depends on a stale or unsupported result is never written from that stale input; it keeps its original `<v>` and is marked stale.
- **Recalculation flag:** set `fullCalcOnLoad="1"` whenever any cell is stale, whenever a `<v>` was omitted, or whenever a value was written from an input the engine did not evaluate. A stale value must never reach a file without asking Excel to recalculate, because Excel would present it as current.
- **Shared formulas:** preserve `t="shared"`, `si` and `ref` for untouched groups. Follower text is synthesised at import and flagged as derived, so it is never written back as authored text. Editing one follower writes a plain formula for that cell and leaves the group intact for the rest.
- **Legacy array formulas:** read-only. Preserve `<f t="array" ref>` and the cached result; editing shows a message and does not change the cell.

### 6.2 External links

- Preserved byte-for-byte; never opened, resolved or refreshed.
- Cached values are displayed.
- The compatibility report lists every external link and says it was not refreshed.
- Structural edits do not rewrite external-workbook references; an edit that would require it is blocked.

### 6.3 Macros (`.xlsm`)

- `vbaProject.bin` is preserved byte-identically. Macros are never executed, inspected or edited.
- Saving an `.xlsm` keeps the macro parts and warns that digital signatures are invalidated.
- **Save As `.xlsx` from a macro workbook** warns that macros will be removed and requires explicit confirmation; the default Save keeps `.xlsm`.
- Opening a macro workbook shows a one-time notice that macros will not run.

### 6.4 Pivot caches

- Definitions and caches are preserved; calcplus never refreshes them.
- Pivot output cells render as the stored values in the worksheet.
- The export report marks the pivot as "preserved but not refreshed" after any edit to its source range.
- A structural edit that would change the pivot's source range is blocked with an explanation.

### 6.5 Digital signatures

- Signature parts are preserved, but any save invalidates them.
- The export report warns before writing, and the saved file is reported as having an invalid signature.

### 6.6 Threaded comments

- Parts are preserved and the indicator renders; replies, mentions and resolved state are read-only in v1.

### 6.7 Protection

- Sheet and workbook protection flags are preserved but not enforced. calcplus can create edits Excel would have blocked; this is a documented deviation.
- The compatibility report states that protection is not enforced.

## 7. Surfacing unsupported content

Two reports, never a wall of dialogs:

- **Compatibility report (on open).** A non-modal banner appears only when Tier B content, external links, unsupported functions, macros, signatures or unenforced protection are present. It summarises counts and opens a panel listing each item, its sheet and anchor, and what calcplus will and will not do. With nothing to report, no banner appears.
- **Export report (on save).** Shown only when something needs attention. Every entry is one of:
  - **Unchanged** — preserved part copied byte-for-byte.
  - **Transformed** — a range or anchor was rewritten.
  - **Preserved but stale** — a chart, pivot or external link no longer matches its source.
  - **Removed** — only possible through an explicit format change such as `.xlsm` → `.xlsx`, and only after confirmation.
  - **Warning** — signature invalidated, macros removed, rich text flattened, protection not enforced.

Both reports are read-only and can be copied as text. The open report is stored in the recovery snapshot so it survives a crash.

## 8. Package limits and hostile input

| Limit | Value |
|---|---|
| Compressed file size | 100 MB |
| Total uncompressed size | 1.5 GB |
| ZIP entry count | 20,000 |
| Compression ratio per entry | 200:1 |
| Single XML part | 512 MB |
| Touched cells / formulas | Per [performance](performance.md): 1M / 200k |

Parsing uses `saxes` with entity resolution disabled and DTDs rejected, so external-entity attacks fail closed. A limit breach is a Tier C rejection with the §5 message.

## 9. Acceptance fixtures

The fixture corpus lives in `fixtures/compat/` and is generated by a committed script where practical. Each fixture names the clause it proves.

| ID | Fixture | Proves |
|---|---|---|
| F1 | `identity-minimal.xlsx` | No-op save returns the original bytes |
| F2 | `identity-full.xlsx` | No-op save is byte-identical with every Tier A feature present |
| F3 | `styles-append.xlsx` | Editing formatting appends styles and leaves existing indices unchanged |
| F4 | `shared-strings-append.xlsx` | Editing text leaves existing shared-string indices unchanged |
| F5 | `formula-cached.xlsx` | Untouched formula cells keep their cached values after an edit elsewhere |
| F6 | `formula-uncached.xlsx` | A formula with no `<v>` recalculates on open and saves a value |
| F7 | `unsupported-function.xlsx` | `_xlfn.XLOOKUP` with a cached value shows stale; without one shows `#NAME?` |
| F8 | `array-shared.xlsx` | Legacy array and shared formulas are read-only and preserved |
| F9 | `charts-images.xlsx` | Chart, image and shape parts are byte-identical after an unrelated edit; report lists them; the dirty worksheet part keeps its unmodelled children (`extLst`, `legacyDrawing`, `oleObjects`) and their order; an insert above the charted range rewrites the series reference |
| F10 | `pivot-external.xlsx` | Pivot and external-link parts preserved; a pivot source-range edit is blocked; external-link references are untouched; report marks staleness |
| F11 | `macros.xlsm` | `vbaProject.bin` byte-identical; no execution; `.xlsx` Save As warns |
| F12 | `threaded-comments.xlsx` | Threaded comments and persons preserved; indicator renders |
| F13 | `validation-cf.xlsx` | Validation and every supported conditional-format rule type render and round-trip |
| F14 | `protection-hidden.xlsx` | Protection, hidden/veryHidden sheets and hidden rows/columns preserved, not enforced |
| F15 | `date1904.xlsx` | 1904 dates display correctly and the flag survives |
| F16 | `print-setup.xlsx` | Print area, orientation, fit-to-width and margins are honoured in preview and PDF |
| F17 | `rich-text.xlsx` | Untouched runs survive; editing warns and flattens to plain text |
| F18 | `encrypted.xlsx` | CFB container rejected with the §5 message; file untouched |
| F19 | `corrupt.xlsx` | Truncated package rejected; no partial import |
| F20 | `zip-bomb.xlsx` | Ratio and size caps reject safely |
| F21 | `strict-ooxml.xlsx` | Strict namespace rejected with a clear message |
| F22 | `producers/` | Excel 365, Excel 2019, LibreOffice and Google Sheets exports all open without repair |
| F23 | `names-defined.xlsx` | Workbook and sheet names, quoted sheet references and range transforms |
| F24 | `csv-roundtrip.csv` | RFC 4180 quoting, delimiter defaults, formulas export as displayed values |

**Repair-free** means the output opens in Excel 365 without a repair prompt. Local scripts use openpyxl, LibreOffice and a relationship verifier where installed; the Excel check is a manual local release step.

## 10. Acceptance criteria

1. A no-op open → save returns a byte-identical file.
2. Every Tier A feature round-trips with semantic equality against an independent reader.
3. Every Tier B part survives with identical uncompressed bytes and resolvable relationships after an unrelated edit.
4. Every output opens in Excel 365 without a repair prompt and in LibreOffice without error.
5. A formula's exact text is never rewritten; a cached value is retained only when the engine evaluated the cell from inputs it also evaluated; any stale content sets `fullCalcOnLoad="1"`.
6. No-op save and edited save both leave style and shared-string indices stable; new entries are appended.
7. A structural edit either transforms every range-bearing construct or blocks the edit; it never leaves a silently wrong chart, pivot or name. Chart series ranges are rewritten; a non-range chart source or a pivot source range is blocked and reported.
8. Tier C inputs are rejected with the §5 message, with no partial import and no file modification.
9. The compatibility report appears only when there is something to report, and lists every Tier B item.
10. The export report uses the five categories in §7 and is shown whenever anything is stale, removed or warned about.
11. `.xlsm` macros survive byte-identically and never execute; `.xlsm` → `.xlsx` requires confirmation.
12. Encrypted workbooks, `.xls`, `.xlsb`, ODS, corrupt files, zip bombs and Strict OOXML are rejected safely within the performance budget.
13. The fixture corpus in §9 passes locally, including the producer matrix.
14. No workbook content leaves the machine during open, save or reporting.
15. A dirty modelled part preserves every unmodelled child element, attribute, namespace declaration and its order.
16. The dirty-part and chart-series behaviours are proved by fixtures, not by inspection.
17. A part that cannot be patched falls back to a tree rewrite below 16 MB, and above it the save is refused with a clear message, the file is not touched and the recovery copy is kept.

## 11. What this contract does not promise

- That calcplus understands, renders or edits Tier B objects.
- That a chart whose series comes from a defined name or another non-range source, or a pivot, stays correct after a structural edit; those edits are blocked instead.
- That unsupported functions recalculate; cached values are a snapshot.
- That a digital signature remains valid after any save.
- That protection is enforced.
- That Strict OOXML, `.xls`, `.xlsb`, ODS or encrypted files open in v1.
- CSV fidelity beyond values and text; CSV is a value channel, not a workbook format.
