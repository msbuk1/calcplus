# 25 — Structural edits and reference updates: how five open-source spreadsheets do it

Researched 2026-11-19. Sources read at these commits: LibreOffice core `946dbb0312d2` (`sc/` module), OnlyOffice `sdkjs` `9d2138866d92`, IronCalc `e57bbbe21e9c` (main; the `v0.9.1` tag is also compared), Univer `4bb7117487b9`, Grist-core `7efae83c9043`. All claims are from the projects' own source code and are cited to file and symbol. Note the path correction: OnlyOffice's spreadsheet module is `sdkjs/cell/`, not `sdkjs/spreadsheet/`.

**Summary.** One structural edit must move every reference in the file: formulas, defined names, chart series, hyperlinks, conditional formats, data validation, drawing anchors, pivot sources, comments.

1. **One central transform wins.** All four spreadsheet engines funnel insert/delete through one shared range transform (LibreOffice `ScRefUpdate` + `sc::RefUpdateContext`, OnlyOffice `DependencyGraph._shiftMoveDelete`, IronCalc `DisplaceData`/`displaced_node`, Univer `RefRangeService` + `util.ts`). Features register what they own; they do not each invent shifting.
2. **One edit is one undo step everywhere.** LibreOffice wraps the whole document change in one list action; OnlyOffice in one History transaction; IronCalc pushes one diff list per call; Univer pushes one mutation bundle per command; Grist folds formula rewrites into the same user action.
3. **Broken references become `#REF!`, not refusals.** All four engines write `#REF!` for deleted targets. Refusal is reserved for edits that lose data or split an array formula (IronCalc, OnlyOffice, LibreOffice). **No project refuses to protect a hyperlink.** Univer instead re-bases hyperlink targets in the same edit and writes an explicit `range=err` marker when the target is gone.
4. **Multi-part consistency comes from one model, not from patching parts.** LibreOffice, OnlyOffice, Grist and Univer hold everything in one in-memory model and save once. calcplus's per-sheet save pipeline is what forces the current link refusal; the fix is a multi-part save transaction, not a smarter refusal.
5. **Threaded comments** are modelled as a root entry plus a reply chain (`parentId`/`children` or a nested list) in LibreOffice, OnlyOffice, Univer and Grist; the root alone owns the cell anchor.

## 1. Method and evidence standard

Every claim below is labelled **Source** — read in the repository at the commit named above. No claim comes from blog posts or documentation summaries. Where I searched for something and found nothing (for example "who updates internal hyperlink targets in OnlyOffice"), the claim says so and names the search; treat those as strong but negative evidence. Version nuance: IronCalc's clean AST displacement code is newer than the `v0.9.1` tag used in note 24; the tag's older approach is called out where it matters.

## 2. Findings at a glance

| Question | LibreOffice Calc | OnlyOffice sdkjs | IronCalc | Univer | Grist-core |
|---|---|---|---|---|---|
| 1. Reference model | One pass `ScDocument::UpdateReference` over a shared `sc::RefUpdateContext`; per-feature `UpdateReference` methods | Central `DependencyGraph` broadcast; listeners rewrite own tokens; raw-range features have own shift helpers | One AST transform `displaced_node` over every formula; separate string shifts for CF, links, merges | Central `RefRangeService`; shared pure functions in `util.ts`; formulas rewritten by `UpdateFormulaController` | References are columns/tables by ID; row edits cannot break anything; renames patch formula text centrally |
| 2. Undo | One list action (`ScUndoInsertCells` + friends) | One History point + transaction | One `DiffList` per API call | One `IUndoRedoItem` of mutation lists | One user action (rewrites folded in) |
| 3. Un-updatable refs | `#REF!` + info bar; refuses for protection, cells pushed off grid, pivot overwrite | `#REF!` (`cError bad_reference`); refuses for array split, pivot, edge data loss; internal link targets left stale | `#REF!` in AST; refuses for array split, grid-edge; names, comments, link targets left stale | `#REF!` tokens; link target `range=err`; anchor deleted removes link; no refusals found | Text kept, dependents show errors; refuses summary group-by column removal |
| 4. File consistency | Whole document in memory; chart refs and drawings updated in the same pass; save rewrites the file | One JS model holds cells, charts, drawings, comments, pivots; save serialises all parts from the model | Export rebuilds the whole package from a fixed part list; unmodelled parts dropped | One workbook snapshot + plugin resources; one save; no `.xlsx` I/O in repo | One document database; one action bundle |
| 5. Threaded comments | `ScThreadedCommentData`: root + `maReplies` + done flag, on `ScPostIt` | `asc_CCommentData`: root + nested `aReplies`, `bSolved` | None. Flat `Vec<Comment>` per sheet | `IBaseComment` with `threadId`, `parentId`, `children`; root adds `ref` | `_grist_Cells` rows: `root` + `parentId`, type 1 = comment |

## 3. LibreOffice Calc (`sc/` in LibreOffice core)

### 3.1 How references are modelled and updated

One document-level pass, with a shared context, calling per-feature updaters.

- The single-reference transform is `ScRefUpdate::Update` ([`sc/source/core/inc/refupdat.hxx`](https://github.com/LibreOffice/core/blob/946dbb0312d2/sc/source/core/inc/refupdat.hxx)). It returns `ScRefUpdateRes`: `UR_NOTHING`, `UR_UPDATED`, `UR_INVALID` ("some part of the reference became invalid"), `UR_STICKY` (whole-row/column references are deliberately not shifted).
- One operation builds a `sc::RefUpdateContext` (mode, range, deltas) and calls `ScDocument::UpdateReference` ([`sc/inc/refupdatecontext.hxx`](https://github.com/LibreOffice/core/blob/946dbb0312d2/sc/inc/refupdatecontext.hxx), [`sc/source/core/data/documen3.cxx`](https://github.com/LibreOffice/core/blob/946dbb0312d2/sc/source/core/data/documen3.cxx) `ScDocument::UpdateReference`).
- That one pass updates, in order: column/row name ranges, database ranges, defined names (`ScRangeName::UpdateReference`), DataPilot/pivot collections, **chart references** (`ScDocument::UpdateChartRef`), area links, data validation (`ScValidationData::UpdateReference`), detective annotations, a UNO broadcast (`ScUpdateRefHint`), and then every sheet via `ScTable::UpdateReference`.
- `ScTable::UpdateReference` ([`sc/source/core/data/table1.cxx`](https://github.com/LibreOffice/core/blob/946dbb0312d2/sc/source/core/data/table1.cxx)) covers sheet-local names, all formula cells, print and repeat ranges, conditional formats (`ScConditionalFormatList::UpdateReference`) and protected ranges. Drawing objects and note captions are updated by `ScTable::UpdateDrawRef` **after** the rows are inserted, "when the new row heights are known" ([`sc/source/core/data/document.cxx`](https://github.com/LibreOffice/core/blob/946dbb0312d2/sc/source/core/data/document.cxx) `ScDocument::InsertRow`).
- Chart series ranges are updated through `ScChartListener` range lists: `ScDocument::UpdateChartRef` runs `ScRefUpdate::Update` over each listener's ranges and flags "data changed" when a range shrinks or dies ([`sc/source/core/data/documen5.cxx`](https://github.com/LibreOffice/core/blob/946dbb0312d2/sc/source/core/data/documen5.cxx)).

So: one central context and one pass, but each feature owns an `UpdateReference` method that applies the shared `ScRefUpdate` transform. Formula tokens are adjusted per token array (`ScTokenArray::AdjustReferenceOnInsertedTab` etc., [`sc/inc/tokenarray.hxx`](https://github.com/LibreOffice/core/blob/946dbb0312d2/sc/inc/tokenarray.hxx)).

### 3.2 Undo

One user-visible step, several undo objects grouped.

- The command path is `ScDocFunc::InsertCells` → `sc::InsertCellsOperation::run()` ([`sc/source/ui/docshell/docfunc.cxx`](https://github.com/LibreOffice/core/blob/946dbb0312d2/sc/source/ui/docshell/docfunc.cxx), [`sc/source/ui/operation/InsertCellsOperation.cxx`](https://github.com/LibreOffice/core/blob/946dbb0312d2/sc/source/ui/operation/InsertCellsOperation.cxx)).
- `runImplementation` calls `GetUndoManager()->EnterListAction(...)` and adds the undo objects (`ScUndoRemoveMerge`, `ScUndoInsertCells`, …) into that one list action. The comment in the source is explicit: "before insertion all the actions are put in the same ListAction".
- `ScUndoInsertCells` holds an undo `ScDocument` snapshot plus `ScRefUndoData`, and its `DoChange` re-runs `rDoc.InsertRow`/`rDoc.DeleteRow` (which re-run the whole reference pass) ([`sc/source/ui/undo/undoblk.cxx`](https://github.com/LibreOffice/core/blob/946dbb0312d2/sc/source/ui/undo/undoblk.cxx)).
- The UNO entry `ScTableRowsObj::insertByIndex` calls `ScDocFunc::InsertCells` and throws `uno::RuntimeException` when the operation fails ([`sc/source/ui/unoobj/docuno.cxx`](https://github.com/LibreOffice/core/blob/946dbb0312d2/sc/source/ui/unoobj/docuno.cxx)). (`XSheetOperation` itself, implemented on `ScCellRangesObj` in `cellsuno.cxx`, is a range-operation API — apply/aggregate over cells — not a structural-edit API.)

### 3.3 References it cannot update

Break with `#REF!`, notify once; refuse only for data loss or protection.

- A reference whose target is deleted becomes `#REF!`: `ScRefUpdate::Update` returns `UR_INVALID` and the token is marked deleted. The pass records the **first** formula cell that gained a `#REF!` in `RefErrorContext` (`sc/inc/refupdatecontext.hxx`), accumulated in [`sc/source/core/tool/token.cxx`](https://github.com/LibreOffice/core/blob/946dbb0312d2/sc/source/core/tool/token.cxx) `lcl_accumulateRefError`.
- After the pass, `ScDocument::BroadcastRefError` sends `SfxHintId::ScRefErrorCreated` ([`sc/source/core/data/document.cxx`](https://github.com/LibreOffice/core/blob/946dbb0312d2/sc/source/core/data/document.cxx)), and the UI shows an info bar via `ShowRefErrorInfoBar` ([`sc/source/ui/view/tabvwsh5.cxx`](https://github.com/LibreOffice/core/blob/946dbb0312d2/sc/source/ui/view/tabvwsh5.cxx)). So the user is told once, with the address of the first broken cell.
- The whole edit is **refused** when: cells are not editable (protected) — `ScEditableTester::CreateAndTestBlockForAction`; the insert would overwrite a pivot (`STR_NO_INSERT_DELETE_OVER_PIVOT_TABLE`); or a non-empty cell would be pushed past the last row (`ScColumn::TestInsertRow`, comment: "At least one cell would get pushed out. Not good.", [`sc/source/core/data/column.cxx`](https://github.com/LibreOffice/core/blob/946dbb0312d2/sc/source/core/data/column.cxx), [`sc/source/core/data/attarray.cxx`](https://github.com/LibreOffice/core/blob/946dbb0312d2/sc/source/core/data/attarray.cxx) `ScAttrArray::TestInsertRow` — merged-attribute breakage is also refused).

### 3.4 Saved-file consistency

Not a problem in this architecture: the whole document is one in-memory model (`ScDocument` for cells, the draw layer for shapes/charts). One edit runs one pass over that model — including chart ranges and drawing anchors — and saving rewrites the whole file from the model. There is no "save one sheet only" step to keep consistent.

### 3.5 Threaded comments

Modelled directly on the OOXML shape.

- `ScThreadedCommentEntry` ("A single entry in a threaded comment (root or reply)" — maps to `CT_ThreadedComment`), and `ScThreadedCommentData` with `maRoot`, `maReplies` and a done flag ([`sc/inc/postit.hxx`](https://github.com/LibreOffice/core/blob/946dbb0312d2/sc/inc/postit.hxx)). The data hangs off `ScPostIt`, one note per cell; person metadata is `ScPersonData` (`CT_Person`). There is a dedicated `InsertThreadedCommentOperation` ([`sc/source/ui/operation/InsertThreadedCommentOperation.cxx`](https://github.com/LibreOffice/core/blob/946dbb0312d2/sc/source/ui/operation/InsertThreadedCommentOperation.cxx)).

## 4. OnlyOffice Document Server (`sdkjs`, module `cell/`)

### 4.1 How references are modelled and updated

Hybrid: one dependency-graph broadcast for anything that behaves like a formula, plus per-feature shift helpers for raw ranges.

- `DependencyGraph` ([`cell/model/Workbook.js`](https://github.com/ONLYOFFICE/sdkjs/blob/9d2138866d92/cell/model/Workbook.js) ~line 624) registers listeners per cell and per range (`startListeningRange`), plus defined-name listeners. Structural edits call `shift`, `move` or `deleteNodes`, all funnelling into `_shiftMoveDelete`, which collects every listener intersecting the changed box and sends it a notify.
- Each listener is a `parserFormula`; `parserFormula.prototype.shiftCells` ([`cell/model/FormulaObjects/parserFormula.js`](https://github.com/ONLYOFFICE/sdkjs/blob/9d2138866d92/cell/model/FormulaObjects/parserFormula.js) line 11402) walks its parsed tokens (`outStack`) and rewrites the affected `cRef`/`cArea`/`cRef3D`/`cArea3D` boxes. A deleted reference is replaced by `new cError(cErrorType.bad_reference)` — the `#REF!` token.
- Listeners include more than cell formulas: table columns (`TableColumn.prototype.onFormulaEvent`, [`cell/model/WorkbookElems.js`](https://github.com/ONLYOFFICE/sdkjs/blob/9d2138866d92/cell/model/WorkbookElems.js) line 10158), pivot sources (`CT_WorksheetSource.prototype.onFormulaEvent`, [`cell/model/PivotTables.js`](https://github.com/ONLYOFFICE/sdkjs/blob/9d2138866d92/cell/model/PivotTables.js) line 15126), conditional-format formulas (`CConditionalFormattingFormulaParent.prototype.onFormulaEvent`, [`cell/model/ConditionalFormatting.js`](https://github.com/ONLYOFFICE/sdkjs/blob/9d2138866d92/cell/model/ConditionalFormatting.js)), data-validation formulas (`CDataFormula.prototype.onFormulaEvent`, [`cell/model/DataValidation.js`](https://github.com/ONLYOFFICE/sdkjs/blob/9d2138866d92/cell/model/DataValidation.js)), and defined names (registered in `DependencyGraph.defNames`).
- Features that own raw ranges get their own offset helpers called in the same operation: `updateConditionalFormattingOffset`, `updatePivotOffset`, `updateSparklineGroupOffset`, `updateSortStateOffset`, `updateProtectedRangeOffset`, `updateUserProtectedRangesOffset` (all on `Worksheet` in `Workbook.js`; called from `_insertRowsBefore` at line 8550). Merges and hyperlinks shift through `RangeDataManager.prototype.shift` (`WorkbookElems.js` line 7670).
- Chart series are a separate, chart-specific path: `Workbook.handleChartsOnMoveRange` (`Workbook.js` line 5716) collects chart data references via `CChartDataRefs.collectRefsInsideRangeForInsertColRow` and rewrites them through `CChartRefBase.handleOnMoveRange` ([`common/Drawings/Format/ChartFormat.js`](https://github.com/ONLYOFFICE/sdkjs/blob/9d2138866d92/common/Drawings/Format/ChartFormat.js) line 12971). The insert-rows UI action calls it explicitly (see 4.2).

### 4.2 Undo

One History point for the whole multi-feature edit.

- `Range.prototype._shiftUpDown` wraps the edit in `AscCommon.History.Create_NewPoint()` + `StartTransaction()` … `EndTransaction()` and shifts merges and hyperlinks inside that transaction (`Workbook.js` ~line 21360).
- The UI action for InsertRows in [`cell/view/WorksheetView.js`](https://github.com/ONLYOFFICE/sdkjs/blob/9d2138866d92/cell/view/WorksheetView.js) (~line 18940) composes everything in one `functionModelAction`: `t.model.insertRowsBefore`, `cellCommentator.updateCommentsDependencies`, `shiftCellWatches`, `t.model.shiftDataValidation`, freeze-pane update, `handleChartsOnMoveRange`, all between `History.Create_NewPoint()` and `History.EndTransaction()`.
- Undo replays the recorded history items in reverse ([`cell/model/UndoRedo.js`](https://github.com/ONLYOFFICE/sdkjs/blob/9d2138866d92/cell/model/UndoRedo.js); item kinds like `AscCH.historyitem_Worksheet_AddRows` are added in `_insertRowsBefore`).

### 4.3 References it cannot update

`#REF!` plus targeted refusals; internal hyperlink targets are the known stale case.

- Deleted reference → `#REF!` token in every listening formula (4.1). Same for formulas referencing a removed sheet (`parserFormula.removeSheet` path in `processNotify`).
- The edit is **refused** before it runs when: it would split an array formula (`Worksheet.prototype.checkShiftArrayFormulas`, `Workbook.js` line 12678, error `CannotChangeFormulaArray`); it would disturb a pivot (`checkShiftPivotTable` line 12234, error `LockedCellPivot`); or a shift would push merged cells or hyperlinks off the grid edge (`Range.prototype._canShiftUpDown` returns null when the edge range holds merges or hyperlinks).
- **Internal hyperlink targets are not updated.** `Hyperlink` stores `LocationSheet`/`LocationRangeBbox` for the target (`WorkbookElems.js` line 6102), `hyperlinkManager.shift` moves only the anchor box, and no code path rewrites the target (search for `LocationRangeBbox` setters outside `WorkbookElems.js` finds only UI code). A link to `Sheet1!A30` goes stale when rows are inserted above row 30 on Sheet1.

### 4.4 Saved-file consistency

The editor keeps the entire workbook in one JS model: cells, drawings, charts, comments, pivots. Saving serialises every part from that model ([`cell/model/Serialize.js`](https://github.com/ONLYOFFICE/sdkjs/blob/9d2138866d92/cell/model/Serialize.js), `common/SerializeChart.js`); the server-side converter turns that into the `.xlsx`. One edit changes one model inside one History transaction; there is no step that writes one part in isolation.

### 4.5 Threaded comments

`asc_CCommentData` ([`cell/model/CellComment.js`](https://github.com/ONLYOFFICE/sdkjs/blob/9d2138866d92/cell/model/CellComment.js) line 176) holds `wsId`/`nRow`/`nCol` (root anchor), `bSolved`, `oParent`/`nLevel`, and `aReplies` — replies are nested inside the root and serialise as the OOXML threaded comment (`WriteThreadedComment`, `c_oSer_ThreadedComment` with `dT`, `personId`, `text` in `Serialize.js`). Comments move with cells on structural edits via `CCellCommentator.updateCommentsDependencies` (`CellComment.js` line 726).

## 5. IronCalc (Rust core)

Read at `main` (`e57bbbe`, 2026-11-19). The `v0.9.1` tag differs; see 5.6.

### 5.1 How references are modelled and updated

One central AST transform; feature-specific displacement for the rest.

- A parsed formula is a tree of nodes. `displaced_node` ([`base/src/expressions/parser/displace.rs`](https://github.com/ironcalc/IronCalc/blob/e57bbbe21e9c/base/src/expressions/parser/displace.rs)) computes the new tree from `Displacement { data, from, to }` — what moved, and where the formula's cell was and is. A formula that moved along with its references does not change at all (the module doc says exactly this).
- `Model::displace_formulas` ([`base/src/actions.rs`](https://github.com/ironcalc/IronCalc/blob/e57bbbe21e9c/base/src/actions.rs) line 450) walks **every formula on every sheet** once, builds the per-cell `Displacement` from `CellsMoved` (so moved cells displace "from where they were to where they are"), and swaps in the new shared-formula index. Spill cells of array formulas are re-pointed to their moved anchor.
- `Model::insert_rows` / `delete_rows` / `insert_columns` / `delete_columns` (`actions.rs` lines 703–1250) each do, in one method: the cell-data move, row-style shift, hyperlink key remap (`displace_links`, line 64), merged-range remap (`displace_merged_cells`, line 79), the formula pass, and a conditional-formatting pass (`displace_cf_ranges`, `displace_cf_sqref`, `displace_cf_rule_formulas` — CF sqref strings and CF formulas are rewritten at string level).
- Hyperlinks are `Link::External { target }` / `Link::Internal { location }` keyed by cell ([`base/src/types.rs`](https://github.com/ironcalc/IronCalc/blob/e57bbbe21e9c/base/src/types.rs) line 237).

### 5.2 Undo

One step per API call.

- `UserModel::insert_rows` ([`base/src/user_model/common.rs`](https://github.com/ironcalc/IronCalc/blob/e57bbbe21e9c/base/src/user_model/common.rs) line 1069) runs the model change, then pushes one `DiffList` containing the merged-cell snapshot diffs and one `Diff::InsertRows { sheet, row, count }` (or `DeleteRows`, `InsertColumns`, `DeleteColumns` — [`base/src/user_model/history.rs`](https://github.com/ironcalc/IronCalc/blob/e57bbbe21e9c/base/src/user_model/history.rs) lines 130–250). Undo/redo replay diff lists (`apply_undo_diff_list` / `apply_diff_list`, [`base/src/user_model/undo_redo.rs`](https://github.com/ironcalc/IronCalc/blob/e57bbbe21e9c/base/src/user_model/undo_redo.rs)). The replay stores per-row diffs with stable indices to avoid shift drift (doc comment in `insert_rows`).

### 5.3 References it cannot update

`#REF!` in the formula; refusal for array splits and grid-edge; stale for names, comments and link targets.

- A reference whose target is deleted becomes `ErrorKind(Error::REF)` — `#REF!` — in the new AST (`displace.rs`, `Displaced::Gone` handling). A range with one deleted end degrades like Excel: `#REF!:B4` (module doc and `base/src/expressions/parser/tests/test_displace.rs`).
- The edit is **refused** when it would split a CSE array formula (`can_insert_rows` etc., `actions.rs` lines 966–1067; error strings "Cannot insert rows because that would break an array formula") or push cells off the grid ("Cannot shift cells because that would delete cells at the end of a row/column").
- **Left stale** (no updater found in `actions.rs`, verified by search): defined names' references (nothing displaces `parsed_defined_names` on row/column edits; no tests in `base/src/test/test_defined_names_formulas.rs` cover it), cell comments (`Comment.cell_ref` is a string, [`base/src/types.rs`](https://github.com/ironcalc/IronCalc/blob/e57bbbe21e9c/base/src/types.rs) line 479), and internal hyperlink target strings (`Link::Internal.location`; `displace_links` only moves the map key). These look like gaps, not policy.

### 5.4 Saved-file consistency

IronCalc's `.xlsx` export rebuilds the whole package from a fixed part list ([`xlsx/src/export/mod.rs`](https://github.com/ironcalc/IronCalc/blob/e57bbbe21e9c/xlsx/src/export/mod.rs); measured in note 24 §7). Everything the model knows stays consistent by construction — the export derives every part from the one model. Parts the model does not hold (charts, drawings, pivots) are dropped, not updated. So IronCalc "solves" multi-part consistency by not supporting the other parts.

### 5.5 Threaded comments

None. Comments are a flat `Vec<Comment>` per worksheet (`text`, `author_name`, `cell_ref`), imported from the legacy comments part ([`xlsx/src/import/worksheets.rs`](https://github.com/ironcalc/IronCalc/blob/e57bbbe21e9c/xlsx/src/import/worksheets.rs) `load_comments`). No replies, no threading.

### 5.6 Difference at the `v0.9.1` tag

At `v0.9.1` the formula update re-stringified and re-parsed: `shift_cell_formula` called `to_string_displaced` and fed the text back through `update_cell_with_formula` (`git show v0.9.1:base/src/actions.rs`). The current tree replaces that with the AST-level `displaced_node`. Both are one central transform; the new one avoids parse/normalisation churn (relevant to note 24 §6 — the model must keep authored text).

## 6. Univer (TypeScript)

### 6.1 How references are modelled and updated

Central service + shared pure functions; features register the ranges they own.

- `RefRangeService` ([`packages/sheets/src/services/ref-range/ref-range.service.ts`](https://github.com/dream-num/univer/blob/4bb7117487b9/packages/sheets/src/services/ref-range/ref-range.service.ts)) intercepts every structural command (`EffectRefRangeParams` — insert/remove row/col, insert/delete range, move rows/cols/range, reorder; [`type.ts`](https://github.com/dream-num/univer/blob/4bb7117487b9/packages/sheets/src/services/ref-range/type.ts)) through `SheetInterceptorService` and, for each registered range, computes what happens to it.
- The computation is one set of pure functions shared by all features: `handleInsertRow`, `handleIRemoveRow`, `handleMoveRows`, `handleBaseInsertRange`, … ([`packages/sheets/src/services/ref-range/util.ts`](https://github.com/dream-num/univer/blob/4bb7117487b9/packages/sheets/src/services/ref-range/util.ts)). They return `IOperator`s (`Set`, `Delete`, `VerticalMove`, `HorizontalMove`). This is the single "range transform".
- Registered holders include: formulas (`FormulaRefRangeService`, [`packages/sheets-formula/src/services/formula-ref-range.service.ts`](https://github.com/dream-num/univer/blob/4bb7117487b9/packages/sheets-formula/src/services/formula-ref-range.service.ts)), conditional-format formulas, data-validation ranges and formulas, hyperlinks (anchor **and** target), thread comments, notes, tables, filters, merged cells (controllers under each `packages/sheets-*` package).
- Formula text is rewritten separately by `UpdateFormulaController` ([`packages/sheets-formula/src/controllers/update-formula.controller.ts`](https://github.com/dream-num/univer/blob/4bb7117487b9/packages/sheets-formula/src/controllers/update-formula.controller.ts)): it walks the formula data of **all units and sheets**, rewrites reference tokens via the engine's sequence nodes (`LexerTreeBuilder`, `packages/engine-formula`), and covers defined names and super-table (structured) references too.

### 6.2 Undo

One step, built from mutation lists.

- `InsertRowCommand` ([`packages/sheets/src/commands/commands/insert-row-col.command.ts`](https://github.com/dream-num/univer/blob/4bb7117487b9/packages/sheets/src/commands/commands/insert-row-col.command.ts)) builds `redos`/`undos` as lists of mutations, pulls in every feature's contributions from `sheetInterceptorService.onCommandExecute` (`intercepted.redos/undos/preRedos/preUndos` — this is where the formula rewrites and ref-range effects enter), runs them with `sequenceExecute`, and finishes with one `undoRedoService.pushUndoRedo({ undoMutations, redoMutations })`.
- `IUndoRedoItem` is exactly those two mutation lists ([`packages/core/src/services/undoredo/undoredo.service.ts`](https://github.com/dream-num/univer/blob/4bb7117487b9/packages/core/src/services/undoredo/undoredo.service.ts)); undo/redo replay them with `sequenceExecute`. Item groups can merge several pushed items into one step.

### 6.3 References it cannot update

`#REF!` for formulas; explicit `err` marker for link targets; no refusals found.

- Deleted or unmovable formula references become the `ErrorType.REF` token (`#REF!`): in `update-formula.controller.ts` (lines 575, 649, 665) and for removed sheets and removed table columns (`getMovedReferenceToken`, `updateTableReference`). Removed defined names also become `ErrorType.REF` (`_getDefinedNameReferenceToken`).
- Hyperlinks are the showcase: `SheetsHyperLinkRefRangeController` ([`packages/sheets-hyper-link/src/controllers/ref-range.controller.ts`](https://github.com/dream-num/univer/blob/4bb7117487b9/packages/sheets-hyper-link/src/controllers/ref-range.controller.ts)) registers both the anchor cell and the **target range** of an internal link (`payload` of the form `#gid=sheet1&range=A1`). When the target moves, the payload is rewritten in the same edit (`UpdateHyperLinkMutation`); when the target is deleted, the payload becomes `#gid=…&range=err` — an explicit broken marker; when the anchor is deleted, the link is removed (`RemoveHyperLinkMutation`).
- No structural-edit refusal found. The one abort case documented in the code: `RemoveSheetCommand` "aborts its whole sequence on the first failure" (`update-formula.controller.ts` comment) — i.e. transaction abort, not per-reference refusal.

### 6.4 Saved-file consistency

One in-memory workbook (univerjs core document) plus plugin-owned resources (links, comments, CF all persist through resource loaders into the same snapshot). Edits are mutation sequences on that one snapshot; save serialises once. There is **no `.xlsx` reader/writer in this repository** (checked: no package writes `xl/workbook.xml`; xlsx exchange ships as a separate product). So consistency across "file parts" is achieved by not having separate parts.

### 6.5 Threaded comments

`IBaseComment` ([`packages/thread-comment/src/types/interfaces/i-thread-comment.ts`](https://github.com/dream-num/univer/blob/4bb7117487b9/packages/thread-comment/src/types/interfaces/i-thread-comment.ts)) carries `id`, `threadId`, `personId`, `dT`, optional `parentId` and `children`, plus `resolved`. `IThreadComment` adds `ref` (the root's cell anchor; docs add text-range anchors). So: a flat store with parent pointers, root owning the anchor, replies linking by `parentId`. The anchor is kept true by `sheets-thread-comment-ref-range.controller.ts`, which registers the comment's range with `RefRangeService` and also `watchRange`es it, so comments move with their cells.

## 7. Grist-core

### 7.1 How references are modelled and updated

References are columns and tables by ID. Row positions do not exist in the reference model, so row inserts and deletes cannot break a formula.

- Formulas are per **column**, and their text references columns (`$Price`) and tables. The data engine parses formulas into generated Python and records every column/table mention with its text position (`gencode.grist_names()`, [`sandbox/grist/gencode.py`](https://github.com/gristlabs/grist-core/blob/7efae83c9043/sandbox/grist/gencode.py)).
- On a column or table **rename**, `UserActions._prepare_formula_renames` ([`sandbox/grist/useractions.py`](https://github.com/gristlabs/grist-core/blob/7efae83c9043/sandbox/grist/useractions.py) line 973) looks up every recorded mention of the renamed item across all formulas and builds text patches (`textbuilder.make_patch`) to rewrite the affected formula strings. The same rename batch also patches ACL rules, dropdown conditions and trigger conditions (`perform_acl_rule_renames`, `perform_dropdown_condition_renames`, `perform_trigger_condition_renames`).
- Dependencies are tracked at column granularity by a dependency graph ([`sandbox/grist/depend.py`](https://github.com/gristlabs/grist-core/blob/7efae83c9043/sandbox/grist/depend.py) `invalidate_deps`): a column change dirties dependent columns, which recompute.

### 7.2 Undo

One user action is one undo entry, and the formula rewrites ride inside it: `formula_updates` from `_prepare_formula_renames` are folded into the same column-update batch (`_updateColumnRecords`, `useractions.py` line 845) and applied with one `doBulkUpdateFromPairs`. Derived actions are grouped with `indirect_actions()` (same action group). The engine builds the inverse doc-actions as it goes (`docactions.py`: e.g. `RenameColumn` appends the reverse rename for undo, [`sandbox/grist/docactions.py`](https://github.com/gristlabs/grist-core/blob/7efae83c9043/sandbox/grist/docactions.py) line 207).

### 7.3 References it cannot update

- Deleting a referenced column does not rewrite anything: `doRemoveColumns` removes it ([`useractions.py`](https://github.com/gristlabs/grist-core/blob/7efae83c9043/sandbox/grist/useractions.py) line 1489), dependents are invalidated and recompute, and their formulas fail at evaluation — the formula text keeps the stale name. There is no `#REF!` marker in formula text.
- Refusal exists in one narrow place: "RemoveColumn: cannot remove a group-by column from a summary table" (raised as `ValueError`, `useractions.py` line 1485). Renames never break anything, because they are always rewritten.

### 7.4 Saved-file consistency

Grist stores everything — schema metadata, formulas, ACLs, comments, cell data — as rows in one document database. A user action writes all affected rows as one action bundle ([`app/common/ActionBundle.ts`](https://github.com/gristlabs/grist-core/blob/7efae83c9043/app/common/ActionBundle.ts)). There are no separate file parts to keep in sync.

### 7.5 Threaded comments

Comments live in `_grist_Cells` metadata rows with `type == 1`, anchored by `(tableRef, rowId, colRef)`, and threaded via `root` and `parentId` fields ([`sandbox/grist/docmodel.py`](https://github.com/gristlabs/grist-core/blob/7efae83c9043/sandbox/grist/docmodel.py) `MetaTableExtras._grist_Cells.setAutoRemove`). Policy on structural deletes: a comment is auto-removed when its row, column or table is removed, or when its parent thread is gone. Comments are never moved (row IDs are stable; there is nothing to move).

## 8. Design options for calcplus

The goal: "one edit updates every affected sheet and file part in one transaction", plus a policy for references that cannot be updated. Our current behaviour — refuse the edit because the pipeline saves only the edited sheet — has **no precedent** in these five projects. The options below can be combined; the last section names the combination I would take.

### Option 1 — One range transform, a reference-holder registry, one save transaction

Build a single `transformRange(range, edit) -> moved | shrunk | deleted` function (LibreOffice's `ScRefUpdate`, IronCalc's `displaced_node`, Univer's `util.ts` are all this shape). Every feature that owns a reference registers as a holder: formulas, defined names, chart series (anchor **and** range), hyperlinks (anchor **and** target), conditional formats, data validation, drawing anchors, comments, pivot sources. A structural edit asks every holder for its part changes, then saves all changed parts in one transaction and records one undo entry.

- **Trade-offs.** Every new feature must register, or its references silently go stale — that is exactly IronCalc's defined-name/comment gap (§5.3). The registry is one audit point, so "what owns this reference?" is answerable. The save layer must grow from per-sheet to multi-part writes; that is real work but it is the work the problem is made of.
- **Backed by:** Univer (closest fit: `RefRangeService` + interceptor-produced undo/redo), LibreOffice (`RefUpdateContext` + one `UpdateReference` pass), OnlyOffice (`DependencyGraph` broadcast + History transaction), IronCalc (for the formula half).

### Option 2 — Whole-workbook model and one save pass (the deeper fix)

Stop treating the saved file as a set of parts the pipeline patches one at a time. Load every supported part into one document model; a structural edit mutates it; saving rewrites all parts from the model in one pass.

- **Trade-offs.** Consistency for every edit type, not only structural ones; the "other sheet" problem disappears. But it is the biggest refactor, and it collides with the preservation-first rule (ticket 04): a rewrite is only safe for parts the model fully owns. IronCalc is the cautionary tale — its rebuild silently drops charts, drawings and pivots (note 24 §7). OnlyOffice and LibreOffice make it work only because their models genuinely cover everything.
- **Backed by:** LibreOffice, OnlyOffice, Grist.

### Option 3 — Break and mark: `#REF!` plus one notice, refuse nothing

For a reference whose target is deleted: write `#REF!` into formulas and names; give hyperlinks an explicit broken target (Univer writes `range=err`; for `.xlsx` we would write the Excel convention, e.g. `#REF!` in the location); notify the user once per edit with the first affected location.

- **Trade-offs.** Matches Excel and LibreOffice behaviour, so files round-trip and users are not blocked. Needs a written marker convention for links and names, and undo must restore exact text. A "broken" workbook is honest; a refused edit leaves users unable to clean up a template without first deleting links by hand.
- **Backed by:** LibreOffice (`RefErrorContext` + info bar), Univer (`range=err`), OnlyOffice (`cError bad_reference`), and Excel's documented `#REF!` behaviour.

### Option 4 — Refuse only what would lose data or split a structure

Keep refusal, but narrow it to the cases these projects actually refuse: the edit would push cells off the grid (LibreOffice `TestInsertRow`, IronCalc's `LAST_ROW` check), split a CSE array formula (IronCalc `can_insert_rows`, OnlyOffice `checkShiftArrayFormulas`), overwrite a pivot (LibreOffice, OnlyOffice), or hit protected content. Everything else updates or breaks.

- **Trade-offs.** These are the cases where breaking a reference is not enough, because data would be destroyed. But keeping the cross-sheet-hyperlink refusal in this class has no support from any project studied; replace it with Options 1 + 3.
- **Backed by:** IronCalc, OnlyOffice, LibreOffice (each for a subset).

**Suggested combination:** Option 1 for the mechanism, Option 3 as the reference policy, Option 4's narrow pre-checks for safety. Revisit Option 2 only if the `.xlsx` layer moves to a model-owned save anyway.

## 9. Open questions

- **Is refusing an edit (our current behaviour for cross-sheet links) used by any project?** No. Not one of the five refuses to keep a hyperlink, defined name or chart reference true. IronCalc and OnlyOffice refuse only for array splits, pivot damage and grid-edge data loss; LibreOffice adds protection. Univer is the direct counter-example: it updates the link target in the same edit. If calcplus keeps refusing, it should be recorded in the spec as a temporary v1 limit, not as a modelled behaviour.
- **How do projects decide between `#REF!` and refusal?** The observed rule: **refuse** when the edit would destroy data or split a structure the model cannot represent (array formulas, pivots, cells pushed off the grid); **break with `#REF!`** when only the reference dies; **notify once** after breaking (LibreOffice's info bar). A third behaviour exists but looks like oversight rather than policy: **leave stale** — OnlyOffice's internal hyperlink targets, IronCalc's defined names and comments. calcplus should treat "stale" as a bug class to design out, not an option.
- **Hyperlink targets: always re-base, or break only when deleted?** Univer re-bases on move and marks `err` on delete. Excel re-bases too. We should confirm the `.xlsx` convention for a broken internal link (Excel writes `#REF!` in the location) and pin it with a fixture before implementing.
- **Charts and drawings in v1 scope?** If charts are pass-through, structural edits must still rewrite series formulas and drawing anchors inside the preserved parts. Decide which holder in the Option 1 registry owns those bytes, or state explicitly that v1 leaves them stale (and then, like calcplus's link policy today, refuses edits that would strand them).
- **Comments on deleted rows:** Grist deletes them, Excel/LibreOffice/OnlyOffice/Univer move them with the cell. Confirm which behaviour calcplus's comment contract wants; deletion-on-row-delete is the minority choice but is defensible for a record-style document.

## 10. Sources

**LibreOffice core (commit `946dbb0312d2`, 2026-11-19)**

- Single-reference transform `ScRefUpdate` / `ScRefUpdateRes`: https://github.com/LibreOffice/core/blob/946dbb0312d2/sc/source/core/inc/refupdat.hxx
- Update context, `RefErrorContext`: https://github.com/LibreOffice/core/blob/946dbb0312d2/sc/inc/refupdatecontext.hxx
- Central pass `ScDocument::UpdateReference`: https://github.com/LibreOffice/core/blob/946dbb0312d2/sc/source/core/data/documen3.cxx
- `ScDocument::InsertRow` (whole operation, draw-layer pass, chart dirty update): https://github.com/LibreOffice/core/blob/946dbb0312d2/sc/source/core/data/document.cxx
- `ScDocument::BroadcastRefError` / UI `ShowRefErrorInfoBar`: same file; https://github.com/LibreOffice/core/blob/946dbb0312d2/sc/source/ui/view/tabvwsh5.cxx
- Chart references `ScDocument::UpdateChartRef`: https://github.com/LibreOffice/core/blob/946dbb0312d2/sc/source/core/data/documen5.cxx
- Per-sheet pass `ScTable::UpdateReference`, `ScTable::UpdateDrawRef`: https://github.com/LibreOffice/core/blob/946dbb0312d2/sc/source/core/data/table1.cxx
- Refusal checks `ScColumn::TestInsertRow` / `ScAttrArray::TestInsertRow`: https://github.com/LibreOffice/core/blob/946dbb0312d2/sc/source/core/data/column.cxx · https://github.com/LibreOffice/core/blob/946dbb0312d2/sc/source/core/data/attarray.cxx
- Operation class, editability/pivot refusals, list action: https://github.com/LibreOffice/core/blob/946dbb0312d2/sc/source/ui/operation/InsertCellsOperation.cxx
- Undo `ScUndoInsertCells`: https://github.com/LibreOffice/core/blob/946dbb0312d2/sc/source/ui/undo/undoblk.cxx
- UNO entry `ScTableRowsObj::insertByIndex`: https://github.com/LibreOffice/core/blob/946dbb0312d2/sc/source/ui/unoobj/docuno.cxx
- Ref-error accumulation in tokens: https://github.com/LibreOffice/core/blob/946dbb0312d2/sc/source/core/tool/token.cxx
- Threaded comment structs `ScThreadedCommentEntry` / `ScThreadedCommentData` / `ScPersonData`: https://github.com/LibreOffice/core/blob/946dbb0312d2/sc/inc/postit.hxx

**OnlyOffice sdkjs (commit `9d2138866d92`, 2026-10-15)**

- `DependencyGraph`, `_shiftMoveDelete`, `Worksheet._insertRowsBefore`, `checkShiftArrayFormulas`, `checkShiftPivotTable`, `handleChartsOnMoveRange`, `renameDependencyNodes`, `Range._shiftUpDown`/`_canShiftUpDown`: https://github.com/ONLYOFFICE/sdkjs/blob/9d2138866d92/cell/model/Workbook.js
- `parserFormula.processNotify` / `shiftCells` (deleted ref → `cError`): https://github.com/ONLYOFFICE/sdkjs/blob/9d2138866d92/cell/model/FormulaObjects/parserFormula.js
- Insert/delete UI action composition and chart call: https://github.com/ONLYOFFICE/sdkjs/blob/9d2138866d92/cell/view/WorksheetView.js
- Feature listeners: https://github.com/ONLYOFFICE/sdkjs/blob/9d2138866d92/cell/model/PivotTables.js · https://github.com/ONLYOFFICE/sdkjs/blob/9d2138866d92/cell/model/ConditionalFormatting.js · https://github.com/ONLYOFFICE/sdkjs/blob/9d2138866d92/cell/model/DataValidation.js · https://github.com/ONLYOFFICE/sdkjs/blob/9d2138866d92/cell/model/WorkbookElems.js
- Chart refs `CChartRefBase.handleOnMoveRange`, `CChartDataRefs`: https://github.com/ONLYOFFICE/sdkjs/blob/9d2138866d92/common/Drawings/Format/ChartFormat.js
- Hyperlink model (`LocationSheet`, `LocationRangeBbox`) and `RangeDataManager.shift`: https://github.com/ONLYOFFICE/sdkjs/blob/9d2138866d92/cell/model/WorkbookElems.js
- Threaded comments `asc_CCommentData`, `updateCommentsDependencies`: https://github.com/ONLYOFFICE/sdkjs/blob/9d2138866d92/cell/model/CellComment.js
- Serialisation incl. `WriteThreadedComment`: https://github.com/ONLYOFFICE/sdkjs/blob/9d2138866d92/cell/model/Serialize.js
- Undo replay: https://github.com/ONLYOFFICE/sdkjs/blob/9d2138866d92/cell/model/UndoRedo.js

**IronCalc (commit `e57bbbe21e9c`, 2026-11-19; tag `v0.9.1` for the older path)**

- AST displacement `displaced_node`, `Displacement`, `#REF!` results: https://github.com/ironcalc/IronCalc/blob/e57bbbe21e9c/base/src/expressions/parser/displace.rs
- Structural edits `insert_rows`/`delete_rows`/`insert_columns`/`delete_columns`, `displace_formulas`, `displace_cf_*`, `displace_links`, `displace_merged_cells`, refusal checks: https://github.com/ironcalc/IronCalc/blob/e57bbbe21e9c/base/src/actions.rs
- `Link` and `Comment` types, `Worksheet`: https://github.com/ironcalc/IronCalc/blob/e57bbbe21e9c/base/src/types.rs
- Undo diffs `UserModel::insert_rows`, `Diff` enum: https://github.com/ironcalc/IronCalc/blob/e57bbbe21e9c/base/src/user_model/common.rs · https://github.com/ironcalc/IronCalc/blob/e57bbbe21e9c/base/src/user_model/history.rs
- Undo replay: https://github.com/ironcalc/IronCalc/blob/e57bbbe21e9c/base/src/user_model/undo_redo.rs
- Displacement tests (`#REF!:B4` shapes): https://github.com/ironcalc/IronCalc/blob/e57bbbe21e9c/base/src/expressions/parser/tests/test_displace.rs
- Older v0.9.1 string-and-reparse path: https://github.com/ironcalc/IronCalc/blob/v0.9.1/base/src/actions.rs
- Comment import: https://github.com/ironcalc/IronCalc/blob/e57bbbe21e9c/xlsx/src/import/worksheets.rs
- Export part list (whole-package rebuild): https://github.com/ironcalc/IronCalc/blob/e57bbbe21e9c/xlsx/src/export/mod.rs

**Univer (commit `4bb7117487b9`, 2026-11-19)**

- `RefRangeService`: https://github.com/dream-num/univer/blob/4bb7117487b9/packages/sheets/src/services/ref-range/ref-range.service.ts
- Command/effect types: https://github.com/dream-num/univer/blob/4bb7117487b9/packages/sheets/src/services/ref-range/type.ts
- Shared range-transform functions: https://github.com/dream-num/univer/blob/4bb7117487b9/packages/sheets/src/services/ref-range/util.ts
- `InsertRowCommand` composition and `pushUndoRedo`: https://github.com/dream-num/univer/blob/4bb7117487b9/packages/sheets/src/commands/commands/insert-row-col.command.ts
- Undo service `IUndoRedoItem`: https://github.com/dream-num/univer/blob/4bb7117487b9/packages/core/src/services/undoredo/undoredo.service.ts
- Formula rewriting `UpdateFormulaController` (`ErrorType.REF`, defined names, cross-sheet): https://github.com/dream-num/univer/blob/4bb7117487b9/packages/sheets-formula/src/controllers/update-formula.controller.ts
- Formula holder registration: https://github.com/dream-num/univer/blob/4bb7117487b9/packages/sheets-formula/src/services/formula-ref-range.service.ts
- Hyperlink anchor + target re-basing (`range=err`): https://github.com/dream-num/univer/blob/4bb7117487b9/packages/sheets-hyper-link/src/controllers/ref-range.controller.ts
- Thread comments `IBaseComment`/`IThreadComment`: https://github.com/dream-num/univer/blob/4bb7117487b9/packages/thread-comment/src/types/interfaces/i-thread-comment.ts
- Thread-comment anchor tracking: https://github.com/dream-num/univer/blob/4bb7117487b9/packages/sheets-thread-comment/src/controllers/sheets-thread-comment-ref-range.controller.ts
- Other holders: `packages/sheets-conditional-formatting-ui/src/controllers/cf-formula-ref-range.controller.ts`, `packages/sheets-data-validation/src/controllers/dv-ref-range.controller.ts`, `packages/sheets-note/src/controllers/sheets-note-ref-range.controller.ts`, `packages/sheets-table/src/controllers/sheet-table-ref-range.controller.ts` (same commit)

**Grist-core (commit `7efae83c9043`, 2026-11-19)**

- Formula renames `_prepare_formula_renames`, rename handling, refusals, `doRemoveColumns`: https://github.com/gristlabs/grist-core/blob/7efae83c9043/sandbox/grist/useractions.py
- Name positions in formulas: https://github.com/gristlabs/grist-core/blob/7efae83c9043/sandbox/grist/gencode.py
- Dependency graph: https://github.com/gristlabs/grist-core/blob/7efae83c9043/sandbox/grist/depend.py
- Undo doc-actions: https://github.com/gristlabs/grist-core/blob/7efae83c9043/sandbox/grist/docactions.py
- Comment rows and auto-remove policy: https://github.com/gristlabs/grist-core/blob/7efae83c9043/sandbox/grist/docmodel.py
- Action bundle: https://github.com/gristlabs/grist-core/blob/7efae83c9043/app/common/ActionBundle.ts
