# 24 — IronCalc/WASM vs HyperFormula for the 52-function contract

Status: resolved
Type: research
Blocked by: none

## Question

Ticket 03 chose HyperFormula (GPLv3) before the formula scope was cut to 52 functions and before we learned that GenOffice runs **IronCalc 0.9.1** (Rust, MIT OR Apache-2.0) for recalculation. Re-evaluate the engine choice for calcplus: browser-only, GPLv3, 52 published functions, no backend.

Compare **IronCalc 0.9.1** (including the `@ironcalc/wasm` and `@ironcalc/wasm-xlsx` packages) with **HyperFormula 3.5** (GPLv3 edition) on:

- **Function coverage:** one-by-one against the 52 functions in [`spec/formula-semantics.md`](../spec/formula-semantics.md).
- **Excel semantics:** error values, type coercion, operator precedence, case-insensitive text, wildcards and text criteria, exact/approximate lookups.
- **Dates:** 1900 and 1904 systems, the fake 29 February 1900, serial dates and times. GenOffice reports IronCalc 0.9.1 has no 1904 handling — verify.
- **Formula storage:** can it preserve authored formula text and cached values without rewriting them?
- **`.xlsx` import/export:** does it preserve untouched package parts, or rewrite the workbook? Could it replace part of the preservation-first OOXML layer, or only the calculation?
- **Browser fit:** WASM bundle size, startup, memory, Web Worker compatibility, performance on large sheets.
- **Licensing** (both are GPLv3-compatible) and **maintenance risk**.
- **Known quirks:** GenOffice documents a strict IronCalc importer, `#ERROR!` results, and formatter gaps such as `mmss.0`.

Recommend one of: **keep HyperFormula**, **switch to IronCalc**, or a **hybrid** (for example IronCalc/WASM for authoritative recalculation plus a lightweight interactive engine). State whether ticket 03 should be amended and what the migration risks are.

## Answer

**Keep HyperFormula 3.5.x (GPLv3); do not switch and do not build a hybrid.** Amend ticket 03 to close the re-evaluation and record the compatibility profile.

- **Dates decide it.** IronCalc 0.9.1 has no 1904 handling (no flag in source; `WorkbookSettings` is only `tz`/`locale`) and gets the fake 29 February 1900 wrong: tested `DATE(1900,2,29)=61`, `DAY(59)=27`, `DAY(60)=28`. HyperFormula reproduces both systems with documented config (`leapYear1900: true` + `nullDate: 1899-12-31`, or `nullDate: 1904-01-01`): tested `60`, `28`, `29`, and `DATE(1904,1,1)=0`.
- **IronCalc's `.xlsx` layer is a rewrite, not a preservation layer.** Tested round-trips dropped charts, pivots, drawings, external links, webextensions and custom parts, flattened shared formulas, and destroyed cached values for unsupported functions after a recalc. It can only be a calculation engine — exactly GenOffice's split, which exists because its own gateway preserves and IronCalc only recalculates.
- **HyperFormula's gaps are small and TS-local:** exact-lookup wildcard behaviour (`VLOOKUP` returns the last match, `MATCH` ignores wildcards), `INT(-8.9)`, `CHOOSE` out-of-range, bare `TRUE`/`FALSE` (named-expression workaround) and `#CYCLE!`→`#CIRC!`. None needs a parser fork; they must ship with the engine port.
- **Browser fit favours HyperFormula:** 171 KB gzip vs 720 KB (calc WASM) / 915 KB (calc+xlsx). At 10k rows performance was comparable; IronCalc's real-world model memory is worse (GenOffice reports ~1.2 GB for a 31 MB workbook).
- **Trade-off:** IronCalc is more Excel-faithful on `INT`, `POWER`, `AND` and lookup wildcards, and is MIT/Apache-2.0. But its date failures are contract blockers, and its xlsx layer cannot help ticket 04.
- **Risks of keeping it:** if the wildcard/`INT`/`CHOOSE` shims slip, the 52-function contract is false; 1904 requires building the engine with the workbook's epoch; browser memory at target sheet sizes is still unproven.

Full research: [24-ironcalc-vs-hyperformula.md](../research/24-ironcalc-vs-hyperformula.md)
