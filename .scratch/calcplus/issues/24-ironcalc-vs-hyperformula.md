# 24 — IronCalc/WASM vs HyperFormula for the 52-function contract

Status: open
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
