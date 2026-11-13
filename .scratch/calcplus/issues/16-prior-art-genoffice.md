# 16 — Prior art: what calcplus can learn from GenOffice

Status: resolved
Type: research
Blocked by: none

## Question

What can calcplus learn from [genspark-ai/genoffice](https://github.com/genspark-ai/genoffice) — an Apache-2.0, TypeScript/Electron AI office suite with a shipping spreadsheet — about architecture, `.xlsx` fidelity, grid strategy, and testing?

## Answer

**GenOffice independently validates the preservation-first `.xlsx` direction and shows where the real cost sits: shared formulas, structural-edit reference rewriting, and a long tail of Excel edge cases.** It also uses Univer (Apache-2.0) rather than a custom grid, so the buy path deserves a stress test before ticket 05's custom-grid recommendation is treated as final.

- Their gateway patches OOXML parts with JSZip and asserts untouched entries survive byte-identical — the same strategy ticket 04 chose.
- Their `formula-cached-fallback` treats cached values as evidence; our model must own authored formula text plus cached results.
- Their renderer carries dozens of named fixes (merges, borders, number formats, RTL, clipboard anchors); milestones must budget for this tail.
- Their native Rust sidecars are unavailable to a browser-only product; do not copy the Electron architecture.
- Their compatibility corpus plus visual baselines are a template for ticket 13.

Full research: [16-prior-art-genoffice.md](../research/16-prior-art-genoffice.md)
