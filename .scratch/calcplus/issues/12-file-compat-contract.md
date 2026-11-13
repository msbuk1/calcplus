# 12 — .xlsx file-compatibility contract

Status: open
Type: grilling
Blocked by: 01, 04

## Question

What exactly does "high-fidelity `.xlsx` round-trip" promise?

Using the chosen import/export approach (04) and the v1 feature matrix (01), define a testable compatibility contract:

- Features that must survive open → edit → save unchanged
- Features read-only but preserved verbatim
- Features that trigger a warning on open or save
- Features that block opening
- Behavior for formulas without cached values, external links, encrypted files, macros, and pivot caches
- How unsupported content is surfaced to the user

Output: a feature-by-feature matrix plus acceptance fixtures.
