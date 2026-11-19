# Fixture groups

Each `.mjs` module in this directory adds one fixture group to the corpus that
`npm run fixtures` (`scripts/generate-fixtures.mjs`) writes. The split exists
so parallel lanes can author fixtures without touching a shared script (see
the lane split in `docs/m6-kickoff.md`).

Contract:

- Export `build(context)`, sync or async. Context is
  `{ root, compatDir, expectedDir, manualDir, oracleDir }` with absolute
  paths. Write each fixture where its clause names it (`file-compat.md` §9).
- Reuse the shared OOXML part templates and `packageEntries` from
  `../parts.mjs`. Append new shared helpers to `parts.mjs` only at its end,
  as a reviewed change; put fixture-specific templates in your own module.
- Never edit `generate-fixtures.mjs`. Adding a module here is all the wiring
  there is (it is picked up automatically, sorted by file name).
- Output must be deterministic (fixed zip timestamps, stable ordering), so
  regeneration is byte-identical and fixture diffs are reviewed like code. A
  diff that only updates expectations is flagged.
- Log every written file with `Wrote <path>`.
