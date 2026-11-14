# Development setup

Decided on 2026-11-14 while closing the spec review. These choices are reversible, so they live here rather than in an ADR; the irreversible ones are in [adr/](adr/README.md).

## Layout

```
src/model/      canonical workbook model and range transforms
src/ooxml/      package layer: read, part patch, write
src/engine/     FormulaEngine port, HyperFormula adapter, worker
src/view/       sheet projection (number formats, conditional-formatting overlay)
src/grid/       canvas renderer, panes, DOM editor
src/chrome/     React shell: ribbon, formula bar, sheet tabs, dialogs
src/commands/   command registry, transactions, undo
bench/          corpus generator and microbenchmarks
fixtures/       small committed fixtures; large files are generated
launcher/       Go module with its own go.mod
```

One npm project, folders by area. Revisit workspaces only if the formula worker needs its own build target.

## Toolchain

- **Node 24 LTS**, pinned in `.nvmrc` and `engines`. **npm**; no workspaces.
- **TypeScript** strict, one `tsconfig.json`, path aliases per area.
- **Vite** builds the app and the engine worker; **Vitest** for unit and integration; **Playwright** for end-to-end and visual; **fast-check** for properties. axe-core returns with the v2 accessibility work.
- **ESLint** flat config plus **Prettier**. Two rules carry architecture: the renderer may not import model modules (ADR 0007), and nothing outside the OOXML layer may write part bytes (ADR 0006).
- The licence gate checks dependencies against the allowlist in [licensing](../.scratch/calcplus/spec/licensing.md).

## Local checks

**No test or check runs on GitHub.** GitHub hosts the source and release
downloads only. Run these on the maintainer's PC:

| Check | Command |
|---|---|
| Typecheck, lint, unit tests, licence gate | `npm run check` |
| Formatting | `npm run format:check` |
| Production build | `npm run build` |
| Chromium end-to-end smoke | `npm run test:e2e` |
| Excel 365 value/error goldens | `scripts/oracle/capture-goldens.ps1` (manual) |

Node comes from `.nvmrc`. Install Chromium once with
`npx playwright install chromium`. Visual baselines and performance numbers are
tied to this local machine; results from other hardware are advisory. The Excel
oracle uses the maintainer's own licence, so there is no runner and no recurring
cost.

## Branches

Direct commits to `main` while the project has one maintainer, with the local
checks passing. Move to pull requests with one approval when a second
contributor lands, as [CONTRIBUTING](../CONTRIBUTING.md) requires.

## Targets this setup must meet

- A per-change local check under 15 minutes
  ([testing and conformance §11](../.scratch/calcplus/spec/testing-conformance.md)).
- Save: 2 s at 250k cells, 8 s at 1M
  ([performance §3.7](../.scratch/calcplus/spec/performance.md)).
