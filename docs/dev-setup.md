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

## CI (GitHub Actions)

| Job | Runner | Trigger | Runs |
|---|---|---|---|
| `quality` | `ubuntu-latest` | every push | typecheck, lint, unit, integration, licence gate |
| `e2e` | `ubuntu-latest` | every push | Playwright smoke |
| `oracle` | `windows-latest` | scheduled and manual | Excel 365 golden regeneration and the manual repair check |

Node comes from `.nvmrc`; npm cache is enabled. The self-hosted Windows oracle runner is set up in milestone 1 ([distribution §8](../.scratch/calcplus/spec/distribution.md)); the Excel licence is the maintainer's own, so no recurring cost is incurred.

## Branches

Direct commits to `main` while the project has one maintainer, with CI green. Move to pull requests with one approval when a second contributor lands, as [CONTRIBUTING](../CONTRIBUTING.md) requires.

## Targets this setup must meet

- A pull request's CI under 15 minutes ([testing and conformance §11](../.scratch/calcplus/spec/testing-conformance.md)).
- Save: 2 s at 250k cells, 8 s at 1M ([performance §3.7](../.scratch/calcplus/spec/performance.md)).
