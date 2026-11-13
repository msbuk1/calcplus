# PROTOTYPE — grid spike: custom canvas vs Univer (throwaway)

This is throwaway code for ticket 17. It is **not production code** and must not be merged into the product tree.

It answers: which grid foundation gets calcplus to Excel fidelity faster — a custom canvas grid or Univer?

- `custom/` — dependency-free canvas grid: four panes, frozen row and column, 1M-row addressable scroll, 100k sparse cells, selection, merges, DOM editing, formula bar, keyboard navigation and an ARIA mirror.
- `univer/` — stock `@univerjs/preset-sheets-core` 1.4.0 with its own chrome, frozen panes, formulas and number formats.

## Run

Custom:

```bash
python3 -m http.server 8091 --directory prototype/grid-spike
# open http://localhost:8091/custom/
```

Univer:

```bash
cd prototype/grid-spike/univer
npm install
npm run dev
# open http://localhost:8092/
```

Production build (for bundle and startup measurement):

```bash
cd prototype/grid-spike/univer
npm run build
npm run preview
# open http://localhost:8093/
```

## Verdict

Recorded in [ticket 17](../../.scratch/calcplus/issues/17-grid-spike-univer.md) and [the grid spike spec](../../.scratch/calcplus/spec/grid-spike.md) on `main`.

**Custom confirmed; Univer rejected; HyperFormula unchanged.**

## Measurements (one Linux machine, Chromium, 1440×900, warm)

| Metric | Custom | Univer |
|---|---:|---:|
| Production bundle, gzip | 7.9 KB | 1.78 MB entry |
| First visible grid, warm | ~125 ms | ~300–400 ms |
| Scroll p95 frame | 16.8 ms | 18.9–20.6 ms |
| Grid cells in a11y tree | 403 | 0 |

Screenshots are in `screens/`; accessibility snapshots are the `*-a11y*.yml` files.
