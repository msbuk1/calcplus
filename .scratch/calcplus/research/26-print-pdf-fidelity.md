# 26 — Print and PDF fidelity: spike findings

Spike run 2026-11-20 for M7 lanes 2-3. Everything below is measured, not reasoned: throwaway probes driven by Playwright in Chromium, PDFs inspected with Poppler and ImageMagick. The probe lives in `/tmp/opencode/pdf-spike/` (a Vite app plus driver scripts); no product code was touched and nothing is committed. Versions: **Chromium 159.0.8154.6** (Chrome for Testing, Playwright chromium-1255), **@playwright/test 1.65.1**, **PDFKit 0.21.0** (with fontkit 3.0.0, fflate, linebreak, png-js), **Vite 7.3.8** (same as the repo), **Node 24.15.0** (repo engines `>=24 <25`), **Poppler 24.02.0** (`pdfinfo`, `pdffonts`, `pdftotext`, `pdftoppm`, `pdfimages`), ImageMagick 6.9.12-98 for pixel probes.

**Summary.**

1. **The contract's central claim holds.** With "Background graphics" off, a CSS `background-color` is suppressed (prints white) while SVG `fill`, `stroke` and `circle` paint prints unchanged (pixel-probed). Grid fills in the preview must be SVG shapes, never CSS backgrounds.
2. **A stack of one-page SVGs paginates cleanly.** Five stacked SVGs give exactly five pages whether each is 841.89pt tall, 1122.519685px tall (fractional), or 1123px tall (0.48px over). No drift, no blank pages, no slicing. An SVG 5% taller than the page is sliced mid-content across two pages: breaks must come from the planner.
3. **Geometry is exact in both units.** SVG user units of points land at exact point positions (error under 0.3pt = glyph side bearings); CSS px map at exactly 0.75pt/px. But Chromium prints A4 as **594.96 × 841.92pt**, not 595.28 × 841.89: content sized to nominal A4 overflows by 0.32pt and spawns an **empty extra page**. PDFKit uses true A4. This page-box mismatch is the top fidelity risk for "print matches PDF".
4. **PDFKit 0.21.0 bundles under Vite with zero polyfill config.** The browser build imports only `fflate`, `fontkit`, `linebreak`, `png-js` — no `stream`, `buffer` or `process` shims. One chunk: 591.9 KB minified / 235.0 KB gzip. One gotcha: the standard fonts are unregistered in the browser build; pass `font: null` (or call `registerStdFonts`).
5. **Subsetting is on by default and licence-clean.** The embedded subset holds six tables (head, hhea, loca, maxp, glyf, hmtx) and **no name table**; the only font name in the file is the PDF `/BaseFont`.
6. **The subset rename works and is mandatory.** Default name is `BZZZZZ+Selawik-Regular`. A ~25-line seam lands exactly `CALP+Selawik` in all three descriptors; the file stays valid (`pdfinfo`, `pdftotext`, `pdftoppm` all read it). Selawik's OFL carries **Reserved Font Name "Selawik"**, so shipping the default name would violate licensing.md §3.6.
7. **Missing glyphs fail silently.** fontkit maps uncovered code points to glyph 0 with no error; Selawik's `.notdef` is invisible, so text just vanishes. Nothing at render time detects this. Plan-time detection is cheap: `fontkit.hasGlyphForCodePoint` per code point, per font, per page.
8. **Right-to-left needs a script rule, not a coverage rule.** Pure Arabic runs shape and order correctly in PDFKit. Mixed text is wrong: `Total مرحبا 123` renders as `Total مرحبا 123` in PDFKit but `Total 123 مرحبا` in the browser (Unicode bidi). Any run mixing RTL script with Latin or digits should trigger the 300 DPI fallback in v1.
9. **Size check matches the contract.** A 20-page A4 report (714 table rows with fills and borders, header bands, one image) is **56.9 KB** vector (101 ms in the browser) versus **16.3 MB** PNG at 300 DPI (15.1 MB at JPEG q0.92). The contract's "~195 KB vs ~16 MB" is the same shape; our text was lighter. Vector is about 300× smaller and the text stays selectable.
10. **Font pair for the PDF lane: Selawik (primary Latin) + Carlito (fallback), all subsets renamed `CALP+…`.** Selawik is Microsoft's OFL metric clone of Segoe UI (Latin only — no Greek or Cyrillic). Carlito is the OFL Calibri clone (Latin, Greek, Cyrillic) and already covers the contract's pinned Aptos→Carlito deviation. Liberation Sans/Serif/Mono extend it for Arial/Times/Courier New. All three carry OFL Reserved Font Names, so renaming every subset is required.

## 1. Method and evidence standard

Four probes, all under `/tmp/opencode/pdf-spike/`, all re-runnable:

- **`probe-fontkit.mjs` / `probe-pdfkit.mjs`** (Node 24): fontkit coverage and subset internals; PDFKit PDFs with the default name and with the rename seam. Output checked with `pdffonts`, `pdfinfo`, `pdftotext` (also `-bbox`), `pdftoppm`, `pdfimages -list`, and a Python script that decompresses the PDF streams and dumps the embedded subset's tables and name records.
- **`probe-vite/`** (Vite 7.3.8): a browser probe bundling pdfkit + fontkit, generating PDFs in Chromium, plan-time coverage checks, a browser-rendered bidi ground truth, the 20-page report (vector and 300 DPI canvas raster). Driven by `drive.mjs` (Playwright).
- **`print-probe.mjs`** with `public/print/*.html`: the SVG print question. Chromium's print-to-PDF path (`page.pdf`) is the same compositor `window.print()` drives; the dialog's "Background graphics" toggle is the `printBackground` option. Pages were rendered at 72-400 DPI and measured with `pdftotext -bbox` (word boxes in points) and 1×1 pixel probes (`convert -crop 1x1 … txt:-`).
- Licence evidence read from the font binaries themselves (name IDs 0/7/13/14) and `/usr/share/doc/fonts-libersion*/copyright`.

Caveat stated once: `page.pdf` bypasses the print dialog, so dialog-only overrides (paper pick, scale, headers/footers) are out of scope here — that is the contract's existing §5 advisory limit. What is measured is the layout and paint behaviour of the print pipeline itself.

## 2. Q1 — SVG through Chromium's print path

### 2.1 Paint with "Background graphics" off

Probe `a-paint.html`: one CSS div with `background-color: #b0b0b0`, one SVG with a `fill` rect of the same grey, a blue `stroke` rect, and a red circle. Printed with `printBackground: false` and `true`; each output rendered at 72 DPI (1 px = 1 pt) and sampled:

| Element | Background graphics **off** | Background graphics **on** |
|---|---|---|
| CSS `background-color` div | **white (suppressed)** | `#B0B0B0` |
| SVG `fill` rect | `#B0B0B0` (prints) | `#B0B0B0` |
| SVG `stroke` rect | `#0000FF` (prints) | `#0000FF` |
| SVG `circle` fill | `#CC4444` (prints) | `#CC4444` |

The contract's §3 claim is verified: SVG shape paint is not a CSS background and survives the toggle. SVG `<text>` is also selectable in the print output (`pdftotext` extracts `SVG-TEXT`).

**Consequence:** every fill and border that must print has to be SVG shape paint in the preview. A CSS background is a trap: it looks right on screen and disappears on paper for a user with the default toggle.

### 2.2 Page breaks with a stack of one-page SVGs

Probe `b-breaks.html` stacks N block-level SVGs under `@page { size: A4; margin: 0 }`, each with `PAGEMARK-i` at the top and `BOTTOM-i` at 820/841.89 of its height. `pdfinfo` counts pages; `pdftotext` per page checks where the marks landed:

| Mode | SVG height | Pages for N=5 | Marks per page |
|---|---|---|---|
| exact | 841.89pt | 5 | `PAGEMARK-i` + `BOTTOM-i` on page *i* |
| frac | 1122.519685px (297mm at 96dpi) | 5 | intact |
| round-up | 1123px (0.48px over one A4) | 5 | intact |
| round-down | 1122px | 5 | intact |
| tall (N=2) | 883.98pt (5% over) | **4** | `PAGEMARK-1` on p1, `BOTTOM-1` on p2 — **sliced** |

So Chromium snaps each SVG block to page boundaries for sub-pixel overflow: no drift accumulates over five pages and an SVG a fraction of a pixel too tall still owns one page. But an SVG clearly taller than the page is **cut mid-drawing** at the page edge. There is no `break-inside: avoid` rescue for replaced content.

**Consequence:** the planner must emit exactly page-sized SVGs (a small safety shrink of 0.5pt is wise — see 2.3). The contract's "breaks must be computed" is confirmed.

### 2.3 Scaling, DPI and the A4 rounding trap

Probe `c-geometry.html` places text marks and 10×10pt squares at known SVG coordinates in points, under `@page { margin: 12mm }`. `c-px.html` repeats the same layout authored in CSS px. Measured with `pdftotext -bbox`:

- **Points are points.** `MARK-16-20` lands at xMin 49.75pt where 12mm margin (34.02pt) + 16pt − glyph bearing ≈ 49.8pt. `MARK-MID-20` at 297.36pt vs 297.64pt expected minus bearing. Error is under 0.3pt everywhere — that is glyph side bearings, not layout error.
- **CSS px are exactly 0.75pt.** The px probe's mark targeting 16pt/20pt measured xMin 15.996pt and a 20pt baseline minus the 12pt ascent (yMin 9.141pt vs 9.145 predicted). No surprises; no hidden device-pixel scaling in the print path.
- **But Chromium's A4 is 594.96 × 841.92pt** (`pdfinfo` on every print output), while PDFKit's A4 is 595.28 × 841.89pt. Our content, sized to the nominal A4 content box, was 0.32pt wider than Chromium's page — and **both c probes produced a second, completely empty page**. The px variant overflowed the same way (793.7px authored vs a 793.28px-wide page).

**Consequence:** the plan's page box and the preview's SVG size must not be assumed equal. Either author the SVG at `100% × 100%` of the page box, or shrink by a fraction of a point, and pin the preview↔PDF page count with a test. The print output and the PDF export will otherwise differ by ~0.3pt on every edge — small, but exit criterion 1 says "match".

## 3. Q2 — PDFKit in the browser

### 3.1 Bundling under Vite

`vite build` on a page importing `pdfkit` and `fontkit` succeeded on the first try with **no polyfill plugin and no alias config**. Evidence from the package: `pdfkit.browser.mjs` has exactly four imports — `fflate`, `fontkit`, `linebreak`, `png-js` — and grep finds no `Buffer` or `process` global use (the hits are `Uint8Array` and `_pageBuffer` identifiers). PDFKit 0.21's browser build is a first-class target, not a webpack-era leftover. `pdfkit`'s `exports` map routes Vite's browser condition to `js/pdfkit.browser.mjs` automatically.

- Bundle cost: **591.9 KB minified, 235.0 KB gzip** for pdfkit + fontkit + fflate + linebreak + png-js together. The repo's own `fflate ^0.8.2` and pdfkit's `fflate ^0.8.3` dedupe to one copy.
- **Gotcha:** `new PDFDocument()` in the browser throws `Standard font "Helvetica" is not registered. Call registerStdFonts()` — the browser build ships no standard-font data. Two fixes: pass `font: null` and always set an embedded font first (zero bytes; what the probe uses), or import the metrics from `pdfkit/standard-fonts/Helvetica.mjs` and call `registerStdFonts(...)`.
- Fonts must be passed as `Uint8Array` (fetch the `.ttf`); file paths throw in the browser build.

### 3.2 Subsetting

PDFKit subsets TrueType by default (`CIDFontType2`, `Identity-H`, ToUnicode CMap; `pdffonts` shows `emb yes, sub yes, uni yes`). The embedded font program was extracted from the PDF stream and parsed: **six tables only** — head, hhea, loca, maxp, glyf, hmtx. There is **no name table** in the subset (fontkit drops it), so the embedded program itself carries no font name and no Reserved Font Name. The sfnt magic is the old Apple `'true'`; Poppler accepts it. The 20-page report's whole font payload is a few KB.

### 3.3 Renaming the subset to `CALP+Selawik`

PDFKit names subsets `${tag}+${postscriptName}` where the tag is six letters derived from an internal font id (`F1` → `BZZZZZ`, `F2` → `CZZZZZ`; postscriptName for Selawik is `Selawik-Regular`). So the default is `BZZZZZ+Selawik-Regular` — which **carries the reserved name** and is not what licensing.md §3.6 allows.

The probe wraps `EmbeddedFont.embedProgram` (~25 lines): intercept `doc.ref()` during the embed, rewrite `FontName` on the descriptor at `ref.end()`, and replace the returned name used for the two `BaseFont` entries. Result, verified by grepping the raw PDF and by `pdffonts`:

```
/FontName /CALP+Selawik
/BaseFont /CALP+Selawik     (descendant font)
/BaseFont /CALP+Selawik     (Type0 font)
name                    type            encoding    emb sub uni
CALP+Selawik            CID TrueType    Identity-H  yes no  yes
```

The file is valid: `pdfinfo` (PDF 1.3, A4 595.28 × 841.89pt), `pdftotext` (text selectable), `pdftoppm` (renders), `pdfimages -list` fine. The rename works in both the Node and the browser build (the same seam runs in `drive.mjs`).

One quirk: Poppler's `sub` column says **no** for the 4-letter tag, because it only recognises 6-letter subset tags (`XXXXXX+`). The font *is* subsetted. Excel writes 6-letter tags and ISO 32000 asks for 6 uppercase letters; the contract's `CALP+` example has 4. Lane 3 should either pick a 6-letter tag (`CALPLUS+Selawik`) or document the quirk.

Alternative seam, noted for completeness: PDFKit accepts a fontkit font object (`src.layout` duck-typing), so `postscriptName` can be overridden on a wrapped font — but that alone cannot shorten the tag; the `embedProgram` wrapper is the one clean seam that reaches every name.

## 4. Q3 — What cannot embed: CJK, RTL, and the fallback trigger

### 4.1 Missing glyphs fail silently (and invisibly)

`fontkit.layout('A中B')` returns glyph ids `[2, 0, 13]` — the missing CJK char becomes glyph 0 (`.notdef`) with **no error thrown**, at plan or render time. PDFKit embeds and draws it. The tofu PDF rendered at 100 DPI shows:

- Latin with Selawik: renders.
- CJK, Arabic, Hebrew with Selawik: **nothing at all** — Selawik's `.notdef` glyph is empty, so the text silently disappears. (Noto Naskh's `.notdef` is a visible box, so "tofu" is font-dependent.)
- `pdftotext` drops the unmapped characters entirely.

This is the worst failure mode: wrong output, zero signals. The trigger cannot be reactive; it must be a plan-time scan.

### 4.2 Plan-time detection is exact and cheap

`font.hasGlyphForCodePoint(cp)` (the same call PDFKit itself uses internally for WinAnsi) ran per code point against each font, in Node and in the browser, with identical results:

| Text | Selawik | Carlito | Noto Naskh |
|---|---|---|---|
| Latin `A`, `é`, `ß`, `Œ` | covered | covered | — |
| Cyrillic `Ж` (U+0416) | **missing** | covered | — |
| Greek `λ` (U+03BB) | **missing** | covered | — |
| CJK `汉字测试` | **missing** | **missing** | — |
| Arabic `مرحبا` | **missing** | — | covered |
| Hebrew `שלום` | **missing** | — | — |
| Emoji `😀` | **missing** | — | — |

(Selawik's Microsoft Learn page confirms the scope: `dlng:'Latn'` — Latin only.) The check is per font, per text run, per page — exactly the granularity the contract's fallback needs. Cost is one cmap lookup per code point.

### 4.3 RTL: shaping is fine, bidi is not

- Pure Arabic (`مرحبا`) with Noto Naskh in PDFKit renders **connected and correctly ordered** — fontkit reverses RTL runs and applies joining forms. It matches the browser visually.
- Mixed text is wrong. The browser ground truth (canvas `fillText`, saved as `bidi-truth.png`) renders `Total مرحبا 123` as **`Total 123 مرحبا`** — Unicode bidi pulls the digits left of the RTL run. PDFKit renders **`Total مرحبا 123`** (measured by 400 DPI crop of the Arabic+digit cluster: the word and `123` adjacent, digits on the right). fontkit does per-run shaping but no bidi resolution across runs.
- A coverage check cannot see this: every glyph exists. The detection has to be a **script rule**: a text run (or tight cell cluster) mixing an RTL script (Arabic, Hebrew) with Latin letters or digits → that page goes to the raster fallback. Pure-RTL runs are probably fine but keeping them on the fallback in v1 is the conservative choice; v2 can add real bidi.

## 5. Q4 — Size: vector PDF versus 300 DPI raster

One shared report generator (browser): 20 A4 pages — per page a filled header band, 30-36 table rows (714 rows total) with alternating row fills, cell borders on every cell, number and percent columns, a footer; on page 1 one 480×320 PNG chart (embedded at 2,957 B). Same content drawn twice: once through PDFKit, once to a 2480×3508 canvas per page (300 DPI).

| Output | Size | Time (Chromium 159) | Notes |
|---|---|---|---|
| PDFKit vector PDF | **56,912 B (55.6 KB)** | 101 ms | text selectable, `CALP+Selawik` subset, 20 pages |
| 300 DPI PNG | **17,050,710 B (16.3 MB)** | 1,229 ms | 20 × ~850 KB pages |
| 300 DPI JPEG q0.92 | **15,843,678 B (15.1 MB)** | 1,541 ms | ringing around text; 7% smaller than PNG |

The contract's "~195 KB vs ~16 MB" is confirmed in shape: our raster matches their ~16 MB almost exactly; our vector is smaller because the contract's report carried more text. Vector is ~300× smaller, ~12× faster to produce, and searchable; raster is the fallback only. JPEG barely helps on line art and damages text edges — the fallback should emit PNG (or JPEG only for pages that are mostly a photo). PDFKit generation of 20 pages took 0.1 s against the contract's 3 s budget.

## 6. Q5 — Fonts: what the repo uses and what to embed

**Today.** `src/styles.css` uses `'Segoe UI', system-ui, -apple-system, sans-serif` for chrome and `Consolas, 'Courier New', monospace` for code-like surfaces. No font files are bundled yet (`THIRD_PARTY_NOTICES.md`: OFL fonts "not installed yet"). `licensing.md` §3.6 names the intended set: **Selawik, Carlito, Liberation** — OFL-1.1 separate works; "do not modify or subset the fonts under their reserved names. If a modified font is ever needed, rename it." Research/21 adds the mapping rule: Carlito for Calibri, Liberation for Arial/Times, and the pinned Aptos→Carlito non-metric deviation.

**Licence evidence from the binaries** (name IDs 0/7/13 read with a Python name-table dump):

| Font | Licence | Reserved Font Name | Rename on subset? |
|---|---|---|---|
| Selawik | OFL-1.1 | **"Selawik"** (nameID 0: "with Reserved Font Name Selawik") | **required** |
| Carlito | OFL-1.1 | none in the copyright record ("Carlito is a trademark of tyPoland" only) | not strictly required; do it anyway |
| Liberation | OFL-1.1 | **"Liberation"** (`/usr/share/doc/fonts-liberation/copyright`: "with Reserved Font Name Liberation"; the source also RFNs Arimo/Tinos/Cousine) | **required** |

**Recommended pair for the PDF export:**

- **Primary Latin: Selawik** (Regular/Bold/Light/Semilight/Semibold files exist). It is Microsoft's own open-source **metric-compatible replacement for Segoe UI** (Microsoft Learn; ArchWiki metric list), which is the face `styles.css` puts on everything around the grid. Latin only — covers Latin-1 but not Greek, Cyrillic or anything east of it. Subset ships renamed, e.g. `CALP+Selawik`.
- **Fallback: Carlito** (OFL). Metric clone of **Calibri** — the commonest workbook body font — and covers Greek and Cyrillic. It is also the contract's pinned substitute for Aptos, so one file serves both jobs. Subset renamed (`CALP+Carlito`).
- **Extension, same pattern:** Liberation Sans (Arial), Liberation Serif (Times New Roman), Liberation Mono (Courier New) for workbook fonts that name them. Consolas has no metric clone (research/21); substitute Liberation Mono and list it in the compatibility report like Aptos.
- Everything outside these families and scripts (CJK, Arabic, Hebrew, Indic…) goes to the per-page 300 DPI fallback, per §4.

## 7. Recommendations for the PDF lane

### 7.1 Bundling

1. Add `pdfkit@^0.21` as a regular dependency (MIT; the commit carries the dependency-review note per the kickoff's boundary rule 4). No `vite-plugin-node-polyfills`, no aliases: the browser build is native.
2. Import `PDFDocument` from `pdfkit`; construct with `font: null` and always select an embedded OFL font before drawing. Call `registerStdFonts` only if a standard font is ever genuinely needed.
3. Ship the OFL font files as assets (Selawik, Carlito first) with their notices — the licence gate already expects them.
4. Load the PDF renderer with dynamic `import()`: the ~592 KB chunk must not sit in the startup bundle. fflate dedupes with the app's copy.
5. Use `fontkit` (already a pdfkit dependency) in the planner for the coverage scan. It is the same library PDFKit maps glyphs with, so the planner and the renderer cannot disagree.
6. Subset rename: keep PDFKit's default subsetting; wrap `embedProgram` as in §3.3 for the `CALP+` prefix. Decide the tag width (see open questions).

### 7.2 Fallback triggers — what is detected when

| When | Detect what | How | Result |
|---|---|---|---|
| Plan time, per text run | code points the resolved font cannot draw | `fontkit.hasGlyphForCodePoint` per code point | page → 300 DPI raster |
| Plan time, per text run | RTL script mixed with Latin letters or digits | script-class scan of the run | page → 300 DPI raster |
| Plan time, per page | font file missing from the bundle and no metric-clone mapping | family → file table (Carlito/Liberation) | substitute, stay vector; unknown family → raster |
| Render time | font file fails to parse / PDFKit throws | try/catch around font load | page → 300 DPI raster, log once |

Coverage must be **per page** (contract: "limited to pages that need it"). Render-time checks cannot rescue missing glyphs: `.notdef` drawing is silent (§4.1).

### 7.3 Fidelity risks worth a test

1. **Preview vs PDF page box.** Chromium prints A4 as 594.96 × 841.92pt; PDFKit writes 595.28 × 841.89pt (measured). Golden test: one `PrintPlan`, both renderers, identical page count and content placement within 0.5pt. Decide once whether the plan's page box is true A4 (and the preview shrinks 0.05%) or Chromium's box (and the PDF shrinks).
2. **Blank overflow pages.** Nominal-A4-sized content spawns an empty page (measured). Golden test: page count from the plan == pages in the print output == pages in the PDF, for pt-sized and px-sized preview content.
3. **SVG slicing.** An SVG over the page height is cut mid-content (measured). Test the planner's worst case: a print area whose last row band overflows the box by 1pt must not produce a sliced page.
4. **Background-graphics toggle.** A pixel test: print the preview with `printBackground: false` and assert every fill and border still appears (SVG paint) while nothing depends on CSS backgrounds (measured behaviour §2.1).
5. **Silent tofu.** A pixel golden for a CJK/Arabic cell on a fallback page — and a unit test asserting the planner *flags* those runs. "No exception thrown" proves nothing here.
6. **Mixed-bidi runs.** Pixel test per research/21's font corpus: `Total مرحبا 123` in preview and PDF must match (today they do not — §4.3).
7. **Width drift after substitution.** Right/centre-aligned numbers recomputed with `widthOfString` (research/21) — test that anchored edges hold when Aptos→Carlito changes advances.
8. **Subset tag width.** `pdffonts` reports `sub no` for 4-letter tags (§3.3). Either test for the chosen tag convention or accept and document.

## 8. Open questions

- **Tag width:** `CALP+Selawik` (contract example, 4 letters) versus `CALPLUS+Selawik` (ISO 32000's 6 letters, Poppler-clean, Excel-like). The mechanics support either; the contract should pick one.
- **Segoe-UI-styled cells in PDF:** Selawik matches Segoe UI metrics but is Latin-only; a Segoe-UI cell containing Greek or Cyrillic currently falls to the raster fallback unless the planner maps those runs to Carlito (wrong shapes, right widths). Lane 1's font mapping should decide.
- **Chromium paper rounding across platforms:** 594.96 × 841.92pt is what Chromium 159 on Linux produced; other builds may round differently. Pin Playwright/Chromium in e2e and treat real printers as the contract's existing advisory zone.
- **CJK embedding is out of v1** (raster fallback), so Noto `.ttc` face-index handling in fontkit is untested and unneeded for now.

## 9. Sources

Measured artefacts and probe code: `/tmp/opencode/pdf-spike/` (throwaway). Tool versions in the header line; exact commands are the driver scripts (`drive.mjs`, `print-probe.mjs`, `probe-pdfkit.mjs`, `probe-fontkit.mjs`) and the Poppler/ImageMagick invocations quoted in each section. External facts used sparingly: Microsoft Learn, [Selawik typography page](https://learn.microsoft.com/en-us/typography/font-list/selawik) (OFL, RFN "Selawik", `dlng:'Latn'`); ArchWiki, [Metric-compatible fonts](https://wiki.archlinux.org/title/Metric-compatible_fonts) (Selawik↔Segoe UI, Carlito↔Calibri); PDFKit 0.21.0 sources in the installed package (`js/pdfkit.browser.mjs`, `js/pdfkit.js`) and fontkit 3.0.0 (`dist/`). Repo inputs: `.scratch/calcplus/spec/print-pdf.md`, `licensing.md` §3.6, `research/21-print-pdf.md`, `docs/m7-kickoff.md`.
