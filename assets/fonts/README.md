# Bundled print fonts

These Open Font License 1.1 files are embedded, subsetted and renamed by the
PDF export (M7 lane 3). `src/print/pdf-fonts.ts` maps workbook families to
them; `src/print/pdf-renderer.ts` renames every subset `CALP+<family>`.

| Files | Family | Maps from | Licence |
|---|---|---|---|
| `Selawik-Regular.ttf`, `Selawik-Bold.ttf` | Selawik | Segoe UI | `Selawik-OFL.txt` (Reserved Font Name **Selawik**) |
| `Carlito-*.ttf` | Carlito | Calibri, Aptos (the pinned non-metric deviation), and the Greek/Cyrillic supplement | `Carlito-OFL.txt` |
| `LiberationSans-*.ttf` | Liberation Sans | Arial, Helvetica | `Liberation-OFL.txt` (Reserved Font Name **Liberation**) |
| `LiberationSerif-*.ttf` | Liberation Serif | Times New Roman, Times | `Liberation-OFL.txt` |
| `LiberationMono-*.ttf` | Liberation Mono | Courier New, Courier | `Liberation-OFL.txt` |

Rules:

- The files ship byte-for-byte as published. Do not edit them.
- Every PDF subset must be renamed (`CALP+Selawik`, `CALP+Carlito`,
  `CALP+LiberationSans`, …) because the embedded subset is a modified version
  and Selawik and Liberation carry Reserved Font Names. `licensing.md` §3.6.
- Selawik has no italic file; the PDF renderer shears the upright face with
  PDFKit's `oblique` option instead of bundling a different family.
- A family that is not in the table above is routed to the 300 DPI canvas
  fallback rather than silently replaced.
