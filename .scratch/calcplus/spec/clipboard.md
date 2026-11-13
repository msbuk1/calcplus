# calcplus v1 — clipboard and Excel interop

Resolved by [06 — Clipboard and Excel interop](../issues/06-clipboard-interop.md) on 2026-11-13. Clipboard is a convenience channel, not the fidelity channel; `.xlsx` is the fidelity channel.

## 1. Promise

- Copy and paste inside calcplus preserves values, formulas, manual formatting and merges.
- Copy from Excel into calcplus preserves values, typing, number formats, fonts, fills, borders, alignment and merged cells with high fidelity. Formulas are best-effort.
- Copy from calcplus into Excel preserves the same set, with formulas best-effort.
- For full fidelity in either direction, use `.xlsx` files.

## 2. Formats

**Read:** `text/html` first, `text/plain` (TSV) second, and accept a dropped `.xlsx` or `.csv` as the lossless escape hatch.

**Write:** one `ClipboardItem` containing both `text/html` (Excel dialect) and `text/plain` (TSV, CRLF line endings, CSV-style quoting).

Browsers expose only `text/plain`, `text/html`, `image/png` and optionally `text/rtf`. Excel's `Biff8`, `Biff12` and `XML Spreadsheet` flavours are unreachable and are never a dependency.

## 3. Excel → calcplus

| Content | Fidelity |
|---|---|
| Values and cell typing | High |
| Number formats via `mso-number-format` | High |
| Fonts, fills, borders, alignment | High |
| Merged cells via `colspan`/`rowspan` | High |
| Formulas via `x:fmla` | Best-effort; Excel does not reliably include it in clipboard HTML |
| Comments, validation, conditional formatting, charts, images, sheet structure | Never travel |

## 4. calcplus → Excel

The HTML uses the Excel dialect:

- `xmlns:x` and `ProgId="Excel.Sheet"`.
- `x:num`, `x:str`, `x:bool`, `x:err` and `x:fmla` for typed cells and formulas.
- `mso-number-format` for number formats.
- `colspan` and `rowspan` for merges.
- `<a href>` for hyperlinks.

The TSV fallback carries values only, with CRLF and CSV-style quoting.

Values, manual formatting and merges are guaranteed. Formulas are best-effort because `x:fmla` paste anchoring is undocumented. Column widths, comments and objects are lost.

## 5. Browser mechanics

- Use `copy` and `paste` events, not `navigator.clipboard.read()`.
- Paste events return the raw HTML with its `<style>` context and no permission prompt.
- `read()` sanitises HTML by default and drops `mso-*` and `ProgId`; Chromium's `unsanitized` option is not portable.
- Keep a hidden focused textarea or contenteditable, because Safari and historically Firefox only fire paste in an editable context.
- Handle the async paste path and the permission prompt for write.
- Paste HTML is sanitised: scripts, event handlers, external resources and unsafe URLs are stripped before parsing.

## 6. In-app clipboard

In-app copy also stores:

- a `data-calcplus-*` HTML payload for typed values and formats;
- an in-memory snapshot of the copied range for lossless internal paste.

The in-memory snapshot is the only path that guarantees formulas, styles and merges exactly. It is cleared when the document closes.

## 7. Limits

- Clipboard HTML is a de-facto dialect, not a contract. Excel-specific attributes are legacy, reverse-engineered and drift across Excel and OS builds.
- Formulas in either direction are best-effort.
- Custom MIME formats are Chromium-only.
- Images are `image/png` only and are v2 as cell content.
- The clipboard never becomes a dependency for correctness; if it fails, the user can still use `.xlsx`.

## 8. Acceptance criteria

1. Copy inside calcplus and paste into another calcplus range preserves values, formulas, manual formatting and merges exactly.
2. Copy from Excel into calcplus preserves values, typing, number formats, fonts, fills, borders, alignment and merges.
3. Copy from calcplus into Excel produces values, formats, styles and merges; formulas paste where Excel accepts `x:fmla`.
4. TSV fallback works when HTML is unavailable.
5. Paste HTML is sanitised; no script, event handler or external resource executes.
6. A 100k-cell paste meets the [performance](performance.md) budget.
7. Dropping a `.xlsx` or `.csv` file opens it; dropping an unsupported type shows an error.
8. The About/help text says `.xlsx` is the full-fidelity channel.
