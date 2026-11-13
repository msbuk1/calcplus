# 06 — Clipboard and Excel interop

Status: resolved
Type: research
Blocked by: none

## Question

How should copy/paste interoperate with Excel and the OS clipboard?

Determine:

- Which clipboard formats Excel emits and consumes (TSV, HTML, RTF, internal binary formats), and what browsers expose.
- What survives a copy from Excel into calcplus: values, formulas, number formats, styles, merged cells, hyperlinks.
- What calcplus must emit for Excel to paste it correctly, including formulas and formats.
- Browser Clipboard API constraints: permissions, async paste, paste events, drag-and-drop.
- The recommended interop contract: which formats to read and write, and the fidelity guaranteed each way.

## Answer

**Read `text/html` first, `text/plain` (TSV) second, and accept `.xlsx`/`.csv` files as the lossless escape hatch; write both HTML and TSV on every copy.** Browsers expose only `text/plain`, `text/html`, `image/png` (plus optional `text/rtf`), so Excel's `Biff8`/`Biff12`/`XML Spreadsheet` flavours are unreachable and must never be a dependency.

- **Excel → calcplus:** values, cell typing, number formats (`mso-number-format`), fonts/fills/borders/alignment, and merged cells (`colspan`/`rowspan`) survive with high fidelity. Formulas survive only if Excel includes `x:fmla` in its clipboard HTML — documented for HTML export, *not* for the clipboard copy — so treat formula paste as best-effort until fixtures prove it. Comments, validation, conditional formatting, charts/images and sheet structure never travel.
- **calcplus → Excel:** emit one `ClipboardItem` with `text/html` (Excel dialect: `xmlns:x`, `ProgId=Excel.Sheet`, `x:num`/`x:str`/`x:bool`/`x:err`/`x:fmla`, `mso-number-format`, `colspan`/`rowspan`, `<a href>`) plus `text/plain` TSV with CRLF and CSV-style quoting. Values/formats/styles/merges are guaranteed; formulas are best-effort (`x:fmla`, and its paste anchoring is undocumented); column widths, comments and objects are lost.
- **Use paste/copy events, not `navigator.clipboard.read()`.** Paste events return raw HTML with its `<style>` context and no permission prompt; `read()` sanitises HTML by default (dropping `mso-*` and `ProgId`) unless Chromium's `unsanitized` option is used. Keep a hidden focused textarea/contenteditable because Safari (and historically Firefox) only fire paste in an editable context.
- **Trade-offs:** HTML is a de-facto dialect, not a contract — Excel-specific attributes are legacy, reverse-engineered, and drift across Excel/OS builds; the contract must be enforced by cross-platform conformance fixtures. Clipboard is a convenience channel, not the fidelity channel; the PRD should say "use .xlsx for full fidelity". Custom MIME formats are Chromium-only, so in-app lossless paste should ride on `data-calcplus-*` HTML attributes plus an in-memory snapshot.

Full research: [06-clipboard-interop.md](../research/06-clipboard-interop.md)
