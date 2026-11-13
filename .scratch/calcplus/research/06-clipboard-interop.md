# 06 — Clipboard and Excel interop

Ticket: [06-clipboard-interop.md](../issues/06-clipboard-interop.md)
Date: 2026-11-13
Scope: copy/cut/paste between calcplus and Excel desktop (Microsoft 365, Windows/macOS),
plus the browser Clipboard API constraints that bound what is possible.

Evidence labels used throughout:

- **[Doc]** — stated by a specification or first-party vendor documentation.
- **[Src]** — read from browser engine source code or a browser bug tracker.
- **[Obs]** — observed in a captured clipboard payload or a working implementation
  (third-party, so weaker than [Doc]/[Src]).
- **[Inf]** — inference from the above; flagged where it matters.

---

## 1. Recommendation (short)

1. **The OS clipboard is the interface; HTML is the high-fidelity channel.**
   Read `text/html` first (Excel's "HTML Format"), fall back to `text/plain` (TSV),
   and accept dropped/pasted `.xlsx`/`.csv` **files** as the lossless escape hatch.
   Never attempt to read Excel's native `Biff8`/`Biff12`/`XML Spreadsheet` formats —
   browsers do not expose them and no browser API can request them. (§3, §4)

2. **Write both `text/html` and `text/plain` on every copy.**
   The HTML carries values, number formats, styles, merges and (best-effort) formulas
   in the dialect Excel already understands (`x:num`, `x:str`, `x:fmla`,
   `mso-number-format`, `colspan`/`rowspan`). The TSV keeps every non-Excel target
   (text editors, terminals, other grids) working. (§6)

3. **Use `paste`/`copy` events, not `navigator.clipboard.read()`**, as the primary path.
   Paste events hand over the *raw* HTML with its `<style>` context and are not
   permission-prompted. `read()` re-serialises the fragment through Chromium's
   sanitiser (losing `mso-*` CSS and the `ProgId` meta) unless
   `read({unsanitized:['text/html']})` is used, which is Chromium-only. (§7)

4. **Guarantee this, and say so in the UI:** Excel ⇄ calcplus transfers
   **values, number formats, cell styles, merged cells, and text/cell typing**
   with high fidelity; **formulas are best-effort** and must be proven by a
   conformance fixture before being advertised; **comments, data validation,
   conditional formatting, charts, images, defined names and sheet structure
   never travel through the clipboard** — the `.xlsx` path exists for those. (§5, §6, §9)

5. **In-app copy/paste must not depend on the OS clipboard.** Keep a serialised
   model snapshot for same-tab paste, and embed `data-calcplus-*` attributes in the
   HTML for cross-tab/other-instance lossless paste. Custom `web ` MIME formats
   are Chromium-only and must not be load-bearing. (§8.4)

---

## 2. Platform clipboard mechanics (what "the clipboard" is)

| Platform | Mechanism | What a browser maps onto it |
|---|---|---|
| Windows | Win32 clipboard: numeric `CF_*` formats plus **registered formats** looked up by name (`RegisterClipboardFormat`) | `text/plain` ⇄ `CF_UNICODETEXT`; `text/html` ⇄ `"HTML Format"` (CF_HTML); `text/rtf` ⇄ `"Rich Text Format"`; `image/png`; `Files` ⇄ `CF_HDROP` [Doc: [Clipboard Formats](https://learn.microsoft.com/en-us/windows/win32/dataxchg/clipboard-formats), [Chromium `clipboard_win.cc`](https://chromium.googlesource.com/chromium/src/+/main/ui/base/clipboard/clipboard_win.cc)] |
| macOS | `NSPasteboard` with UTI-style type names | `public.html` → `text/html`; `public.utf8-plain-text` → `text/plain`; `public.rtf` → `text/rtf` [Doc: [NSPasteboard.PasteboardType.html](https://developer.apple.com/documentation/appkit/nspasteboard/pasteboardtype/html); Obs: [Excel clipboard note §1.3](https://github.com/weisberg/agent_tools/blob/main/docs/excel_format_on_clipboard.md)] |
| Linux/X11 | Selections: `CLIPBOARD` (Ctrl+C/V), `PRIMARY` (middle-click), `SECONDARY` | The Clipboard API spec explicitly leaves multi-selection OSes under-specified ("Some OSs contain multiple clipboard (e.g. Linux, 'primary', 'secondary', 'selection'). Define from which of those data is read."). Chromium reads the `kCopyPaste` buffer, i.e. `CLIPBOARD`, not `PRIMARY`. [Doc: [Clipboard API §7.3.2](https://w3c.github.io/clipboard-apis/); Src: `clipboard_host_impl.cc`/`clipboard_win.cc` buffer enums] |

Two consequences that shape the design:

- **No application can ask a browser for an arbitrary OS clipboard format.** The
  web platform exposes only the mandatory types (`text/plain`, `text/html`,
  `image/png`), a small optional set (`text/uri-list`, `image/svg+xml`), and
  `web `-prefixed custom formats. [Doc: [Clipboard API §6.4–6.5](https://w3c.github.io/clipboard-apis/)]
- **Browsers wrap/unwrap platform envelopes for you.** On Windows, Chromium turns
  `text/html` into a CF_HTML document with `Version/StartHTML/EndHTML/StartFragment/EndFragment`
  headers before writing, and strips the header when reading. [Src: `ClipboardWin::WriteHTML`
  calls `clipboard_util::HtmlToCFHtml`; `ClipboardWin::ReadHTMLInternal` calls
  `CFHtmlExtractMetadata`] Do **not** hand-author CF_HTML headers in JS.

---

## 3. What Excel puts on the clipboard, and what it accepts

### 3.1 Formats Excel emits on Copy

Excel writes roughly 20 representations of the same range at once. A captured
Windows dump (Excel 97-era but the set is stable) lists:

```
CF_ENHMETAFILE, CF_METAFILEPICT, CF_BITMAP,
"Biff8", "Biff5", "BIFF4", "Biff3", "Biff",
CF_SYLK, "Wk1", CF_DIF, "XML Spreadsheet", "HTML Format",
CF_UNICODETEXT, CF_TEXT, "Csv", "Rich Text Format",
"Embed Source", "Native"
```
[Obs: [mschaef.com clipboard dump](https://mschaef.com/what_is_in_your_clipboard);
Obs: [ClipboardEnumViewer format list](https://github.com/CarlTBarnes/ClipboardEnumViewer)
— the modern list also includes `Biff12`, `Hyperlink`, `ObjectLink`, `OwnerLink`,
`Link Source`, `Ole Private Data`]

Microsoft documents the clipboard-paste side of this set: BIFF5/BIFF8/BIFF12,
SYLK, DIF, tab-delimited text, CSV, RTF, HTML, pictures, and embedded/linked
objects. [Doc: [File formats that are supported in Excel → "File formats that use
the Clipboard"](https://support.microsoft.com/en-us/office/file-formats-that-are-supported-in-excel-a28ae1d3-6d19-4180-9209-5a950d51b719)]

Notes:

- **XML Spreadsheet (SpreadsheetML 2003)** is a real clipboard flavour with
  `ss:Formula`, `ss:HRef`, `ss:MergeAcross`/`ss:MergeDown`, `ss:StyleID`.
  Microsoft documents its copy/paste semantics explicitly ("When copying and
  pasting data, Excel and the Spreadsheet component only support elements and
  attributes in the ss namespace… Formulas are not adjusted when data is copied
  and pasted"). [Doc: [XML Spreadsheet Reference](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-xp/aa140066(v=office.10));
  Doc: [XML in Excel and the Spreadsheet Component §Copying and Pasting](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-xp/aa140062(v=office.10))]
- **`Biff8`/`Biff12` are Excel's own binary workbook streams.** They are what makes
  Excel→Excel paste lossless (formulas with reference adjustment, comments,
  validation, conditional formats). There is no public specification of the
  *clipboard framing* (it is not the same as an `.xls`/`.xlsb` file container),
  and it is unreachable from a browser. [Inf from the format names + the fact that
  browsers expose no such type; see §4]
- Excel also puts a **screenshot of the range** on the clipboard (`CF_BITMAP`,
  `CF_METAFILEPICT`, `CF_ENHMETAFILE`, and PNG on modern builds). [Obs: SheetJS,
  "Excel for Windows stores a screenshot of the selected cells as an image";
  [SheetJS clipboard docs](https://docs.sheetjs.com/docs/demos/local/clipboard)]
  Not needed for v1; useful later for "Copy as Picture".

### 3.2 Format matrix (Excel side)

Legend: ✅ yes, ◐ conditional, ❌ no.

| Clipboard format | Excel emits | Excel consumes on paste | Browser exposes | Usable by calcplus |
|---|---|---|---|---|
| `CF_UNICODETEXT` / `CF_TEXT` (TSV) | ✅ displayed text, tab/CRLF separated | ✅ (as Text / Unicode Text) | ✅ as `text/plain` | ✅ fallback channel |
| `"HTML Format"` (CF_HTML) / `public.html` | ✅ full HTML doc + `<style>` | ✅ (documented: "When you copy text from another program, Excel pastes the text in HTML format") | ✅ as `text/html` | ✅ **primary channel** |
| `"Rich Text Format"` / `public.rtf` | ✅ formatted table, values only | ✅ ("Formatted text") | ◐ Chrome (Win/mac) and Safari may expose `text/rtf`; Firefox generally does not | ◐ optional fallback, not needed |
| `"XML Spreadsheet"` | ✅ SpreadsheetML 2003 | ✅ | ❌ | ❌ |
| `"Biff12"` / `"Biff8"` / `"Biff5"` | ✅ | ✅ | ❌ | ❌ |
| `"Csv"`, `CF_SYLK`, `CF_DIF`, `"Wk1"` | ✅ | ✅ | ❌ (the browser's `text/plain` comes from CF_UNICODETEXT, not these) | ❌ |
| `"Hyperlink"`, `"OwnerLink"`, `"ObjectLink"` | ◐ | ✅ | ❌ | ❌ (hyperlinks can still travel inside HTML via `<a href>`) |
| `"Embed Source"`, `"Native"`, OLE object | ✅ (whole-workbook OLE payload) | ✅ | ❌ | ❌ |
| Bitmap/metafile/PNG of the range | ✅ | ✅ (as picture) | ✅ `image/png` | ◐ could support "paste as image" later |
| `CF_HDROP` (files) | ❌ (not for cell copies) | ✅ | ✅ `Files` | ✅ for `.xlsx`/`.csv` file paste/drop |

**Honest gap:** Microsoft documents the *pasteable* set but not the exact
priority order or the clipboard binary framing. Third-party controls implement a
priority of `Biff8 > Biff12 > Biff5 > Csv > TabDelimited > DrawingObject`
[Obs: [DevExpress `PasteFormat`](https://docs.devexpress.com/OfficeFileAPI/DevExpress.XtraSpreadsheet.SpreadsheetClipboardOptions.PasteFormat)].
For a browser the order is irrelevant: only HTML/TSV/files reach us.

### 3.3 The HTML flavour Excel writes (and reads)

Captured Excel (365/2024, Windows) clipboard HTML has this shape:

```html
<html xmlns:o="urn:schemas-microsoft-com:office:office"
      xmlns:x="urn:schemas-microsoft-com:office:excel"
      xmlns="http://www.w3.org/TR/REC-html40">
<head>
  <meta name=ProgId content=Excel.Sheet>
  <meta name=Generator content="Microsoft Excel 15">
  <style><!--
    td { mso-number-format:General; font-size:11.0pt; ... }
    .xl65 { mso-number-format:"\#\,\#\#0"; }
  --></style>
</head>
<body>
<table>… <td class=xl65 x:num="516382306730">516.382.306.730</td> …</table>
</body>
</html>
```
[Obs: captured payloads in [Qiita (CF_HTML 1.0, offsets, `mso-number-format`)](https://qiita.com/kob58im/items/f5d6b8d93467ef97e9c6)
and [Univer issue #6087 (Excel 2024 + Excel for the web, single-cell copy)](https://github.com/dream-num/univer/issues/6087);
Doc for the dialect: [Microsoft Office HTML and XML Reference](https://learn.microsoft.com/en-us/previous-versions/office/developer/office2000/aa155477(v=office.10))]

The `x:` attributes are Excel's typed-value hints, documented for Excel's HTML
export:

| Attribute | Meaning | Example |
|---|---|---|
| `x:num` | cell is numeric; value is the **full-precision** number when the display text is rounded/formatted | `<td x:num="12344.6789">12,345</td>` |
| `x:str` | cell is text even if it looks numeric | `<td x:str>00123</td>` |
| `x:bool` | boolean | `<td x:bool="TRUE">TRUE</td>` |
| `x:err` | error value | `<td x:err="#DIV/0!">#DIV/0!</td>` |
| `x:fmla` | cell formula (English, A1 style, `=` prefix) | `<td x:fmla="=SUM(B2:B5)" x:num>10</td>` |

[Doc: ["Formulas and Data Types" (Microsoft Office HTML reference mirror)](https://stigmortenmyre.no/mso/html/excel/xlconformulas.htm);
Doc: ["Cell Formatting" — `x:num` + `mso-number-format`](https://stigmortenmyre.no/mso/html/excel/xlconcellformatting.htm)]

Key observed properties of Excel's HTML:

- **Number formats ride in CSS**, as `mso-number-format:"<Excel format code>"` on a
  class or inline style. This is how a formatted value (`$1,234.56`, `mm/dd/yyyy`,
  `0.00%`) is distinguishable from its raw number. [Doc: same "Cell Formatting"
  page; Obs: both captured payloads above]
- **Merged cells use ordinary `colspan`/`rowspan`.** [Obs: [Clipboard Inspector](https://clipboardinspector.com/blog/parse-pasted-spreadsheet-data);
  Obs: [Excel HTML reverse-engineering notes §6.6](https://github.com/weisberg/agent_tools/blob/main/docs/excel_format_on_clipboard.md)]
- **A single-cell copy can degrade to a bare text run** — no `<table>`, only a
  style block and the value text. Any parser must handle "no table ⇒ one cell".
  [Obs: Univer #6087 Excel 2024 payload]
- **Excel for the web writes the same dialect but adds `data-ccp-*` /
  `ccp_infra_*` attributes** and may wrap the document inside another
  `<html><body>` envelope. [Obs: Univer #6087; [reverse-engineering notes §15](https://github.com/weisberg/agent_tools/blob/main/docs/excel_format_on_clipboard.md)]
- Excel's HTML export writes `x:fmla` for formula cells ([Doc] for the attribute
  and its export semantics). **Whether the *clipboard copy* always includes it is
  not documented by Microsoft**; several independent implementers depend on it
  ([Obs: svgrid](https://svgrid.com/docs/help/cells/paste-find-structure),
  [Obs: ooxml-excel-editor `clipboard-html.ts`](https://github.com/ojaDeveloper/ooxml-excel-editor/blob/main/src/core/edit/clipboard-html.ts)).
  Treat Excel→calcplus formula transfer as **verify-before-promise** (see §10).

### 3.4 What Excel accepts from HTML

Microsoft: "When you copy text from another program, Excel pastes the text in HTML
format, regardless of the format of the original text," and HTML is listed as a
pasteable clipboard format. [Doc: File formats that are supported in Excel, above]
The reverse-engineering evidence (plus long-standing practice) is that Excel's
HTML parser honours the same `x:` attributes and `mso-*` CSS when they are present,
and that it only takes the "native Excel HTML" path when the document declares the
Excel namespaces and `ProgId=Excel.Sheet`. [Obs: reverse-engineering notes §2.2,
§8; corroborated by vendor guidance to emit `<td x:num>` and
`mso-number-format` [StackOverflow](https://stackoverflow.com/questions/354476/html-to-excel-how-can-tell-excel-to-treat-columns-as-numbers/357106)]

Treat the following as the **write contract** (§6) — these are the levers that are
documented or widely reproduced:

- `xmlns:x="urn:schemas-microsoft-com:office:excel"` + `<meta name=ProgId content=Excel.Sheet>`
- `x:num` / `x:str` / `x:bool` / `x:err` / `x:fmla` on `<td>`
- `mso-number-format` in a `<style>` class or inline
- `colspan`/`rowspan` for merges
- `<a href="…">` for hyperlinks (best-effort; see §5)
- HTML `height`/`width` and `mso-*` properties for row height, wrap, rotation,
  fill pattern, borders

---

## 4. What browsers actually expose

### 4.1 Formats visible to web content

| Source OS format | Chrome/Edge (Win) | Chrome (macOS) | Firefox (Win/mac) | Safari (macOS) | Firefox/Chrome (Linux, LibreOffice source) |
|---|---|---|---|---|---|
| Unicode text | `text/plain` | `text/plain` | `text/plain` | `text/plain` | `text/plain` |
| HTML Format / public.html | `text/html` | `text/html` | `text/html` | `text/html` | `text/html` |
| Rich Text Format / public.rtf | `text/rtf` | `text/rtf` | ❌ (historically only `text/plain` leaked through) | `text/rtf` (listed even when HTML is not) | `text/rtf` if offered |
| PNG/bitmap picture | `image/png` | `image/png` | `image/png` | `image/png` | `image/png` |
| Excel native (Biff/XMLSS/OLE) | ❌ | ❌ | ❌ | ❌ | ❌ |

Sources:
- Mandatory types are `text/plain`, `text/html`, `image/png`; optional are
  `text/uri-list`, `image/svg+xml`, `web ` custom formats. [Doc: [Clipboard API §6.4–6.5](https://w3c.github.io/clipboard-apis/)]
- Chromium's Windows `ReadAvailableTypes` maps `CF_RTF` → `text/rtf`, PNG, and
  `text/uri-list` alongside text/html. [Src: [`ui/base/clipboard/clipboard_win.cc`](https://chromium.googlesource.com/chromium/src/+/main/ui/base/clipboard/clipboard_win.cc)]
- Firefox's `clipboardData.types` has historically hidden RTF (only `text/plain`
  appeared); the same bug notes Safari lists `text/rtf` and can omit `text/html`
  when the source is RTF-only. [Src: [Mozilla bug 1254028](https://bugzilla.mozilla.org/show_bug.cgi?id=1254028)]
- Firefox on Windows writes `CF_HTML` for `text/html` copies (interop fix in bug
  586587), so the HTML channel is symmetric in both directions. [Src: [Mozilla bug 586587](https://bugzilla.mozilla.org/show_bug.cgi?id=586587)]
- SheetJS's tested matrix for browser *reading* is `text/plain`, `text/html`,
  `text/rtf`; for *writing*, `text/rtf` is not supported. [Obs/testimonial:
  [SheetJS clipboard docs](https://docs.sheetjs.com/docs/demos/local/clipboard)]

### 4.2 Browser API support (async clipboard)

From MDN's browser-compat-data (authoritative mirror of MDN tables):

| Feature | Chrome | Firefox | Safari |
|---|---|---|---|
| `navigator.clipboard` | 66 | 63 | 13.1 |
| `readText()` | 66 (needs `clipboard-read` permission) | 125 (user gesture + paste prompt) | 13.1 |
| `read()` | 76 (permission) | 127 (gesture + prompt) | 13.1 |
| `write()` | 76 (gesture required from 107) | 127 (gesture) | 13.1 |
| `writeText()` | 66 (gesture from 107) | 63 (gesture) | 13.1 |
| `text/html` read/write | 86 | 127 | 13.1 |
| `ClipboardItem.supports()` | 121 | 127 | 18.4 |
| `web ` custom formats | 104 | preview only | ❌ |

[Doc: [BCD `api/Clipboard.json`](https://github.com/mdn/browser-compat-data/blob/main/api/Clipboard.json),
[`api/ClipboardItem.json`](https://github.com/mdn/browser-compat-data/blob/main/api/ClipboardItem.json);
rendered: [MDN Clipboard.read](https://developer.mozilla.org/en-US/docs/Web/API/Clipboard/read)]

Practical reading of that table:

- **Writes are fine everywhere modern** as long as they happen in a user gesture
  (click, keydown, or the `copy`/`cut` event). Chrome 107+ only requires transient
  activation, not a persistent permission. Safari always requires the gesture.
- **`read()` is the fragile one**: it may prompt, and Firefox/Safari only allow it
  from a gesture. There is no way to silently poll the clipboard. Do not build
  Ctrl+V on it.
- **Custom MIME formats cannot be relied on** for lossless in-app paste
  (Safari none, Firefox preview). Use HTML data attributes instead.

---

## 5. What survives Excel → calcplus, per property

| Property | Survives via HTML? | Survives via TSV? | Notes / confidence |
|---|---|---|---|
| Cell values (number, text, bool, error) | ✅ exact, when `x:num`/`x:str`/`x:bool`/`x:err` present | ◐ displayed text only; type must be re-inferred | Number precision: `x:num` carries the full value while the `<td>` text may be rounded. Dates arrive as serial numbers + a format code — good. [Doc: "num Attribute … precise value"] |
| Number formats (currency, %, date, custom) | ✅ `mso-number-format` gives the Excel format code back | ❌ only the rendered string | Best single reason to prefer HTML. Excel escapes separators (`mso-number-format:"\#\,\#\#0"`). [Obs: captured payloads] |
| Formulas | ◐ `x:fmla` (see §3.3 caveat) | ❌ never — TSV contains the computed value | If `x:fmla` is absent the paste is values-only; do not reconstruct formulas from text. |
| Font family/size/bold/italic/underline/strike/colour | ✅ CSS (`font-family`, `font-size`, `font-weight`, `text-decoration`, `color`) | ❌ | Excel uses `windowtext` for "auto" colour. |
| Fill colour/pattern | ✅ `background` + `mso-pattern` | ❌ | Solid fills need both properties on write. |
| Borders | ✅ CSS borders (+ `mso-border-*` for diagonals) | ❌ | Border edges are per-side; adjacent-cell doubling is a known artefact. |
| Horizontal/vertical alignment, wrap, indent | ✅ `text-align`, `vertical-align`, `white-space`, `mso-*` | ❌ | "General" alignment is not representable as a CSS value; treat as "unset". |
| Text rotation | ✅ `mso-rotate` | ❌ | |
| Merged cells | ✅ `colspan`/`rowspan` | ❌ (merged value appears once, blanks elsewhere) | Grid expansion must account for spans before indexing cells. |
| Hyperlinks | ◐ `<a href>` | ❌ | The URL is in the anchor. Whether Excel/calcplus reconstructs a real hyperlink on paste is not documented; test. |
| Row heights | ◐ `height` on `<tr>`/`<td>` | ❌ | Excel uses `mso-height-source:userset` to say the height is explicit. |
| Column widths | ◐ `<col width>` present in HTML, but **Excel ignores widths on paste** (auto-fits) | ❌ | [Obs: reverse-engineering notes §8/§14] |
| Cell comments/notes | ❌ | ❌ | Not in HTML/TSV. |
| Data validation | ❌ | ❌ | |
| Conditional formatting | ❌ | ❌ | Only the *resulting* style of the current state is baked into CSS. |
| Charts, images, shapes | ❌ as objects | ❌ | They may appear only as a separate picture flavour on the OS clipboard. |
| Defined names, sheet names/structure, multiple sheets | ❌ | ❌ | A clipboard payload is a rectangular range, not a workbook. |
| External/remote references | ❌ | ❌ | Only the last computed value is present. |
| Hidden rows/columns | ❌ | ❌ | Clipboard copies only visible cells; gaps are collapsed. |

**Verdict for Excel → calcplus:** HTML gives a credible "paste from Excel" that
preserves what users see: values, formats, styles, merges. Formulas are the only
material uncertainty, and they are conditional on `x:fmla`. Comments, validation,
conditional formatting and objects are out of scope for the clipboard and must be
handled by `.xlsx` import.

---

## 6. What calcplus must emit for Excel to paste correctly

### 6.1 Payload shape

Write **one `ClipboardItem` containing both representations**:

```
text/html  →  a single HTML document (no CF_HTML header)
text/plain →  the same range as TSV
```

Chromium maps `text/html` to CF_HTML (adding the byte-offset envelope itself) and
macOS maps it to `public.html`; `text/plain` becomes `CF_UNICODETEXT` /
`public.utf8-plain-text`. [Src: `ClipboardWin::WriteHTML`; Doc: CF_HTML spec]
Chromium and Firefox reject more than one `ClipboardItem` per `write()`, so put
both types in the same item. [Doc: BCD notes for `Clipboard.write`]

HTML skeleton (minimal but Excel-parsable):

```html
<html xmlns:o="urn:schemas-microsoft-com:office:office"
      xmlns:x="urn:schemas-microsoft-com:office:excel"
      xmlns="http://www.w3.org/TR/REC-html40">
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <meta name="ProgId" content="Excel.Sheet">
  <meta name="Generator" content="Microsoft Excel 15">
  <style><!--
    table { border-collapse: collapse; }
    td { mso-number-format:General; font-size:11.0pt;
         font-family:Calibri, sans-serif; vertical-align:bottom;
         white-space:nowrap; mso-rotate:0; }
    .xl1 { font-weight:700; background:#FFF2CC; mso-pattern:black none;
           mso-number-format:"\#\,\#\#0\.00"; border:1.0pt solid #000000; }
  --></style>
</head>
<body>
  <table border=0 cellpadding=0 cellspacing=0>
    <tr height=20>
      <td class=xl1 x:num="1234.5" data-calcplus-v="1234.5"
          data-calcplus-nf="0.00" data-calcplus-f="=A1*2">1,234.50</td>
      <td x:str>00123</td>
    </tr>
  </table>
</body>
</html>
```

Rules:

1. **Do not emit a CF_HTML header** (`Version:`/`StartHTML:`…). The browser builds it.
2. Declare the `x:` namespace and `ProgId` — without them Excel takes a reduced
   "generic web HTML" path and drops `mso-*` styling. [Obs: reverse-engineering
   notes §2.2; consistent with Microsoft's HTML reference]
3. **Type every cell explicitly.** `x:num` for numbers (attribute value = full
   precision, element text = displayed value), `x:str` for text that looks
   numeric, `x:bool`, `x:err`. This is what prevents Excel from guessing a date or
   a formula out of a text cell.
4. **Put the Excel number-format code in `mso-number-format`.** calcplus already
   owns OOXML format codes (ticket 04/07), so the value is the same string Excel
   would store in `numFmt@formatCode`; escape `\ , . # $ %` as Excel's HTML writer
   does. Without it Excel treats pasted values as plain text.
5. **Formulas:** emit `x:fmla="=…"` with an English, A1-style formula and the
   computed value as element text. Add `x:num` when the result is numeric. This is
   the only formula channel a browser can offer Excel. **Caveat:** how Excel
   anchors `x:fmla` references when pasting at an arbitrary cell is not documented;
   see §10.
6. **Merged cells:** `colspan`/`rowspan` on the top-left `<td>` and omit covered
   cells. Cells that abut a merged block may need explicit `border-*` overrides
   (the class may only describe interior edges).
7. **Hyperlinks:** wrap cell content in `<a href="…">`. Best-effort; Excel may
   import the anchor or just the text.
8. **Styles:** inline styles win over class styles in Excel's parser; emit a class
   per unique format combination (`.xl1`, `.xl2`, …) to keep the payload small,
   and use inline only for the per-cell exceptions (indent, rotation, border
   overrides).
9. **Character encoding:** UTF-8, declared in `<meta>`. Excel's CF_HTML is defined
   as UTF-8 on Windows; mismatched encoding is the classic "mojibake on paste" bug.
  [Doc: [HTML Clipboard Format](https://learn.microsoft.com/en-us/windows/win32/dataxchg/html-clipboard-format)]
10. **Private attributes for calcplus→calcplus:** `data-calcplus-v/f/nf/s` on each
    `<td>`. Unknown `data-*` attributes are ignored by Excel and survive DOM
    parsing, giving cross-tab lossless paste without custom MIME types.

### 6.2 TSV rules

- Columns separated by `\t`, rows by `\r\n` (matching Excel's CF_UNICODETEXT).
- **Use the displayed text** (the same string Excel would put on the clipboard),
  not the raw value. A date leaves as `2026-09-01` (locale-formatted), a formula as
  its result, an empty result as an empty field. [Obs: [ReoGrid article](https://web.reogrid.net/articles/paste-from-excel-javascript-spreadsheet);
  SheetJS]
- **Quote a field that contains `\t`, `\r`, `\n`, or `"`**, doubling embedded
  quotes — Excel does exactly this (it is valid tab-separated CSV).
  [Obs: [Super User answer on Excel's quoting](https://superuser.com/questions/324271/how-to-copy-multi-line-text-from-excel-without-quotes/693096);
  [Microsoft Q&A](https://learn.microsoft.com/en-us/answers/questions/5052307/why-is-excel-adding-quotes-when-i-copy-a-cell-with);
  [Clipboard Inspector](https://clipboardinspector.com/blog/parse-pasted-spreadsheet-data)]
- **Formula-injection guard:** a text cell whose display text begins with `=`, `+`,
  `-`, or `@` would be re-interpreted as a formula by Excel when pasted as text.
  In HTML, `x:str` makes the intent explicit. In the TSV fallback, prefix such
  text with `'` (Excel's text marker) or have the HTML path take precedence — but
  document the behaviour either way.
- Cap the payload (e.g. 100k cells / a few MB) and warn; the Windows clipboard
  history truncates around 4 MB and browsers are not specified for huge strings.

### 6.3 Fidelity calcplus → Excel

| Property | Via HTML | Confidence |
|---|---|---|
| Values + cell typing | ✅ | high |
| Number formats | ✅ | high (documented `mso-number-format` mechanism) |
| Font family/size/bold/italic/underline/strike/colour | ✅ | high |
| Fill | ✅ | medium (needs `mso-pattern` companion) |
| Borders | ✅ | medium (edge overrides needed) |
| Alignment / wrap / rotation | ✅ | medium |
| Merged cells | ✅ | high (`colspan`/`rowspan` is how Excel itself does it) |
| Hyperlinks | ◐ | low–medium (anchor import not documented) |
| Row heights | ◐ | low (Excel may auto-fit) |
| Column widths | ❌ | Excel auto-fits on paste |
| Formulas | ◐ `x:fmla` | **verify** — reference anchoring undocumented |
| Comments, validation, conditional formats, objects, names, sheets | ❌ | — |

---

## 7. Browser Clipboard API constraints (the ones that bite)

1. **Secure context is mandatory.** `navigator.clipboard` exists only over HTTPS
   (or `localhost`). A statically hosted, backend-less app still needs TLS.
   [Doc: MDN `Clipboard.read`/`write`, ClipboardItem — "Secure context"]
2. **`write()` needs transient activation.** A click, key press, or the `copy`/`cut`
   event qualifies. Async work (fetch/worker) before the write can lose activation;
   build the payload synchronously and pass it to `write()` immediately. Chromium
   also rejects if the document is not focused. [Src: `clipboard_promise.cc`
   `ValidatePreconditions`; Doc: BCD notes]
3. **`read()` may prompt and may be denied.** Chrome requires the `clipboard-read`
   permission; Firefox 127+ shows a paste prompt; Safari requires a gesture. There
   is no silent polling. [Doc: BCD `Clipboard.read` notes / MDN]
4. **The `paste` event never prompts and carries the data.** It fires on the focused
   element and bubbles; the `clipboardData` store is read-only in a paste handler
   (`setData` there has no effect on the clipboard). [Doc: Clipboard API §6.3–6.4]
5. **Safari (and historically Firefox) only deliver paste when an editable element
   has focus.** WebKit bug 75891: paste events attached to `document`/`body` do
   not fire when focus is not in an input/textarea/contenteditable; the workaround
   is an offscreen (not `display:none`) focused `<textarea>` or contenteditable
   grid root. Implement this in the grid shell from day one. [Src: [WebKit bug
   75891](https://bugs.webkit.org/show_bug.cgi?id=75891)]
6. **`read()` sanitises HTML by default in Chromium.** The sanitised path runs
   `CreateStrictlyProcessedMarkupWithContext`, which re-parses, strips
   `<base>/<link>/<meta>` (including `ProgId`), inlines `<style>` through the CSSOM
   (dropping unknown `mso-*` properties), and re-serialises. `read({unsanitized:
   ['text/html']})` returns the raw string, but is Chromium-only. The **paste event
   returns raw HTML** (`SystemClipboard::ReadHTML` straight through `getData`), so
   it is the fidelity-preserving path. [Src: `clipboard_reader.cc`,
   `serialization.cc`, `clipboard_promise.cc`, `data_object_item.cc`]
7. **One `ClipboardItem` per `write()`.** Passing an array of two items rejects with
   `NotAllowedError` in Chrome/Firefox. Put multiple MIME types inside one item.
   [Doc: BCD note on `Clipboard.write`]
8. **Custom formats are Chromium-only** (`web ` prefix, Chrome 104+; Firefox preview;
   Safari none). Never make them load-bearing. [Doc: BCD `ClipboardItem.type_web`]
9. **Linux has multiple selections.** `read()`/`readText()` read `CLIPBOARD`, not
   `PRIMARY`; middle-click paste is outside the API. The spec leaves the choice
   implementation-defined. [Doc: Clipboard API §7.3.2 note; Src: Chromium buffers]
10. **Drag and drop is not a substitute for paste.** HTML5 DnD exposes whatever
    string types the drag source offers (`text/plain`, `text/html`,
    `text/uri-list`, `Files`), but Excel's drag is an OLE drag, and what survives
    into a browser page is inconsistent across OS/browser pairs. Implement:
    `drop` with `Files` (xlsx/csv) as the reliable path, then `text/html`, then
    `text/plain`; do not claim Excel-drag fidelity. [Doc: [WHATWG HTML DnD](https://html.spec.whatwg.org/multipage/dnd.html);
    Inf for Excel's OLE drag]
11. **Clipboard contents are untrusted input.** Parse with `DOMParser` into an inert
    document; never `innerHTML`. Whitelist attributes/tags before mapping to the
    model. Cap rows/columns/cells and string lengths.
12. **`document.execCommand('copy')` is deprecated but remains the fallback** for
    non-secure contexts and older Safari. It must run inside a `copy` event and set
    both types with `setData`. [Doc: [MDN execCommand](https://developer.mozilla.org/en-US/docs/Web/API/Document/execCommand)]

---

## 8. Recommended interop contract

### 8.1 Read (Paste / Ctrl+V)

```
on paste (document capture, focused grid root / hidden textarea):
  dt = event.clipboardData
  if dt.files contains .xlsx/.xlsm/.csv       → import file (highest fidelity; ticket 04)
  else if "text/html" in dt.types             → parseExcelHtml(dt.getData("text/html"))
  else if "text/plain" in dt.types            → parseTsv(dt.getData("text/plain"))
  else if "text/rtf" in dt.types              → optional: strip RTF to text (v2)
  else                                        → no-op
  preventDefault()
```

`parseExcelHtml` algorithm:

1. If the string starts with `Version:` / contains `StartHTML:`, strip the CF_HTML
   header (may happen on Firefox/Safari, which can return the envelope).
2. If `<!--StartFragment-->` / `<!--EndFragment-->` are present, parse only that
   slice; otherwise parse the whole document.
3. `DOMParser.parseFromString(html, "text/html")`.
4. If there is no `<table>`, treat the body's text as a single cell (single-cell
   Excel copies look like this).
5. Build the grid by walking `<tr>`/`<td>`, consuming `colspan`/`rowspan`; reject
   or clamp at the sheet limits.
6. Per cell, in order: `x:err` → error; `x:bool` → boolean; `x:fmla` → formula
   (+ `x:num`/text for the cached value); `x:num` → number (attribute value, else
   element text); `x:str` → text; else element text. Strip `<br>` to `\n`
   (`mso-data-placement:same-cell`).
7. Resolve formatting: inline `style` first, then the `<style>` rule for the
   element's class; map font/weight/style/decoration/colour/background/border/
   alignment/white-space/`mso-number-format`/`mso-rotate`.
8. Hyperlink: first `<a href>` inside the cell.
9. Emit a command (ticket 09) that pastes as one undoable batch.

### 8.2 Write (Copy / Cut)

```
on copy/cut (or a click handler for the ribbon buttons):
  payload = buildPayload(selection)       // synchronous, from the model
  item = new ClipboardItem({
     "text/html":  new Blob([payload.html], {type:"text/html"}),
     "text/plain": new Blob([payload.tsv],  {type:"text/plain"})
  })
  await navigator.clipboard.write([item]) // in gesture
  fallback: execCommand("copy") with a one-shot copy handler that setData()s both
```

- Cut: same payload; clear the source cells as part of the same undoable command
  (or on paste, if we keep a "pending cut" marker — decide in ticket 09).
- Always keep a copy of `payload` + the model snapshot in an in-memory clipboard,
  keyed by a `data-calcplus-copy-id` embedded in the HTML.
- Cross-instance paste first tries the in-memory snapshot (same tab); otherwise the
  `data-calcplus-*` attributes in the HTML; otherwise the Excel-compatible parse.

### 8.3 Paste Special (v1)

| Mode | Source representation | Notes |
|---|---|---|
| All (default) | HTML if present, else TSV | Excel-like default |
| Values only | HTML values / TSV | drop formulas and formats |
| Formulas only | `x:fmla` / `data-calcplus-f` | cells without formulas become blank |
| Formats only | HTML styles / `data-calcplus-nf` | no values |
| Transpose | any | apply after parse |
| Text | TSV only | force raw text, no type inference |

Defer "Paste as Picture", "Paste Link", arithmetic operations (add/subtract…) to v2.

### 8.4 Guarantees to state in the PRD

| Direction | Guaranteed | Best-effort | Never |
|---|---|---|---|
| Excel → calcplus | values, cell types, number formats, fonts/fills/borders/alignment, merges | formulas (`x:fmla`), hyperlinks, row heights, column widths | comments, validation, conditional formats, charts/images/shapes, names, sheets, external links |
| calcplus → Excel | values, cell types, number formats, fonts/fills/borders/alignment, merges | formulas (`x:fmla`), hyperlinks, row heights | column widths, comments, validation, conditional formats, charts/images/shapes, names, sheets |
| TSV fallback either way | displayed text only | type inference | formulas, formats, merges |
| `.xlsx` file either way | full per ticket 04 | — | — |

The UI should say "For formulas and full fidelity, use the .xlsx file"; the
clipboard is a convenience channel, not the fidelity channel.

---

## 9. What is browser-impossible (state plainly)

- **Reading Excel's native clipboard formats** (`Biff8`, `Biff12`, `XML Spreadsheet`,
  OLE `Embed Source`). No web API can enumerate or request arbitrary OS formats;
  only the mandatory/optional MIME set is reachable. [Doc: Clipboard API §6.4–6.5]
- **Writing Excel's native formats**, so Excel→calcplus→Excel cannot round-trip
  comments, validation, conditional formatting, names or sheet structure.
- **Silent clipboard reads.** Every read is gesture-gated and may prompt.
- **Delayed rendering** (Excel's "cut clears on paste" trick). A browser writes a
  snapshot; it cannot know when/where it was pasted.
- **Reading `PRIMARY` (Linux middle-click) via the API.** Out of scope by spec.
- **Guaranteeing the `text/html` flavour from Excel on every platform/build.**
  Single-cell copies and some Excel builds degrade the HTML; always keep the TSV
  fallback. [Obs: Univer #6087]

---

## 10. Open questions and conformance tests to build (ticket 13)

These are the assumptions that must become fixtures before the contract is
advertised:

1. **Does Excel 365 (Windows + macOS) put `x:fmla` in the clipboard HTML for a
   multi-cell formula range?** Capture the raw `text/html` from the paste event on
   both platforms; assert formula appears. If not, demote Excel→calcplus formulas
   to "not supported via clipboard" in the PRD.
2. **How does Excel anchor `x:fmla` references on paste?** Copy `=B2*2` from C3 and
   paste into `F10`; observe whether Excel keeps `B2` or rebases. Then define the
   export rule (verbatim vs rebased) and, if needed, a "paste as values" default.
3. **Does Excel import `<a href>` as a hyperlink** on HTML paste, or only as text?
4. **Does Excel honour `height`/`rowspan` on paste**, or auto-fit? (Column widths
   are documented/observed to be ignored.)
5. **Does Excel's TSV parser honour quoting** for embedded tabs/newlines on paste?
   (Its writer quotes; verify the reader.)
6. **Safari paste target**: confirm on current Safari whether a document-level paste
   fires without a focused editable element; if not, ship the hidden-textarea shim.
7. **Firefox HTML fidelity**: confirm `getData("text/html")` from Excel returns the
   full document including `<style>` (not just the fragment).
8. **Excel for the web** as a source: its HTML uses `data-ccp-*` and a nested
   document; assert the parser handles it.
9. **Single-cell copies** from Excel: assert the no-`<table>` path.
10. **Large range limits** (e.g. 50k cells): measure payload size, write time, and
    paste time per browser.

---

## 11. Sources

**Specifications and platform APIs**
- W3C, *Clipboard API and events* (Editor's Draft) — https://w3c.github.io/clipboard-apis/
- WHATWG, *HTML Standard — Drag and drop* — https://html.spec.whatwg.org/multipage/dnd.html
- MDN, *Clipboard API* — https://developer.mozilla.org/en-US/docs/Web/API/Clipboard_API
- MDN, *Clipboard.read()* — https://developer.mozilla.org/en-US/docs/Web/API/Clipboard/read
- MDN, *Clipboard.write()* — https://developer.mozilla.org/en-US/docs/Web/API/Clipboard/write
- MDN, *ClipboardItem* — https://developer.mozilla.org/en-US/docs/Web/API/ClipboardItem
- MDN, *Element: paste event* — https://developer.mozilla.org/en-US/docs/Web/API/Element/paste_event
- MDN, *Document.execCommand()* — https://developer.mozilla.org/en-US/docs/Web/API/Document/execCommand
- MDN browser-compat-data, `api/Clipboard.json` — https://github.com/mdn/browser-compat-data/blob/main/api/Clipboard.json
- MDN browser-compat-data, `api/ClipboardItem.json` — https://github.com/mdn/browser-compat-data/blob/main/api/ClipboardItem.json
- Apple, *NSPasteboard.PasteboardType.html* — https://developer.apple.com/documentation/appkit/nspasteboard/pasteboardtype/html

**Browser engine source and bug trackers**
- Chromium, `content/browser/renderer_host/clipboard_host_impl.cc` (allowed types; policy gating) — https://chromium.googlesource.com/chromium/src/+/main/content/browser/renderer_host/clipboard_host_impl.cc
- Chromium, `ui/base/clipboard/clipboard_win.cc` (CF_HTML read/write, available types, RTF) — https://chromium.googlesource.com/chromium/src/+/main/ui/base/clipboard/clipboard_win.cc
- Chromium, `ui/base/clipboard/clipboard_constants.h` (`kMimeType*`) — https://chromium.googlesource.com/chromium/src/+/main/ui/base/clipboard/clipboard_constants.h
- Chromium, `third_party/blink/renderer/modules/clipboard/clipboard_promise.cc` (permissions, `unsanitized`) — https://chromium.googlesource.com/chromium/src/+/main/third_party/blink/renderer/modules/clipboard/clipboard_promise.cc
- Chromium, `third_party/blink/renderer/modules/clipboard/clipboard_reader.cc` (HTML sanitisation on read) — https://chromium.googlesource.com/chromium/src/+/main/third_party/blink/renderer/modules/clipboard/clipboard_reader.cc
- Chromium, `third_party/blink/renderer/core/editing/serializers/serialization.cc` (`CreateStrictlyProcessedMarkupWithContext`) — https://chromium.googlesource.com/chromium/src/+/main/third_party/blink/renderer/core/editing/serializers/serialization.cc
- Chromium, `third_party/blink/renderer/core/clipboard/data_object_item.cc` (`getData` raw HTML/RTF) — https://chromium.googlesource.com/chromium/src/+/main/third_party/blink/renderer/core/clipboard/data_object_item.cc
- WebKit bug 75891, *Safari fails to fire paste events* — https://bugs.webkit.org/show_bug.cgi?id=75891
- Mozilla bug 586587 (Firefox writes CF_HTML) — https://bugzilla.mozilla.org/show_bug.cgi?id=586587
- Mozilla bug 1254028 (Firefox hides RTF in `clipboardData.types`) — https://bugzilla.mozilla.org/show_bug.cgi?id=1254028
- Chromium issue 40851502 (`clipboard.write` text/html behaviour by platform) — https://issues.chromium.org/issues/40851502

**Microsoft documentation (Excel and Windows clipboard)**
- *HTML Clipboard Format* (CF_HTML grammar, offsets, UTF-8) — https://learn.microsoft.com/en-us/windows/win32/dataxchg/html-clipboard-format
- *Clipboard Formats* (standard vs registered formats) — https://learn.microsoft.com/en-us/windows/win32/dataxchg/clipboard-formats
- *File formats that are supported in Excel* (clipboard-format list, HTML note) — https://support.microsoft.com/en-us/office/file-formats-that-are-supported-in-excel-a28ae1d3-6d19-4180-9209-5a950d51b719
- *File format reference for Word, Excel, and PowerPoint* (clipboard formats) — https://learn.microsoft.com/en-us/office/compatibility/office-file-format-reference
- KB 274308, *Add HTML code to the clipboard* (CF_HTML "such as Excel, Word") — https://learn.microsoft.com/en-us/troubleshoot/developer/visualstudio/cpp/language-compilers/add-html-code-clipboard
- *Microsoft Office HTML and XML Reference* (archived; the `x:` attribute dialect) — https://learn.microsoft.com/en-us/previous-versions/office/developer/office2000/aa155477(v=office.10)
- *Formulas and Data Types* (`fmla`, `bool`, `err`, `num`, `str` attributes; Office HTML reference mirror) — https://stigmortenmyre.no/mso/html/excel/xlconformulas.htm
- *Cell Formatting* (`x:num` + `mso-number-format`; Office HTML reference mirror) — https://stigmortenmyre.no/mso/html/excel/xlconcellformatting.htm
- *XML Spreadsheet Reference* (`ss:Formula`, `ss:HRef`, `ss:MergeAcross`, `ss:StyleID`) — https://learn.microsoft.com/en-us/previous-versions/office/developer/office-xp/aa140066(v=office.10)
- *XML in Excel and the Spreadsheet Component* — "Copying and Pasting" — https://learn.microsoft.com/en-us/previous-versions/office/developer/office-xp/aa140062(v=office.10)

**Observation / implementation references (weaker evidence, used where Microsoft does not document)**
- Clipboard dump of an Excel copy (format list and sizes) — https://mschaef.com/what_is_in_your_clipboard
- ClipboardEnumViewer format-ID list (`Biff12`, `Hyperlink`, `XML Spreadsheet`) — https://github.com/CarlTBarnes/ClipboardEnumViewer
- Qiita, *Excel clipboard HTML dump* (CF_HTML 1.0 offsets, `ProgId`, `mso-number-format`) — https://qiita.com/kob58im/items/f5d6b8d93467ef97e9c6
- Univer issue #6087 (Excel 2024 / Excel for the web clipboard HTML captures) — https://github.com/dream-num/univer/issues/6087
- *Excel Clipboard HTML specification* (reverse-engineering notes for the write dialect) — https://github.com/weisberg/agent_tools/blob/main/docs/excel_format_on_clipboard.md
- SheetJS, *Clipboard Data* (browser read/write MIME matrix) — https://docs.sheetjs.com/docs/demos/local/clipboard
- Clipboard Inspector, *Parse Pasted Spreadsheet Data* (TSV quoting, merges) — https://clipboardinspector.com/blog/parse-pasted-spreadsheet-data
- Super User, *Excel quotes cells containing LF/CR/TAB* — https://superuser.com/questions/324271/how-to-copy-multi-line-text-from-excel-without-quotes/693096
- Microsoft Q&A, *Why is Excel adding quotes when I copy a cell with line breaks* — https://learn.microsoft.com/en-us/answers/questions/5052307/why-is-excel-adding-quotes-when-i-copy-a-cell-with
- ReoGrid, *Paste from Excel into a JavaScript Spreadsheet* (displayed-text TSV, HTML paste) — https://web.reogrid.net/articles/paste-from-excel-javascript-spreadsheet
- svgrid, *Clipboard / Paste Special* (Excel HTML `x:fmla`/`x:num` on paste) — https://svgrid.com/docs/help/cells/paste-find-structure
- ooxml-excel-editor, `clipboard-html.ts` (parser for Excel clipboard HTML) — https://github.com/ojaDeveloper/ooxml-excel-editor/blob/main/src/core/edit/clipboard-html.ts
