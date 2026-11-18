# 28 — Hostile-input caps and Tier C detection rules

Status: research complete
Date: 2026-11-18
Ticket: milestone 6, slice 6E (producers and hostile input)
Method: primary specifications (ECMA-376 via the c-rex rendering, MS-OFFCRYPTO, MS-CFB, JAXP Security Guide, libexpat security docs, .NET XmlReader docs, OWASP ASVS), the USENIX WOOT 2019 zip-bomb paper and its talk page, the installed dependency READMEs and docs (fflate 0.8.2, saxes 6.0.0 per package.json), Microsoft Support articles on Excel limits and corrupt-file behaviour, and the existing calcplus contracts. No fixture experiments were run; every claim that needs a fixture check is marked UNVERIFIED.

## TL;DR

The file-compatibility contract already states cap numbers in §8; this note supplies the sourced rationale for each, confirms most of them, and recommends two changes: relax the per-entry compression-ratio tripwire (legitimate repetitive sheets inside the guaranteed envelope can exceed 200:1) and add an explicit wall-clock abort so a hostile file always fails faster than the slowest legitimate open. Detection must run in a fixed order — container magic, ZIP structure, streaming caps, part inventory, Strict namespace check — and modelling must never start before the Strict check passes. The XML parser posture is the strongest result in this note: saxes cannot expand DTD-defined entities at all, so billion-laughs and external-entity attacks are structurally impossible as long as the implementer never populates `parser.ENTITIES` and rejects any `DOCTYPE`.

Recommended caps (full rationale in §1; `*` marks a project decision, not a standard):

| Cap | Recommended default | Unit | Confirm or change vs contract §8 |
|---|---|---|---|
| Compressed file size | 100 MB | bytes on disk | Confirm |
| Total uncompressed size, all entries | 1.5 GB | bytes of inflate output | Confirm |
| ZIP entry count | 20,000 | entries | Confirm |
| Compression ratio, whole package | 200:1 | uncompressed ÷ compressed | Add (contract states it per entry; move it here) |
| Compression ratio, single entry | 1,000:1 | uncompressed ÷ compressed | Change (was 200:1 per entry) |
| Single-entry uncompressed size | 512 MB | bytes of inflate output | Confirm |
| `DOCTYPE` / DTD in any part | 0 allowed | occurrence | Confirm (contract: DTDs rejected, entities not resolved) |
| Accumulated character data per text node or attribute | 1 MB | characters | Add |
| Package open phase (decompress plus first parse), wall clock | 30 s | seconds, then abort | Add |
| Touched cells / formulas | 1M / 200k | cells, per performance contract | Confirm (no change; the refusal line already exists) |

## 1. Concrete cap numbers

### 1.1 Compressed file size: 100 MB (confirm)

A 100 MB ceiling bounds the `ArrayBuffer` held before ZIP parsing and the central-directory scan. Legitimate coverage is comfortable: the performance contract's largest generated variants are 50 MB, so 100 MB is twice the largest file the product promises to open [S15][S16]. No standard mandates a number — OWASP ASVS requires that compressed uploads be checked against a maximum size but sets no value [S1] — so this is a project decision with 2x headroom over the promised envelope. Real-world Excel workbooks above 100 MB compressed exist (UNVERIFIED how common), and they would be rejected; that is consistent with the performance contract, which already refuses anything above 1M touched cells, because such files cannot be opened inside the memory envelope anyway [S15].

### 1.2 Total uncompressed size: 1.5 GB (confirm)

The total caps the sum of inflate output across all entries, counted during streaming (see §4), never trusted from central-directory headers. Rationale is time plus completeness rather than retained memory, because the loader streams and discards (only modelled cells are kept, and those are separately capped by the 1M touched-cell refusal [S15]). Decompression throughput in the browser is hundreds of MB/s, so 1.5 GB represents a small number of seconds — inside the 15 s degraded-envelope open budget [S15]. It is far above any legitimate workbook inside the envelope: 1M touched cells of verbose worksheet XML is on the order of 100–200 MB uncompressed (UNVERIFIED: measure the `degraded-1m` corpus shape once generated). The number is a project decision; no archive standard or OWASP guidance sets it [S1].

### 1.3 ZIP entry count: 20,000 (confirm)

Legitimate `.xlsx` packages contain tens to low hundreds of entries (one part per sheet, drawing, image, pivot part, plus relationships and metadata [S17]). 20,000 is roughly two orders of magnitude above the largest plausible legitimate producer output, so false positives are not expected, while it bounds central-directory parse time and `PartStore` memory. Reference points, not precedents: the ZIP format itself allows 65,535 entries without Zip64 [S13], and Zip64 removes even that ceiling, so the format gives no safety — the cap must come from the application. Project decision; confirm.

### 1.4 Compression ratio: 200:1 whole-package (add), 1,000:1 per entry (change)

This is the one contract number the evidence does not support as written. The contract states 200:1 per entry, but a legitimate repetitive sheet inside the guaranteed envelope can exceed that: 250,000 identical cells produce roughly 15 MB of near-identical row XML, and deflate on such input routinely exceeds 500:1 (UNVERIFIED exact figure: measure a generated repetitive sheet before slice 6E starts). A strict 200:1 per-entry tripwire would therefore reject legitimate workbooks the product promises to open.

The safe restatement uses two tripwires for the two different things a ratio check defends against:

- **Whole-package 200:1 (new).** Keeps the contract's number but applies it to totals (total uncompressed ÷ total compressed). Ordinary `.xlsx` content compresses roughly 3–10x (XML text with markup), so 200:1 on the whole package is about 20x headroom over normal content while still catching any bomb that relies on bulk expansion. The 200 figure itself remains a project decision; published guidance describes ratio checking without setting a default [S13][S14].
- **Per-entry 1,000:1 (relaxed from 200:1).** A single deflate stream cannot legitimately exceed roughly 1,032:1, the maximum compression ratio of the deflate format cited in the zip-bomb literature [S2][S3]. Anything above ~1,000:1 from one entry's single decompression pass is therefore proof of foul play (lied-about header sizes or overlapping data, §3), not of repetitive content — repetitive content is bounded by what one deflate pass can do. The 1,032:1 ceiling is a sourced format property; rounding to 1,000:1 is a project decision.

Both ratios must be computed from **bytes actually produced during streaming**, not from central-directory header fields, because overlapping-entry bombs deliberately lie in the headers (§3) [S2]. A high-ratio but small entry still trips the per-entry wire even when the totals are under 1.5 GB — that is the correct behaviour, since nested bombs such as 42.zip (42 KB expanding to petabytes across nested layers [S3]) are small in absolute terms at every layer except the last.

For the record on precedent strength: Apache Commons Compress documents ratio tracking for bomb detection but sets no default threshold [S13]; a commonly cited community answer uses 100:1 [S14] (mark as community practice, not a standard, and note it is stricter than what the false-positive analysis above allows). Expat's billion-laughs defence tolerates 100x amplification before rejecting [S9]; .NET's `XmlReader` disables DTD processing by default and separately caps entity-derived characters [S10]; JAXP defaults to 64,000 entity expansions and 50M total entity characters [S11]. None of these is directly a ZIP ratio, and all are marked as what they are: library defaults from other ecosystems, useful as order-of-magnitude company, not as authority for our numbers.

### 1.5 Single-entry uncompressed size: 512 MB (confirm)

The largest single legitimate part is the biggest worksheet XML or `sharedStrings.xml`. 512 MB of worksheet XML represents on the order of 8M cells — far above the 1M touched-cell refusal line [S15] — so no legitimate in-envelope file trips it. Its purpose is to bound the largest single inflate stream (and the largest single saxes parse) independently of the 1.5 GB total, so one giant part cannot monopolise the budget. Project decision; confirm. Enforce during streaming per §4, not from the header's declared size.

### 1.6 XML entity limits: DOCTYPE rejected, 1 MB per text node or attribute (add)

The parser posture makes numeric entity-expansion limits almost unnecessary, but both halves are still recommended:

- **Zero `DOCTYPE` allowed (confirm).** OOXML parts never legitimately contain a document type declaration — the schemas define no DTD subset — so any `<!DOCTYPE` in any part is hostile or corrupt by construction. saxes surfaces it as a `doctype` event [S6]; the loader must abort on that event (fail closed). Precedent direction, not precedent numbers: OWASP and every major XML stack converge on disabling DTDs entirely as the single measure that defeats external-entity attacks, billion-laughs, and SSRF-via-XML together [S8][S10].
- **Why numeric entity caps are defence in depth, not the primary defence.** saxes expands exactly the five predefined entities (`&amp; &lt; &gt; &apos; &quot;`) and explicitly does nothing with DTD-defined entities; custom entities only expand if application code adds them to `parser.ENTITIES` [S6]. External entities are never fetched. So with the stock configuration, billion-laughs (nested internal entities expanding kilobytes to gigabytes [S4]) and external-entity attacks cannot function — there is no expansion machinery to abuse. The implementer rule is therefore a negative one: never populate `parser.ENTITIES`, never fetch anything a part references. This immunity should be pinned by a test (a billion-laughs part parses to inert text or is rejected at the `DOCTYPE`, without memory growth).
- **1 MB per text node or attribute (add, project decision).** saxes deliberately dropped sax's `MAX_BUFFER_LENGTH` chunk limiting, and its FAQ says size bounding is the application's job [S6]. The loader must therefore bound accumulation itself: abort if any single text node, CDATA section, or attribute value exceeds 1 MB. Headroom rationale: Excel cells hold at most 32,767 characters [S12], so 1 MB is ~30x the largest legitimate cell text; shared-string items, cell values, and attribute values in legitimate producers stay far below it (UNVERIFIED against LibreOffice and Google Sheets exports — spot-check the producer matrix). For company: JAXP's total-entity default is 50M characters [S11] and Expat starts rejecting at 8 MiB of amplified output [S9]; our 1 MB per-node cap is stricter than both, which is appropriate because OOXML has no legitimate use for megabyte-scale text nodes.

### 1.7 Parse time and memory budget: 30 s wall-clock abort for the open phase (add)

The byte caps will trip first in practice, but a wall clock is needed for pathological-but-under-cap inputs (many small entries, pathological XML nesting depth, adversarial central directories). Recommendation: abort the package-open phase (decompress plus first parse, before modelling) after 30 s of wall time. Rationale ties to the performance contract: the slowest legitimate open (1M touched cells, degraded) is budgeted at 15 s [S15], so 30 s is 2x the worst legitimate case — a hostile file always fails faster than the slowest file the product promises to open. Mark as a project decision needing reference-machine measurement (UNVERIFIED: confirm 15 s holds on the reference machine and that no legitimate producer-matrix file approaches 30 s). Memory needs no new cap: the performance contract already refuses estimated working sets above 1.2 GB and warns at 1 GB [S15]; the loader's estimate at open (compressed size plus modelled-cell projection) feeds that existing gate.

## 2. Detection rules for Tier C rejection, in order

These run in fixed order before any content is modelled — and before the range transformer from research 27 runs, since the Strict detector must precede the transformer [S18]. Each rule either passes or produces exactly one Tier C rejection; the first failure wins and nothing is partially imported [S19].

### Rule 0 — Container magic (first 8 bytes)

- `50 4B 03 04` (`PK\003\004`, local file header) or other `PK` signatures → ZIP path, continue to Rule 1.
- `D0 CF 11 E0 A1 B1 1A E1` → OLE Compound File (CFB) container [S5][S7]. This is the encrypted/legacy branch, not the corrupt branch. Then inspect the CFB directory:
  - Streams `EncryptionInfo` plus `EncryptedPackage` present → **password-protected workbook**. This is the documented OOXML encryption envelope: the whole ZIP package encrypted into `EncryptedPackage`, parameters in `EncryptionInfo` [S5][S7]. Sub-types the implementer can distinguish but must treat identically in v1: Agile encryption (`EncryptionInfo` version 4.4, XML `keyData`/`fileKey` elements [S5][S7]) versus Standard/Extensible (`EncryptionInfo` versions x.2/x.3 with binary header structures [S7], DataSpace streams). IRM (rights-managed) files use the same CFB envelope with an extensible-encryption `EncryptionInfo` (UNVERIFIED exact stream layout — msoffcrypto-tool treats 4.3 as unsupported extensible encryption [S7]); report with the same message, since v1 cannot open either.
  - BIFF `Workbook` stream present, no encryption streams → **legacy `.xls`** → unsupported-format message (contract §5), not the corrupt message.
  - CFB magic with neither → **corrupt** message. (A bare CFB file that is none of the above cannot be a workbook.)
- Anything else (PDF `%PDF`, HTML, plain text, empty file) → **corrupt / not-a-valid-workbook** message. Extension must not be trusted: the contract requires a container check, and renaming `report.pdf` to `report.xlsx` must still reject [S19].

Why magic order matters: an encrypted file is a valid, intentional file the user can fix (supply the password in Excel); a corrupt file is an accident. Confusing the two — reporting "damaged" for a password-protected file — sends the user down the wrong recovery path. The CFB magic plus stream inventory is the discriminator [S5][S7].

### Rule 1 — ZIP structure: end of central directory and central directory parse

Parse the end-of-central-directory record (plus the Zip64 locator/end structures) from the tail, then walk the central directory. Any failure — no end record, truncated directory, inconsistent entry count, entry pointing outside the file — is **corrupt**, rejected with the damaged-file message. This matches the contract's stated signal ("ZIP central-directory parse fails" [S19]) and matches Excel's behaviour at this severity: a file with an unreadable container is not repairable and Excel reports it cannot open the file because the format or extension is not valid [S4]. No repair is attempted in v1: Excel's "found unreadable content … recover?" path applies to damaged XML parts inside an otherwise sound container [S4], and building an Excel-compatible repairer is out of scope — the contract's F19 fixture is "truncated package rejected; no partial import" [S19]. The gap (container sound, XML damaged) is an open scope question in §6.

Additionally, verify each entry's CRC-32 as it is streamed (Rule 3): a CRC mismatch is corruption, same message. CRC catches truncation and bit-rot that a structurally valid central directory misses.

### Rule 2 — Pre-decompression accounting from headers (trip early, but trust nothing)

Before inflating anything, sum the central directory's declared compressed sizes, uncompressed sizes, and entry count, and compare against the §1 caps (100 MB, 1.5 GB, 20,000). Breach → reject immediately without decompressing. This is cheap and catches honest-but-oversized files. It must never be the only check: overlapping-entry bombs forge these very fields (§3), so Rule 3 re-enforces every byte cap against actual inflate output [S2].

### Rule 3 — Streaming caps during decompression (the real enforcement)

Decompress entry by entry through streaming inflate (fflate's `Unzip` class with per-file handlers such as `UnzipInflate`/`UnzipPassThrough`, which decompress each entry without materialising the whole archive [S20]), maintaining running counters: bytes out per entry (1,000:1 and 512 MB wires), bytes out total (1.5 GB wire, 200:1 package wire), entries seen (20,000 wire), elapsed wall time (30 s wire). On any breach: abort the stream, release buffers, reject. Never `await` a whole-archive decompression into memory first — fflate's one-shot `unzip` that returns all files materialises everything and must not be the hostile-path API [S20]. Never recurse: OOXML parts are never themselves ZIPs, so each entry is inflated at most once; anything resembling a nested archive (embedded object bytes) is preserved opaque and never re-entered. This single-pass, no-recursion, count-output-bytes design is exactly the defence the better-zip-bomb paper's findings demand: per-file-ratio-only and nesting-only checks are both evaded by construction (single-layer 28M:1 bombs exist [S2]), while output counting is not.

### Rule 4 — Part inventory: `[Content_Types].xml` and `_rels/.rels` resolve

Parse the two root parts and locate the workbook part via the officeDocument relationship, per the OPC rules in ECMA-376 Part 2 [S17]. Missing content types, missing root relationships, or no officeDocument target → **corrupt** message (this is what a file that is a valid ZIP but not a valid package looks like — e.g. a renamed `.zip` of documents). Extension/format mismatches (`.xls`, `.xlsb`, ODS) that arrive as ZIPs (notably ODS, which is also a ZIP) are caught here by content-type/relationship shape and get the unsupported-format message, not the corrupt one.

### Rule 5 — Strict namespace check (namespace discriminator, not content check)

Confirming the contract's existing decision: Strict rejection is a namespace test. Strict (ISO/IEC 29500 Strict, ECMA-376 Part 1) parts use `http://purl.oclc.org/ooxml/` namespaces; Transitional (Part 4) parts use `http://schemas.openxmlformats.org/` [S21][S22]. The Open XML SDK itself maps the two families to each other by namespace substitution, which confirms namespaces are the discriminator [S22]. Implementer rule: after parsing `xl/workbook.xml`, inspect the root element's namespace URI (and the `conformance` attribute, which openpyxl models as `strict` versus `transitional` [S23]); if any examined part root is in the `purl.oclc.org` family, reject the whole file with the Strict message before modelling anything and before the range transformer runs [S18]. Real-world confirmation that Strict files exist in the wild and break Transitional-only readers: openpyxl issues `UserWarning: File contains an invalid specification` on Strict files rather than reading them [S23], and re-saving through Excel (which defaults to Transitional; Strict requires the explicit "Strict Open XML" save type) converts them [S23]. Check every part root cheaply during inventory if desired, but one positive is enough to reject. UNVERIFIED: obtain an Excel-saved Strict fixture (F21) and confirm the namespace appears on the workbook root as expected rather than only on some parts.

### Rule 6 — XML posture during part parsing (saxes configuration)

- Reject on the `doctype` event (Rule: zero DOCTYPE, §1.6) [S6].
- Construct the parser with `xmlns: true` (namespace-aware, needed for Rule 5 and for part parsing generally) and default entity behaviour; never assign to `parser.ENTITIES` [S6].
- Never resolve external references of any kind: no DTD fetching, no schema fetching, no relationship-target fetching (external-link targets are preserved bytes, never opened [S19]). In the browser there is additionally no local-file surface for an external entity to reach, but the posture must not depend on that — mark any Node-only tooling advice (validators, converters) as out of scope for this posture.
- Bound per-node character accumulation at 1 MB (§1.6); bound nesting depth implicitly via streaming row-at-a-time handling rather than DOM building for large parts (worksheets, shared strings).
- Macros: `vbaProject.bin` is copied as opaque bytes, never executed, never inspected [S19] — no macro-based payload can run because there is no execution step, in any context.

## 3. Attack patterns covered

- **Nested zip bombs** (the classic 42.zip: tens of KB expanding to petabytes across layers of nested archives [S3]). Defeated twice: Rule 3 never recurses, and every layer would independently trip the running total and ratio wires.
- **Overlapping-entry, non-recursive bombs** (Fifield 2019: 10 MB expanding to 281 TB at 28M:1 in a single layer, at the format's limits [S2]). This is the pattern that defeats header-only and nesting-only defences, and it is the reason Rules 2 and 3 both exist with Rule 3 counting actual output bytes [S2].
- **Header lying generally.** Declared uncompressed sizes, entry counts, and ratio-relevant fields are attacker-controlled; only counted output is trusted (§1.4, Rule 3) [S2].
- **Billion-laughs entity expansion** (kilobyte document with nested entities expanding to gigabytes [S4]). Structurally impossible through saxes with stock configuration (§1.6), with DOCTYPE rejection as the tripwire and the 1 MB node cap as backstop [S6].
- **External-entity attacks (XXE)** (entity referencing local files or URLs, causing exfiltration or SSRF [S8]). saxes never resolves external entities or fetches DTDs [S6]; the loader additionally fetches nothing. Browser context further removes the local-file target, but the parser posture stands alone.
- **Truncation and bit-rot.** Caught by Rules 1 (structure) and 3 (CRC mismatch). Reported as corrupt, never partially imported [S19].
- **Password-protected and IRM files.** Detected, never decrypted, never modified; reported as protected so the user knows the file is fine and the remedy is in Excel (Rule 0) [S5][S7].

## 4. Streaming enforcement approach

Concretely, for the implementer (verify API names against the installed fflate 0.8.2 sources; the shapes below are per the fflate docs [S20]):

1. Read the file once into an `ArrayBuffer` (already bounded by the 100 MB Rule 2 check — check `file.size` before reading).
2. Scan the central directory from the end record without inflating anything (Rule 1 + Rule 2).
3. Open a streaming `Unzip` session; for each entry, attach a per-file inflate handler and count output bytes in the chunk callback, aborting the session on any §1 breach (Rule 3). Route Tier B (preserved) parts: small ones may be retained as bytes; large ones (media, caches) must be streamed to their preserved store with the same counters, never held twice.
4. Feed decompressed part bytes to saxes in chunks via `write()` (saxes is a streaming SAX parser by design [S6]); handle `doctype` as fatal; accumulate text per node with the 1 MB bound (§1.6).
5. Parse `[Content_Types].xml`, rels, and workbook first (Rules 4–5); only after the Strict check passes does any worksheet modelling or range transformation begin [S18].
6. On any breach at any stage: destroy the unzip session, drop parsed state, surface the single Tier C message for the winning rule (§5). The file on disk is never modified by any Tier C path [S19].

Memory posture throughout: at no point is more than one entry's in-flight chunk plus the modelled cells retained; the 1.5 GB total is output counted, not output kept. JSZip-style whole-archive materialisation is acceptable for tests but must not be the open path, consistent with prior research [S24].

## 5. Rejection message content (per category)

Slice 6C owns the report plumbing; this section specifies only the message content each Tier C path must carry, so the user can tell the categories apart. Wording below extends the contract's §5 messages [S19] with the one action line each category needs. The contract sentence stays verbatim as the headline; the action line follows it.

| Category | Headline (contract §5, verbatim) | Required action line |
|---|---|---|
| Password-protected or IRM-protected | "This workbook is encrypted or password-protected. calcplus cannot open it." | "Open it in Excel, remove the password or rights protection, and save an unprotected copy." Optionally name the detected flavour ("password to open" vs "rights-managed") when Rule 0 distinguishes them. |
| Unsupported format (`.xls`, `.xlsb`, ODS) | "This file format is not supported. Open it in Excel and save as `.xlsx`." | Name the detected format ("This is a legacy Excel 97–2003 workbook / …binary workbook / OpenDocument spreadsheet"). |
| Corrupt, truncated, or non-ZIP | "This file is damaged or is not a valid workbook." | "It may be an incomplete download or transfer. Re-download or re-save it; Excel's Open and Repair may recover its contents." |
| Zip bomb or oversized | "This workbook is too large or unusually compressed to open safely." | No user action fixes hostility; for legitimately large files: "If this is a genuine workbook, split it into smaller files." The detail payload (for logs and the 6C panel, not necessarily the headline) should record which cap tripped and on which entry. |
| Strict OOXML | "This workbook uses an OOXML variant calcplus does not support yet." | "Open it in Excel and save it as a regular Excel Workbook (`.xlsx`); Excel's default save type is the supported variant." |
| XML bomb / DOCTYPE (no contract row; hostile input) | Use the oversized/hostile headline above. | Same detail-payload rule as zip bombs (which part contained the DOCTYPE). Whether this gets its own headline is a 6C/wording decision (open question 4). |

Two reporting rules: never show two Tier C messages (first failing rule wins, §2 ordering), and never import or modify anything on any Tier C path [S19].

## 6. Decisions and open questions

Resolved by the project on 2026-11-18. The four main calls were taken as recommended:

- **Per-entry ratio is 1,000:1, whole-package stays 200:1** (question 1, option taken: loosen). The single-part limit is relaxed because a legitimate sheet of repeated values can exceed 200:1 on one deflate stream, while anything above about 1,032:1 proves a faked header. A fixture measurement of a 250k-identical-cell sheet still confirms the number.
- **The open-phase abort is 30 seconds with its own "took too long to open" message** (question 2, option taken: own message), so a slow file reads differently from a broken one.
- **Cap detail goes in the compatibility panel, not the headline** (question 4, option taken: panel). The headline stays one clear line; which limit tripped and where appears in the panel slice 6C already builds, so one slice owns it.
- **Damaged files of either kind are refused as corrupt with a message naming what is broken** (question 5, option taken: refuse). v1 has no repair pass; this is a recorded deviation from Excel, which offers to repair damaged parts.

Defaults adopted for the smaller items, overridable later:

- The strict fixture for question 6 is to be captured from an Excel save during the slice 6E fixture work, with the namespace checked on the workbook root so the detector fires on the first part parsed.
- A container that is neither the expected compound-file shape nor a ZIP maps to corrupt, matching rule 0, unless a producer-matrix run shows a legitimate writer emitting such a file (question 7).
- Rights-managed and password-protected files share one rejection message, since the remedy differs and a split needs design input (question 8).
- The entity-expansion message headline is owned by slice 6E alongside the other Tier C rejections.

Still open, to be settled by fixture measurement or the producer matrix:

1. Confirm the 1,000:1 per-entry figure against a generated 250k-identical-cell sheet. If the measurement contradicts the analysis, keep the contract's original 200:1.
2. Confirm the 30 second abort holds on the reference machine, and that no producer-matrix file approaches it.
3. Verify the 1 MB per-node text cap against LibreOffice and Google Sheets exports (question 3). The Excel 32,767-character cell ceiling suggests wide headroom, but producer pretty-printing and whitespace are unverified.

## 7. Sources

| ID | Source |
|---|---|
| S1 | OWASP ASVS 5.0, §5.2.3 — compressed files must be checked against maximum uncompressed size and file count before uncompressing (no numeric values set) — https://github.com/OWASP/ASVS/blob/v5.0.0/5.0/en/0x14-V5-File-Handling.md |
| S2 | David Fifield, "A Better Zip Bomb", USENIX WOOT 2019 — single-layer overlapping-entry bomb: 10 MB → 281 TB (28M:1); deflate single-stream ceiling ~1,032:1; header/nesting-only defences evaded — https://www.usenix.org/conference/woot19/presentation/fifield and https://www.usenix.org/system/files/woot19-paper_fifield_0.pdf |
| S3 | Wikipedia, "Zip bomb" — 42.zip: 42 KB of nested archives expanding to petabytes; nesting overcomes deflate's per-stream ratio limit — https://en.wikipedia.org/wiki/Zip_bomb |
| S4 | Wikipedia, "Billion laughs attack" — nested-entity exponential expansion; defences include per-parser memory caps — https://en.wikipedia.org/wiki/Billion_laughs_attack |
| S5 | Microsoft Learn, [MS-OFFCRYPTO]: `\EncryptionInfo` Stream; Agile Encryption (4.4, `keyData`/`fileKey` XML); ECMA-376 Document Encryption — https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto/87020a34-e73f-4139-99bc-bbdf6cf6fa55 and https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto/4f2e28a1-624a-4bed-b3ec-816f0df484ff |
| S6 | saxes README (installed major 6.0.0 per package.json) — expands only the five predefined entities; DTD entity declarations ignored unless the application populates `parser.ENTITIES`; `doctype` event; no `MAX_BUFFER_LENGTH` (application bounds chunks); streaming `write()` design — https://raw.githubusercontent.com/lddubeau/saxes/master/README.md |
| S7 | Microsoft Learn, [MS-OFFCRYPTO] `EncryptionInfo` example (Standard/CryptoAPI versions, header structures); [MS-CFB] compound-file magic `D0 CF 11 E0 A1 B1 1A E1`; msoffcrypto-tool version matrix (4.4 Agile, x.2 Standard, x.3 Extensible) — https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto/7ffe54c5-a76c-4eec-9690-dcc4640aac83 and https://github.com/nolze/msoffcrypto-tool/blob/master/msoffcrypto/format/ooxml.py |
| S8 | OWASP XML External Entity Prevention Cheat Sheet — disable DTDs entirely defeats XXE plus billion-laughs; `XML_PARSE_NOENT`/`DTDLOAD` must stay off — https://cheatsheetseries.owasp.org/cheatsheets/XML_External_Entity_Prevention_Cheat_Sheet.html |
| S9 | libexpat XML Security docs — billion-laughs protection: reject when output exceeds 100x input beyond 8 MiB (defaults: amplification 100.0, threshold 8,388,608 bytes) — https://libexpat.github.io/doc/xml-security/ |
| S10 | Microsoft Learn, .NET `XmlReaderSettings` — DTD processing disabled by default; `MaxCharactersFromEntities` caps entity-derived characters (compat note: reader expansion limited to 10,000,000 characters) — https://learn.microsoft.com/en-us/dotnet/api/system.xml.xmlreadersettings.dtdprocessing and https://github.com/microsoft/dotnet/blob/main/Documentation/compatibility/xmltextreader-dtd-entity-expansion-is-limited-to-10000000-characters.md |
| S11 | Oracle JAXP Security Guide / JDK constants — `entityExpansionLimit` default 64,000; `totalEntitySizeLimit` default 50,000,000 — https://docs.oracle.com/javase/8/docs/technotes/guides/security/jaxp/jaxp.html and https://github.com/openjdk/jdk/blob/master/src/java.xml/share/classes/jdk/xml/internal/JdkConstants.java |
| S12 | Microsoft Support, Excel specifications and limits — cell content max 32,767 characters — https://support.microsoft.com/en-us/excel/excel-specifications-and-limits |
| S13 | Apache Commons Compress docs — ratio-tracking interface for bomb detection (no default threshold); ZIP entry-count ceiling 65,535 without Zip64 — https://commons.apache.org/proper/commons-compress/examples.html and https://commons.apache.org/proper/commons-compress/zip.html |
| S14 | Community practice (mark: not a standard) — 100:1 uncompressed-to-compressed tripwire suggested in Apache Commons Compress zip-bomb discussion — https://stackoverflow.com/questions/72474144/apache-commons-compress-as-solution-to-zip-bomb |
| S15 | calcplus performance contract (`.scratch/calcplus/spec/performance.md`) — 250k/50k guaranteed, 1M/200k degraded, 15 s degraded open, 1 GB warn / 1.2 GB refuse, 50 MB largest generated variant |
| S16 | calcplus prior file research 04 (`.scratch/calcplus/research/04-xlsx-io.md`) — CFB magic detection, Tier C categories, fflate-streaming/saxes-worker-safe architecture, whole-archive materialisation warning |
| S17 | ECMA-376 Part 2 (Open Packaging Conventions) via c-rex; Microsoft Learn SpreadsheetML structure — ZIP package, content types, relationship-targeted workbook location — https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/structure-of-a-spreadsheetml-document |
| S18 | calcplus research 27 (`.scratch/calcplus/research/27-chart-pivot-range-transforms.md`) — Strict detector must run before the range transformer; unknown `extLst` blocks edits |
| S19 | calcplus file-compatibility contract (`.scratch/calcplus/spec/file-compat.md`) §§5/8/9 — Tier C table and messages, existing cap numbers, no-partial-import rule, fixtures F18–F22 |
| S20 | fflate docs — `Unzip` streaming class with per-file handlers (`UnzipInflate`/`UnzipPassThrough`); unzip filter support; one-shot `unzip` materialises all files — https://github.com/101arrowz/fflate and https://github.com/101arrowz/fflate/blob/master/docs/classes/Unzip.md |
| S21 | StackOverflow OOXML Strict-vs-Transitional discriminator — Strict namespaces contain `purl.oclc.org`, Transitional `schemas.openxmlformats.org` — https://stackoverflow.com/questions/24642803/any-way-to-tell-if-an-arbitrary-docx-file-is-in-the-strict-office-open-xml-format-vs-the-transitional-format |
| S22 | Microsoft Learn, Open XML SDK — Strict parts load by mapping `https://purl.oclc.org/ooxml/` namespaces to `https://schemas.openxmlformats.org/` counterparts — https://learn.microsoft.com/en-us/office/open-xml/about-the-open-xml-sdk |
| S23 | openpyxl workbook model (`conformance` strict/transitional); Strict-file failure reports (`UserWarning: invalid specification`; re-save through Excel fixes) — https://openpyxl.readthedocs.io/en/stable/_modules/openpyxl/packaging/workbook.html and https://stackoverflow.com/questions/62800822/openpyxl-cannot-read-strict-open-xml-spreadsheet-format-userwarning-file-contains-an |
| S24 | calcplus prior research 04 §6 — JSZip materialises the whole archive (tests only); fflate streams in the browser |

### Local experiments

None were run for this note. The UNVERIFIED items in §§1.4, 1.6, 1.7, 2 (Rules 0 and 5), and open questions 1–3 each name the fixture or measurement that would resolve them; running that matrix is recommended as the first task of slice 6E.
