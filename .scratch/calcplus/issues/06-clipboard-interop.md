# 06 — Clipboard and Excel interop

Status: claimed
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
