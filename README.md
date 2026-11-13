# calcplus

A browser-based, single-user clone of Excel desktop (Microsoft 365): a core
credible spreadsheet with grid editing, formatting, formulas, multi-sheet
workbooks and high-fidelity `.xlsx` round-trip.

**Status: specification.** This repository currently contains the build-ready
spec package, not a working application. The map is the entry point:

- [Spec map](.scratch/calcplus/map.md) — destination, decisions and open work
- [Glossary](CONTEXT.md) — the project vocabulary
- [Milestones](.scratch/calcplus/spec/milestones.md) — the eight-phase build plan
- [Testing and conformance](.scratch/calcplus/spec/testing-conformance.md) — how
  "faithful" is proved

## Licence

calcplus is free software under the **GNU General Public License, version 3
only** (`GPL-3.0-only`).

- [LICENSE](LICENSE) — the full GPLv3 text
- [NOTICE](NOTICE) — copyright, warranty disclaimer and source link
- [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md) — dependency audit and
  obligations
- [CONTRIBUTING.md](CONTRIBUTING.md) — DCO sign-off and contribution rules

The GPLv3 choice is forced by the HyperFormula GPLv3 edition, which is the
formula engine. Source code is available at
<https://github.com/msbuk1/calcplus>. Each release publishes a source tag and
archive alongside the binaries.

calcplus is an independent project. It is not affiliated with, endorsed by, or
sponsored by Microsoft Corporation. Excel is a trademark of Microsoft
Corporation. calcplus reads and writes `.xlsx` files but does not include
Microsoft code, logos or other Microsoft assets.

## Privacy

No accounts, no telemetry, no workbook uploads. Workbooks never leave the
machine. The hosted build loads only its own static files; the local launcher
makes no outbound calls.
