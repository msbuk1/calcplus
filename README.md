# calcplus

A browser-based, single-user spreadsheet with close compatibility with Excel
desktop (Microsoft 365) for core spreadsheet work: grid editing, formatting,
formulas, multi-sheet workbooks and high-fidelity `.xlsx` round-trip.

**Status: v1 released (version 0.1.0).** calcplus covers the full v1 feature
matrix: grid editing, formatting, formulas and defined names, several sheets,
sort and AutoFilter, conditional-formatting rendering, high-fidelity `.xlsx`
round-trip, print and PDF, and the Windows and Linux launchers. The release
gate passes in full. See the [release notes](docs/release-notes.md) for what is
in the release, the downloads and first-run notes.

## Download and run

Download from the [latest release](https://github.com/msbuk1/calcplus/releases/latest):

| Platform | File |
|---|---|
| Windows x64 | [`calcplus.exe`](https://github.com/msbuk1/calcplus/releases/latest/download/calcplus.exe) (portable, unsigned), or `calcplus-<version>-windows-x64.zip` with the licence notices |
| Linux x64 | `calcplus-<version>-linux-x64.tar.gz` (static binary and bundled app) |
| Hosted | `calcplus-<version>-hosted.tar.gz` (static build to serve over HTTPS) |

Every release also carries `SHA256SUMS`, the hosted build's checksum list and a
SPDX software bill of materials.

**Windows.** Double-click `calcplus.exe`. A console window opens, the browser
opens the app, and closing the console stops the server. The build is not
code-signed, so SmartScreen may show "Windows protected your PC" on the first
launch: choose **More info**, check that the app is `calcplus.exe`, then choose
**Run anyway**.

**Linux.** Extract the tarball and run the binary:

```bash
tar -xzf calcplus-<version>-linux-x64.tar.gz
cd calcplus-<version>-linux-x64
./calcplus
```

Run `chmod +x calcplus` first if the file is not executable.

The launcher serves the app from `http://127.0.0.1:47823/`, binds to loopback
only, works offline, and needs no Node, Python or other runtime. Chrome or Edge
can open and save a workbook in place; Firefox opens through a file picker and
saves by download.

**Verify a download.** Each release carries `SHA256SUMS` covering every
artifact. Compare it with the value you compute locally:

```powershell
Get-FileHash .\calcplus.exe -Algorithm SHA256
```

```bash
sha256sum calcplus-<version>-linux-x64.tar.gz
```

**Hosted.** The hosted tarball is the built static app. Serve it over HTTPS from
any static host; the File System Access API needs a secure context. It is the
same build the launcher embeds.

## Run from source

Requires Node 24 LTS and npm.

```bash
npm install
npx playwright install chromium   # once, for the E2E smoke test
npm run dev          # Vite app
npm test             # Vitest unit and integration tests
npm run test:e2e     # Playwright Chromium smoke test
npm run bench:browser # Chromium/CDP memory harness (M1 stand-up)
npm run check        # typecheck, lint, unit tests and licence gate
npm run format:check # Prettier check
npm run build        # production build
```

**No tests run on GitHub.** The test suite and stored goldens are run locally
on the maintainer's PC. The Excel 365 oracle is captured manually as described
in `scripts/oracle/README.md`; no GitHub runner or scheduled Excel job is used.
Stored oracle goldens are not committed yet.

The M4A visual baseline is recorded on the maintainer's pinned Chromium runner.
The test pins its canvas viewport, so chrome height changes do not reshape the
image. Regenerate the PNG deliberately with:

```powershell
npx playwright test tests/e2e/formatted-visual.spec.ts --update-snapshots
```

Start here:

- [Spec package](.scratch/calcplus/spec/README.md) — reading order and all contracts
- [PRD](.scratch/calcplus/spec/prd.md) — goals, scope and non-goals
- [Technical design](.scratch/calcplus/spec/technical-design.md) — architecture and components
- [Milestones](.scratch/calcplus/spec/milestones.md) — the eight-phase build plan
- [Spec map](.scratch/calcplus/map.md) — destination, decisions and open work
- [Glossary](CONTEXT.md) — the project vocabulary
- [Development setup](docs/dev-setup.md) — repo layout, toolchain and local checks
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
