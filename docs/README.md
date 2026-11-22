# calcplus documentation

These pages are the published form of the calcplus v1 documents. The
contracts under [`.scratch/calcplus/spec/`](../.scratch/calcplus/spec/README.md)
are the source of truth: where a page here and a contract disagree, the
contract wins and the page is corrected.

| Page | What it covers |
|---|---|
| [Product requirements](prd.md) | The v1 product, its users, scope, non-goals and acceptance criteria |
| [Technical design](technical-design.md) | The architecture outline, invariants, data flow and the main components |
| [ADR index](adr/README.md) | Every recorded architecture decision, with status |
| [Known deviations](known-deviations.md) | Deliberate differences from Excel and the current limitations, each traced to a contract or ticket |
| [Release gate record](release-gate.md) | The M8 release-gate sweep: every testing-and-conformance §13 item, its status and its evidence |
| [Development setup](dev-setup.md) | Repository layout, toolchain and local checks |
| [Command inventory](command-inventory.md) | The generated map of commands and ribbon placement |
| [M6 kickoff](m6-kickoff.md) | Handover notes for file fidelity |
| [M7 kickoff](m7-kickoff.md) | Handover notes for output and hardening, including the launcher lane |

## Running calcplus

- **Hosted:** any static host serves the build over HTTPS; the File System
  Access API needs a secure context.
- **Local:** the Go launcher in [`launcher/`](../launcher/README.md) serves the
  bundled build from `http://127.0.0.1:47823/`, opens the browser and works
  offline. The Windows build is unsigned; the launcher README explains the
  SmartScreen warning.

## Licence and privacy

calcplus is GPL-3.0-only. See [LICENSE](../LICENSE), [NOTICE](../NOTICE) and
[THIRD_PARTY_NOTICES.md](../THIRD_PARTY_NOTICES.md). No accounts, no telemetry
and no workbook uploads: workbooks stay on the machine.
