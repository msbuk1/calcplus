# 25 — Distribution and local-run model

Status: resolved
Type: grilling
Blocked by: none

## Question

How does an individual get calcplus, run it, and keep it up to date — given that v1 is a browser app with no backend?

Decide:

- **Delivery:** static host (GitHub Pages / Netlify / Cloudflare Pages), a one-command local static server, an installable PWA, or a desktop wrapper (Tauri/Electron). Which is the v1 promise, which are optional, and which are later?
- **Local run:** what the user actually does — open a URL, run a command, click an installed icon. Do we ship a `serve` script or launcher?
- **Secure context:** `file://` is not supported; document the `localhost`/HTTPS requirement and the failure mode.
- **File access:** File System Access API where available, download/upload fallback elsewhere; how this interacts with ticket 08.
- **Offline:** service-worker caching, app updates, version skew, and whether workbooks are local files or IndexedDB.
- **Updates:** how a new version reaches the user; cache invalidation; no silent data loss.
- **Privacy and trust:** no backend, no accounts, no telemetry by default; what the user should be told.
- **Desktop wrapper:** if deferred, what would trigger it, and what the architecture must keep possible.

Output: the v1 distribution contract and acceptance criteria for "an individual can install and run calcplus".

## Answer

The v1 distribution contract is in [distribution.md](../spec/distribution.md).

- **Two paths:** a static hosted URL (latest build, nothing to install) and a bundled launcher (local and offline).
- **Launcher:** one executable per OS (Windows, macOS, Linux), loopback-only static server, opens the default browser; no runtime to install.
- **Offline and updates:** the launcher is offline by nature; the hosted URL is online. Launcher updates are manual and versioned; the app shows its version.
- **Browsers:** Chromium-first (File System Access API), with Safari/Firefox upload/download fallback.
- **Privacy:** no accounts, no telemetry, no workbook uploads; the launcher makes no outbound calls.
- **Desktop app:** not v1; only if OS integration is needed.
- **Open:** launcher technology and release packaging — [26 — Launcher technology and release packaging](26-launcher-technology.md).
