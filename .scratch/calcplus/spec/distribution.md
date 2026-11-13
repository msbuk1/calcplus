# calcplus v1 — distribution and local run

Resolved by [25 — Distribution and local-run model](../issues/25-distribution.md) on 2026-11-13.

## 1. Promise

calcplus is a static browser app with no backend. An individual can run it two ways:

- **Hosted:** a static host (GitHub Pages, Netlify, Cloudflare Pages) serves the latest build at a URL. Nothing to install; requires a network.
- **Local:** a bundled launcher starts a loopback-only static server and opens the default browser. No Node, Python or other runtime to install; works offline.

A true desktop app (Tauri/Electron) is not part of v1. It is only justified if OS integration becomes necessary.

## 2. Launcher

- One executable per OS: Windows, macOS (Intel + Apple Silicon, or universal) and Linux.
- Starts a static HTTP server bound to `127.0.0.1` on a free port; never binds to a LAN interface.
- Serves the bundled static build and opens the default browser at `http://127.0.0.1:<port>`.
- No runtime dependency; portable executable unless the launcher research decides an installer is better.
- Version shown in the app and/or launcher. Updating means replacing the launcher with a newer release.
- No background updater and no outbound network calls in v1.
- The launcher is a distribution shell, not a desktop app: it does not embed a browser engine or provide native menus.

## 3. Hosted path

- The static build deploys to any static host; HTTPS is required for the File System Access API.
- The hosted URL always serves the latest build.
- No server-side logic, database, accounts or API.
- SPA routing must work on the chosen host (404 fallback or hash routing).

## 4. Browsers and file access

- **Chrome/Edge (desktop):** File System Access API for open and save in place; file handles are remembered where permitted.
- **Safari/Firefox:** upload through a file input and download through a Blob/`<a download>` link; no in-place save.
- The launcher opens the default browser, so both paths must work.
- `file://` is unsupported: modules, WebAssembly, service workers and the File System Access API require `http://localhost` or HTTPS. Show a clear message if the app is opened from `file://`.
- Unsupported browsers get a clear compatibility message rather than a broken grid.

## 5. Offline and updates

- **Launcher:** fully offline after download; the app files are local.
- **Hosted:** requires a network for the initial load; no service worker or PWA in v1.
- Launcher updates are manual: download and replace the executable. The app shows its version.
- Hosted updates happen on deploy; normal HTTP caching only.
- No silent data loss on update; workbooks are local files (or IndexedDB per ticket 08).

## 6. Privacy and trust

- No accounts, no telemetry, no workbook uploads.
- Workbooks never leave the machine.
- The hosted app loads only its own static files; the launcher makes no outbound calls.
- State this in the UI (About/Privacy) and the README.

## 7. Acceptance criteria

1. A user can download one file for their OS, run it, and the app opens in the default browser without installing Node, Python or another runtime.
2. The launcher binds only to `127.0.0.1`; the server is not reachable from the LAN.
3. The launcher works with no network connection.
4. The hosted URL serves the latest build and requires no install.
5. Chrome/Edge can open and save a workbook in place; Safari/Firefox can upload and download it.
6. Opening from `file://` or an unsupported browser shows a clear message.
7. No workbook data is transmitted anywhere; the launcher makes no outbound calls.
8. The app shows its version; replacing the launcher updates it.
9. A Tauri/Electron desktop app is not shipped in v1; the architecture keeps it possible.

## 8. Open decision

Launcher technology and release packaging — [26 — Launcher technology and release packaging](../issues/26-launcher-technology.md): Go vs Rust vs Bun vs Tauri, code signing and notarization, release artifacts and checksums.
