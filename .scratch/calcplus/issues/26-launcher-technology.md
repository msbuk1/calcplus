# 26 — Launcher technology and release packaging

Status: resolved
Type: research
Blocked by: none

## Question

How should the calcplus launcher be built and shipped for Windows, macOS and Linux?

The launcher is a small distribution shell: it starts a static HTTP server bound to `127.0.0.1`, serves the bundled web app, and opens the default browser. It is not a desktop app and must not embed a browser engine.

Evaluate:

- **Go, Rust, Bun (`bun build --compile`), and a Tauri shell** (as a rejected or alternative shape) on single-executable output, cross-compilation, binary size, startup time, security surface, maintainability, and fit with the TypeScript codebase.
- **macOS:** universal binary (Intel + Apple Silicon), code signing, notarization, and Gatekeeper behaviour for unsigned builds.
- **Windows:** Authenticode signing, SmartScreen warnings, portable `.exe` vs installer.
- **Linux:** AppImage, `.deb` or tarball, desktop entry, and the absence of a signing requirement.
- **Release packaging:** GitHub Releases per OS, checksums, version stamping, and how the static build is bundled into the launcher.
- **Security:** loopback-only binding, a free port, no directory traversal, no outbound calls, no remote access.
- **Update story:** manual replacement per ticket 25; what the launcher should show, and what it should not check.

Recommend one launcher technology and a release matrix, with the signing/notarization cost called out. Note anything that would force a desktop wrapper instead.

## Answer

**Use a stdlib-only Go launcher.** `//go:embed` the static build, serve it with `http.FileServerFS` bound to `127.0.0.1`, and ship one static binary per OS. Go is the only candidate that combines a path-safe file server, an explicit loopback listener, trivial `GOOS`/`GOARCH` cross-compilation and no third-party dependencies; measured 6.9 MB, ~2 ms startup, ~9 MB RSS. Rust is a credible smaller alternative (1.8 MB) but costs per-target linkers/SDKs and a much larger dependency graph. Bun is the best language fit, but a 90–116 MB executable, a `0.0.0.0` default bind, JIT entitlements for macOS signing and no universal macOS target are the wrong trade-offs; Tauri is a desktop wrapper and is rejected for v1.

- **Release:** Windows x64 portable `.exe` (Authenticode); macOS universal `.app` in a notarized/stapled `.dmg`; Linux x64 static tarball (optional `.deb` and arm64); `SHA256SUMS` and GitHub Releases.
- **Cost:** ~US$220/year (Apple US$99 + Microsoft Artifact Signing ~US$120) or ~US$630/year with an OV certificate, plus macOS CI. Unsigned builds hit Gatekeeper/SmartScreen.
- **Key finding:** a random port changes the browser origin and breaks IndexedDB/remembered file handles. Prefer a stable port (`127.0.0.1:47823`), reuse an existing instance, and fail with guidance on collision; ticket 08 must confirm the storage model.
- **Desktop wrapper only if** native menus/file associations, a guaranteed browser engine, lifecycle control or a blocked-loopback environment is required.

Full research: [26-launcher-technology.md](../research/26-launcher-technology.md)
