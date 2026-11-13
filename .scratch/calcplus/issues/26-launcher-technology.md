# 26 — Launcher technology and release packaging

Status: open
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
