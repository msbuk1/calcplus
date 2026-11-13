# 26 — Launcher technology and release packaging

Researched 2026-11-13. Prototypes were built and measured on Linux x64 (Pop!_OS 24.04, kernel 7.1.5) with **Go 1.26.3**, **Rust 1.96.0**, **Bun 1.4.4** and Node 24.15.0, serving the same 1.1 MB static payload. No macOS or Windows host was available: macOS/Windows signing and packaging claims below are **documented**, not executed.

**Recommendation: Go.** Build a stdlib-only launcher, embed the static build with `//go:embed`, serve it with `http.FileServerFS` on `127.0.0.1`, and ship one static binary per OS (plus a minimal macOS `.app` wrapper so the notarization ticket can be stapled). Rust is a credible smaller alternative but costs more in cross-compilation and dependency surface. Bun is viable and is the best language fit, but a 90–116 MB executable, a `0.0.0.0` default bind, JIT entitlements for macOS signing, and no universal macOS target are the wrong trade-offs for a distribution shell. Tauri is a desktop wrapper and is rejected for v1.

The release matrix is Windows x64 portable `.exe`, macOS universal `.app` in a notarized/stapled `.dmg`, and Linux x64 static tarball (optional `.deb` and arm64), with `SHA256SUMS` and GitHub Releases. Signing budget is roughly **US$220/year** (Apple Developer Program + Microsoft Artifact Signing), or roughly **US$630/year** if an OV certificate is used instead; macOS CI minutes are extra.

## 1. Method and evidence standard

Claims are labelled:

- **Tested** — built and executed in this environment on 2026-11-13. The harness starts each binary, reads the `READY <port>` line from stdout, issues raw HTTP requests over a socket, samples `/proc/<pid>/status` for RSS, and terminates the process. Five runs per binary; medians reported.
- **Documented** — stated in a primary source (vendor documentation, specification, or source code).
- **Thin** — plausible but not directly documented or tested here; called out explicitly.

The measurements are indicative, not a benchmark. They cover process start to listening, one 1.1 MB asset served repeatedly, and Linux RSS. They do not measure browser startup, first-paint, or Windows/macOS memory.

## 2. What the distribution contract requires

From [distribution.md](../spec/distribution.md):

- One executable per OS (Windows, macOS Intel + Apple Silicon or universal, Linux); no Node/Python/runtime to install.
- A static HTTP server bound to `127.0.0.1` on a free port; never a LAN interface.
- Serves the bundled static build and opens the default browser.
- Fully offline; no outbound calls; no background updater.
- Version shown; updates are manual replacement.
- A distribution shell, not a desktop app: no embedded browser engine, no native menus.

Acceptance criteria 1–3 and 7–9 are the hard constraints: one download, loopback-only, offline, no outbound calls, version visible, no desktop app. The launcher must be judged as a **small, security-sensitive static server**, not as an application framework.

## 3. Options at a glance

| Dimension | Go | Rust | Bun `--compile` | Tauri shell |
|---|---|---|---|---|
| Single executable | Yes, static by default with `CGO_ENABLED=0` | Yes; static with musl | Yes, embeds Bun runtime | No: `.app`/DMG, `.exe`/installer, AppImage/deb |
| Cross-compile from one Linux host | Yes: `GOOS`/`GOARCH`; tested Windows, macOS x64/arm64, Linux arm64 | Not really: target std + linker/SDK; macOS SDK and MSVC needed | Yes for per-arch targets; downloads target Bun | No: per-OS build; Linux needs WebKitGTK |
| macOS universal | Two builds + `lipo` | Two builds + `lipo` | **No documented universal target**; `lipo` path unverified | Per-OS bundle; universal not evaluated |
| Linux binary | Static ELF; no glibc dependency (tested) | Dynamic glibc by default; musl target for static | Dynamic glibc (tested); musl target available | App bundle depends on system WebKitGTK |
| Binary size (same payload) | 6.9 MB | 1.8 MB | 90–116 MB | Minimal app claimed <600 KB, plus OS webview |
| Startup to listening (median) | 2.0 ms | 1.8 ms | 6.5 ms | Not measured |
| RSS after serving (median) | 8.7 MB | 4.1 MB | 24.7 MB | Not measured |
| HTTP server | `net/http` stdlib; path-safe `FileServerFS` | `axum`/`tokio` + `rust-embed`; path handling owned by code | `Bun.serve`; **hostname defaults to `0.0.0.0`** | Custom protocol/webview; no loopback server needed |
| Path traversal | `FileServerFS` cleans the path and stays inside the embedded FS; `ServeFile` rejects `..`; tested 404 | Exact-match lookup in prototype; tested safe | Exact-match whitelist required; URL normalization is subtle | Webview origin; different model |
| Host header check | Manual; tested 403 | Manual; tested 200 without it | Manual; tested 200 without it | N/A |
| macOS signing | Standard Mach-O; no JIT entitlements | Standard Mach-O; no JIT entitlements | Bun docs recommend JIT/unsigned-memory entitlements | Tauri signing workflow |
| Language fit | Second language, tiny isolated shell | Second language, more code/deps | **Best**: TypeScript, same toolchain | Rust + frontend |
| Licence | BSD-3-Clause | MIT/Apache-2.0 | MIT | MIT/Apache-2.0 |
| Verdict | **Recommended** | Strong second | Viable, not recommended | Rejected for v1 |

All four licences are GPLv3-compatible, so the app's GPLv3 distribution is not a differentiator.

## 4. Measured results

### 4.1 Build and runtime

| Binary | Build command (summary) | Size | Start → listening | RSS |
|---|---|---:|---:|---:|
| Go | `CGO_ENABLED=0 go build -trimpath -buildvcs=false -ldflags="-s -w -X main.version=1.2.3"` | 6.9 MB | 1.99 ms | 8.7 MB |
| Rust | `cargo build --release` (`opt-level="z"`, LTO, `panic=abort`, `strip`) | 1.8 MB | 1.79 ms | 4.1 MB |
| Bun | `bun build --compile --minify --asset ./dist server.ts` | 110.2 MB | 6.46 ms | 24.7 MB |

The Go binary is statically linked (`ldd` reports "not a dynamic executable"); the Rust and Bun binaries link `libc`/`libgcc` dynamically in this build. Go's static output is a practical advantage on Linux: one binary runs across distributions without a glibc baseline. Rust can match this with `x86_64-unknown-linux-musl`; Bun offers musl targets but the default Linux build is glibc.

### 4.2 Cross-build sizes

Go, from the Linux host:

| Target | Size |
|---|---:|
| `windows/amd64` | 7.06 MB |
| `darwin/amd64` | 7.01 MB |
| `darwin/arm64` | 6.57 MB |
| `linux/arm64` | 6.49 MB |

Bun, cross-compiled from the Linux host:

| Target | Size |
|---|---:|
| `bun-windows-x64` | 116.35 MB |
| `bun-darwin-arm64` | 90.25 MB |

A Go universal macOS binary would be roughly the sum of the two slices (about 13.6 MB before `lipo` headers). A Bun universal binary, if `lipo` worked, would be roughly 200 MB; Bun does not document that path. Rust was not cross-built here; its Linux binary is 1.8 MB.

### 4.3 Security probes

All three prototypes bound only loopback (`ss -ltnp` showed `127.0.0.1:<port>`), served the app offline in a network namespace with only loopback up, and returned 200 for a normal SPA route. Traversal and Host results:

| Probe | Go | Rust | Bun |
|---|---|---|---|
| `GET /%2e%2e/%2e%2e/etc/passwd` | 404 | 200 (index fallback) | 200 (index fallback) |
| `GET /..%2f..%2fetc/passwd` | 404 | 404 | 404 |
| `Host: evil.example` on `/` | 403 (custom check) | 200 | 200 |
| `GET /assets/` | 200 (index fallback in prototype) | 200 (index fallback) | 200 (index fallback) |

The 200s for the first traversal probe are **not** `/etc/passwd`: the URL parser normalized the path and the SPA fallback served `index.html`. The important lesson is that status codes are misleading; the body and the lookup strategy are what matter. Go's `http.FileServerFS` keeps requests inside the embedded `fs.FS`, and `ServeFile` rejects `r.URL.Path` values containing a `..` path element; Rust and Bun prototypes used an exact-match map, which is safe by construction but must be written deliberately. A production launcher should reject any path containing `..` with 400 before the SPA fallback, and should never join a request path onto a filesystem path.

Go's `http.FileServer` will serve a directory listing if a directory has no `index.html`; the prototype's SPA fallback masked this, but a production wrapper must reject directory requests or ensure every directory has an index. Bun's `Bun.serve` does not validate the `Host` header; Go's prototype did, via a one-line check. These are small code differences with real security consequences.

### 4.4 Offline

Each binary served `200` inside `unshare -rn` after bringing loopback up. This is a useful acceptance test: it proves the launcher itself needs no network. It does not prove the browser will not make requests; the app's CSP and code must ensure that.

## 5. Go

### 5.1 Single executable and embedding

`//go:embed` initializes an `embed.FS` with files from the package directory at compile time; the docs explicitly show `http.FileServer(http.FS(content))` for static content ([Go embed](https://pkg.go.dev/embed)). The launcher is roughly:

```go
//go:embed dist
var content embed.FS

sub, _ := fs.Sub(content, "dist")
ln, _ := net.Listen("tcp4", "127.0.0.1:47823")
http.Serve(ln, http.FileServerFS(sub))
```

`http.FileServerFS` serves an `fs.FS`; `ServeFile` documents that it rejects requests whose `r.URL.Path` contains a `..` path element, and `FileServer` cleans the path before serving ([net/http](https://pkg.go.dev/net/http#FileServer), [ServeFile](https://pkg.go.dev/net/http#ServeFile), [source](https://cs.opensource.google/go/go/+/go1.27.4:src/net/http/fs.go)). This is the main security reason to prefer Go's stdlib over a hand-rolled server: the traversal guard is already there and well tested.

`net.Listen` documents that an empty or unspecified host listens on all interfaces, while `"127.0.0.1:0"` listens on loopback and chooses a free port; `Listener.Addr()` reveals the chosen port ([net.Listen](https://pkg.go.dev/net#Listen)). Use `tcp4` and the literal `127.0.0.1`, not `localhost`, to avoid IPv6/DNS ambiguity.

### 5.2 Cross-compilation and universal binaries

Cross-compilation is a first-class Go feature: set `GOOS` and `GOARCH`; `CGO_ENABLED=0` keeps the result pure Go and static ([Go environment variables](https://go.dev/doc/install/source#environment), [go help environment](https://pkg.go.dev/cmd/go#hdr-Environment_variables)). This environment produced working Windows, macOS x64/arm64 and Linux arm64 binaries from one Linux host with no extra toolchain.

A macOS universal binary is two Go builds plus Apple's `lipo -create` ([Apple: building a universal macOS binary](https://developer.apple.com/documentation/apple-silicon/building-a-universal-macos-binary)). `lipo` is part of the macOS toolchain, so the macOS runner that signs and notarizes can also create the fat binary; no Linux-side fat-binary support is needed.

### 5.3 Size, startup, memory

6.9 MB, 2.0 ms to listening, 8.7 MB RSS on this payload. The binary is larger than Rust's because the Go runtime and `net/http` are linked in, but 6.9 MB is negligible next to a multi-megabyte web app and is 16× smaller than the Bun executable. Startup and memory are comfortably below any user-noticeable threshold.

### 5.4 Maintainability and fit

The launcher is a few hundred lines of stdlib Go with no third-party modules: no `go.sum` supply chain, no framework churn, no runtime package manager. The cost is a second language in a TypeScript repository. That cost is bounded because the launcher does not share app logic; it is a distribution shell. For a small, sporadically maintained, security-sensitive component, stdlib Go is easier to audit than a Node/Bun runtime or a Rust dependency graph.

### 5.5 Caveats

- `http.FileServer` directory listings: wrap the handler to reject directories, or guarantee an `index.html` in every directory.
- Windows resources (icon, version metadata) are not in the Go standard library; add a small resource tool or accept a plain `.exe` with the version in the filename and `--version`. This is cosmetic, not architectural.
- `-H=windowsgui` hides the console but also removes the obvious "close this window to stop" affordance; keep the console for v1, or add an explicit quit path.

## 6. Rust

### 6.1 Single executable and embedding

Rust produces a small native executable. `rust-embed` embeds a directory at compile time ([rust-embed](https://docs.rs/rust-embed/latest/rust_embed/)); `axum` + `tokio` provide the HTTP server. The prototype used an exact-match `Assets::get(path)` lookup and was 1.8 MB with `opt-level="z"`, LTO, `panic=abort`, and `strip`.

### 6.2 Cross-compilation

Rust's cross-compilation is target-by-target: `rustup target add` installs only the standard library, and a linker (and often an SDK) is still required ([rustup cross-compilation](https://rust-lang.github.io/rustup/cross-compilation.html)). Windows MSVC needs the MSVC linker; macOS needs the Apple SDK; Linux musl needs a musl toolchain. In practice this means per-OS CI runners or extra tooling (`cargo-zigbuild`, `cargo-xwin`), which is more moving parts than Go's `GOOS`/`GOARCH`. A macOS universal binary still needs two builds plus `lipo`.

### 6.3 Size, startup, memory

Best of the three: 1.8 MB, 1.8 ms, 4.1 MB RSS. If binary size or memory were the dominant constraint, Rust would win. They are not: the download also contains the web app, and 6.9 MB vs 1.8 MB is not a user-visible difference for a spreadsheet.

### 6.4 Security and maintainability

Rust is memory-safe, and `tower-http::ServeDir` has a path-sanitisation story ([ServeDir](https://docs.rs/tower-http/latest/tower_http/services/struct.ServeDir.html)); but the embedded-FS path requires glue, and the prototype's exact-match handler had to own MIME types, traversal checks, and headers. The dependency graph is much larger than Go's stdlib (axum, tokio, hyper, tower, rust-embed and their transitive crates), which is more supply-chain surface to review and update. The launcher would also be the only Rust code in a TypeScript project unless Rust is already used elsewhere.

### 6.5 Verdict

A technically strong second choice, especially if a future native component (for example a Rust/WASM engine) makes Rust a first-class project language. For a standalone shell, Go's simpler cross-compilation and smaller dependency surface outweigh Rust's size advantage.

## 7. Bun (`bun build --compile`)

### 7.1 Single executable and embedding

`bun build --compile` bundles the program and a copy of the Bun runtime into one executable; `--asset ./dist` embeds a directory tree, and `Bun.embeddedFiles` / `import.meta.dir` expose it at runtime ([Bun executables](https://bun.com/docs/compiler/executables)). The prototype used `--asset ./dist`, built a recursive map of exact paths at startup, and served only those paths. This is the right pattern: Bun's URL handling normalizes `..`, so an exact whitelist is safer than joining request paths.

### 7.2 Cross-compilation and universal binaries

Bun documents targets for `bun-windows-x64`, `bun-windows-arm64`, `bun-linux-x64`, `bun-linux-arm64`, `bun-linux-*-musl`, `bun-darwin-x64` and `bun-darwin-arm64` ([Bun executables](https://bun.com/docs/compiler/executables)). There is **no universal macOS target**. A universal binary would require `lipo` of the two Darwin outputs; Bun does not document or test that, and a compiled Bun executable is a self-extracting runtime with appended payload, so whether `lipo` preserves its internal offsets is unverified (**thin**). Treat "Bun universal macOS" as unsupported until proven otherwise. Cross-compiling also downloads the target Bun runtime (observed: ~50 MB per target) and Windows metadata flags other than `hideConsole` cannot be used when cross-compiling because they depend on Windows APIs ([Bun executables](https://bun.com/docs/compiler/executables)).

### 7.3 Size, startup, memory

110.2 MB Linux, 116.4 MB Windows, 90.2 MB macOS arm64; 6.5 ms to listening; 24.7 MB RSS. Bun's own docs say the compiled binary is "still way too big" and recommend `--trim`/AOT experiments; `bun build --compile --trim` only removes the CLI in that mode, and AOT (`bun compile`) is experimental and cannot cross-compile ([Bun size](https://bun.com/docs/compiler/size), [Bun troubleshooting](https://bun.com/docs/compiler/troubleshooting)). For a download that a user keeps on disk, 90–116 MB is a real cost and a larger attack surface.

### 7.4 Security

The critical footgun: `Bun.serve` defaults `hostname` to `"0.0.0.0"` ([Bun HTTP server](https://bun.com/docs/runtime/http/server)). A launcher that forgets `hostname: "127.0.0.1"` binds every interface and violates acceptance criterion 2. `port: 0` does select a free port and `server.port` exposes it, so the mechanics are fine once the hostname is set. Bun's runtime includes `fetch`, DNS, sockets, the package manager, the test runner and the JIT; even if the launcher does not call them, they are in the binary. That is a much larger surface than a Go stdlib server.

macOS signing is also more delicate. Bun's own guide recommends a `codesign` entitlements file with `com.apple.security.cs.allow-jit`, `allow-unsigned-executable-memory`, `disable-executable-page-protection`, `allow-dyld-environment-variables` and `disable-library-validation` ([Bun codesign guide](https://bun.com/docs/guides/runtime/codesign-macos-executable)). Those entitlements weaken the hardened runtime that notarization expects. They may be acceptable for a JIT runtime, but a Go launcher needs none of them, which is a materially better security posture for a local server that only serves static files.

### 7.5 Maintainability and fit

Bun is the best language fit: TypeScript, one toolchain, easy sharing of build scripts and types. The launcher is also the simplest code of the four. If "one language" is an absolute project constraint and a 100 MB download is acceptable, Bun is defensible. But the launcher does not execute app code; it only serves files and opens a browser. The language-fit benefit is therefore small compared with the size, bind-default, and macOS-entitlement costs.

### 7.6 Verdict

Viable, not recommended. If Bun is chosen, use an exact-match asset whitelist, explicitly bind `127.0.0.1`, keep the JIT entitlements to the minimum Bun actually needs, and ship **two** macOS artifacts (x64 and arm64) rather than claiming a universal binary.

## 8. Tauri (rejected shape)

Tauri builds desktop applications that use the system webview rather than bundling a browser engine: WKWebView on macOS, WebView2 on Windows, WebKitGTK on Linux ([Tauri: What is Tauri?](https://v2.tauri.app/start/)). That is smaller than Electron, but it is still a desktop app:

- Distribution is an `.app` bundle/DMG, a Windows installer, or AppImage/deb/RPM, not one executable per OS ([Tauri distribute](https://v2.tauri.app/distribute/)).
- Linux builds require the WebKitGTK development packages (`libwebkit2gtk-4.1-dev`) and related system libraries, and the resulting app depends on the system WebKitGTK runtime ([Tauri prerequisites](https://v2.tauri.app/start/prerequisites/)); Windows uses WebView2. This is a runtime dependency, contrary to the "no runtime to install" promise.
- A Tauri app normally has native window chrome and can add native menus, tray icons and an updater. The distribution contract explicitly excludes those and says the launcher is not a desktop app.
- Signing/notarization is still required on macOS; a free Apple Developer account cannot notarize, and ad-hoc signing still requires the user to whitelist the app ([Tauri macOS signing](https://v2.tauri.app/distribute/sign/macos/)).

Tauri would become the right answer only if the product needs native OS integration (menus, tray, file associations, drag-and-drop onto the icon), a guaranteed browser engine/API, or app lifecycle control that a browser tab cannot provide. It is not the v1 shape. Electron is not a candidate at all: it embeds a browser engine, which the contract forbids.

## 9. The free-port trap: origin-scoped browser storage

This is the most important launcher decision the ticket did not explicitly ask about, and it changes the "free port" wording.

A web origin is the triple **(scheme, host, port)**; two URLs are same-origin only if all three match ([RFC 6454 §3.2, §5](https://www.rfc-editor.org/rfc/rfc6454)). Browser storage is origin-scoped: IndexedDB, `localStorage` and remembered File System Access handles are partitioned by origin. If the launcher picks a random port on every launch, the origin changes (`http://127.0.0.1:49152` → `http://127.0.0.1:53001`), and the app appears to have lost its autosave, recent files and remembered file handles. The distribution contract promises that "file handles are remembered where permitted" ([distribution.md §4](../spec/distribution.md)), so a random port directly contradicts the contract.

**Recommendation:** bind a **stable preferred port** (for example `127.0.0.1:47823`) and treat it as part of the app's identity.

- On startup, try the preferred port. If it is free, bind it.
- If it is in use, probe a small health endpoint (`GET /.well-known/calcplus` returning `{"name":"calcplus","version":"..."}`). If it is another calcplus instance, open the existing URL and exit; this is single-instance behaviour without a native API.
- If another program owns the port, **fail with an actionable message** ("port 47823 is in use; close the other program or run `calcplus --port N`, which uses separate storage"). Do **not** silently fall back to a random port: a silent fallback looks like data loss.
- Keep `--port N` as an escape hatch, and warn that it creates a separate origin and therefore separate browser storage.

This still satisfies "on a free port": the preferred port must be free for the launch to succeed. It also avoids a second server and a second origin. The trade-off is that a rare port collision fails the launch; a clear message is better than silently orphaning the user's autosave.

If ticket 08 decides that v1 uses only explicit file open/save with no IndexedDB, no autosave and no remembered handles, a random port becomes acceptable. That is not the current contract, and it would be a significant reduction in the promised UX. Ticket 08 should confirm the storage model before the launcher fixes the port strategy.

`http://127.0.0.1:<port>` is a **secure context**: the Secure Contexts algorithm treats `127.0.0.0/8` and `::1` as potentially trustworthy, and explicitly notes that the port has no effect on trustworthiness ([W3C Secure Contexts §3.1](https://www.w3.org/TR/secure-contexts/#is-origin-trustworthy)). So the File System Access API works on loopback HTTP without TLS. Use the literal `127.0.0.1`, not `localhost`: `localhost` can resolve to IPv6 or be redirected by resolver configuration, while the literal address is unambiguous.

## 10. Security surface and checklist

### 10.1 Binding and lifecycle

1. Bind `tcp4` to `127.0.0.1` on the stable preferred port; never `0.0.0.0`, never a LAN interface, and do not rely on `localhost` resolution.
2. Bind before opening the browser; read the actual port from the listener, not from a constant, and construct the URL from it.
3. Implement single-instance detection (health endpoint + preferred port). Do not start a second server.
4. On failure to bind, print the URL and a clear error; exit non-zero. Do not leave a half-started process.
5. Open the default browser with an argument vector, never a shell string: `open <url>` on macOS, `xdg-open <url>` on Linux, `rundll32 url.dll,FileProtocolHandler <url>` on Windows. If it fails, print the URL for manual copy.
6. Keep a visible console on Windows/Linux with "calcplus <version> — serving <url>. Close this window to stop." On macOS, test the `.app` quit affordance; if the bundle has no way to quit, document it or add an explicit web-UI quit path.

### 10.2 Request handling

7. Serve only embedded assets. Prefer Go's `http.FileServerFS` with a wrapper that rejects directory requests; alternatively build an exact-match whitelist at startup. Never join a request path onto a filesystem path.
8. Reject any path containing a `..` segment, encoded or not, with 400 before the SPA fallback. Test `%2e%2e`, `..%2f`, backslashes, double-encoding and null bytes.
9. SPA fallback only for extensionless, known routes; return 404 for missing assets. Do not turn every miss into `index.html` (it masks traversal and makes debugging harder).
10. Validate the `Host` header against `127.0.0.1:<port>`; return 403 otherwise. This is cheap DNS-rebinding hardening. Optionally reject cross-origin `Origin` headers on any non-GET request (v1 has none).
11. Set correct MIME types (`text/javascript`, `text/css`, `application/wasm`, `image/svg+xml`) and `X-Content-Type-Options: nosniff`.
12. Set a strict CSP: `default-src 'self'; script-src 'self' 'wasm-unsafe-eval'; style-src 'self' 'unsafe-inline'; img-src 'self' data: blob:; connect-src 'self'; font-src 'self'; worker-src 'self' blob:; object-src 'none'; base-uri 'none'; frame-ancestors 'none'`. Adjust only if the app genuinely needs more; never add external CDNs.
13. Set `Referrer-Policy: no-referrer`, `Cross-Origin-Resource-Policy: same-origin`, and `X-Frame-Options: DENY` (or rely on `frame-ancestors`).
14. Cache hashed assets immutably (`Cache-Control: public, max-age=31536000, immutable`); serve `index.html` and `version.json` with `no-store` so a replaced launcher is not masked by browser cache.

### 10.3 No outbound calls

15. The launcher must contain no HTTP client, update check, telemetry, crash reporting, licence check or DNS use. In Go, importing `net/http` brings the client code into the dependency graph even if unused; review the code and keep the only socket operation the listener. In Bun, the runtime includes `fetch`/DNS/sockets; that is a reason to prefer a smaller runtime, not a reason it cannot be audited.
16. Test with `unshare -rn` (loopback up) and with a network monitor: the launcher must serve the app with no external route. Add a CI grep/audit for `http.Get`, `fetch`, `net.Dial`, `Bun.connect`, update URLs and telemetry SDKs.
17. The app's CSP `connect-src 'self'` prevents the page from calling out; the launcher cannot enforce the app's JavaScript, so this is a second layer.

### 10.4 Acceptance tests

- `ss -ltnp` shows only `127.0.0.1:<port>`.
- `curl --path-as-is` traversal probes return 400/404 and never a host file.
- `Host: evil.example` returns 403.
- `GET /assets/` does not list a directory.
- The app loads offline in a network namespace.
- Browser storage survives a launcher restart (stable port) and is separate under `--port`.
- A second launch opens the existing instance and does not start a second server.
- Killing the console/window stops the listener.

## 11. Release packaging

### 11.1 macOS

**Universal binary.** Build `darwin/amd64` and `darwin/arm64`, then `lipo -create -output calcplus calcplus-amd64 calcplus-arm64` ([Apple: universal binary](https://developer.apple.com/documentation/apple-silicon/building-a-universal-macos-binary)). A universal binary runs natively on both; Rosetta is only needed for x86_64-only builds. Apple's September 2026 notice says **macOS 26 is the final release supporting Intel Macs and Rosetta; macOS 27 is Apple silicon only** ([Apple news](https://developer.apple.com/news/?id=k1mtkt1k)). Universal is therefore a transitional requirement: ship it for v1 to cover macOS ≤26 users, but plan to drop the x86_64 slice when macOS 26 falls out of support. This does not change the v1 contract; it changes the long-term artifact count.

**Signing and notarization.** You need a paid Apple Developer Program membership (US$99/year) and a **Developer ID Application** certificate; a free account cannot notarize ([Apple Developer Program](https://developer.apple.com/programs/), [Tauri macOS signing](https://v2.tauri.app/distribute/sign/macos/)). The notary service requires code signing, a Developer ID certificate, the hardened runtime, a secure timestamp and no `get-task-allow` entitlement ([Apple: notarizing macOS software](https://developer.apple.com/documentation/security/notarizing-macos-software-before-distribution)). Use `codesign --options runtime --timestamp --sign "Developer ID Application: ..."`, then `xcrun notarytool submit ... --wait`, then `xcrun stapler staple` ([Apple: customizing the notarization workflow](https://developer.apple.com/documentation/security/customizing-the-notarization-workflow)). Notarization completes for most software within 5 minutes, and for 98% within 15 minutes, with a limit of 75 submissions per day ([same](https://developer.apple.com/documentation/security/customizing-the-notarization-workflow)).

**Packaging for offline first launch.** Apple states: "Although tickets are created for standalone binaries, it's not currently possible to staple tickets to them" ([Apple: customizing the notarization workflow](https://developer.apple.com/documentation/security/customizing-the-notarization-workflow)). A bare downloaded Mach-O therefore cannot carry its ticket offline. Ship a **minimal `.app` bundle** containing the universal Go binary and an `Info.plist`: sign and notarize the `.app`, staple the ticket to the bundle, put the stapled bundle in a `.dmg`, then notarize and staple the DMG. The `.app` is a double-click wrapper, not a desktop app: it has no embedded browser engine and no native menus; it starts the loopback server and opens the default browser. This also gives the user a normal Finder icon and a drag-to-Applications install. A `.zip` containing the stapled `.app` is an acceptable fallback, but a DMG is the better default.

**Gatekeeper for unsigned builds.** macOS Catalina and later require notarization for Developer ID software; unsigned/unnotarized apps show "Apple cannot check ... for malicious software" or "the developer cannot be verified" ([Apple support: safely open apps](https://support.apple.com/en-us/102445)). Since macOS Sequoia, users can no longer Control-click to override Gatekeeper; they must go to **System Settings > Privacy & Security > Open Anyway** ([Apple news](https://developer.apple.com/news/?id=saqachfa)). Managed Macs may block this entirely. Unsigned is therefore not a viable default for a consumer download; it is only acceptable as a developer/local build.

**Operational note.** The original Developer ID Sub-CA expires on 1 February 2027; certificates issued by it stop working, and new certificates should come from the G2 Sub-CA ([Apple news](https://developer.apple.com/news/?id=w4atic4c)). Build a certificate-renewal task into the release calendar.

### 11.2 Windows

**Portable `.exe` vs installer.** v1 should ship a **portable `.exe`**: one file, no admin rights, no registry, no uninstaller, consistent with "replace the executable to update". An installer (MSI/NSIS) adds desktop shortcuts and uninstall but also more signing/packaging surface and a per-machine install model that the contract does not need. SmartScreen applies to both; an installer does not avoid it.

**Authenticode.** Sign with `signtool sign /fd SHA256 /tr <RFC3161 URL> /td SHA256 <file>`; `/fd` and `/td` are required in current SDKs, and a timestamp keeps the signature valid after the certificate expires ([SignTool](https://learn.microsoft.com/en-us/windows/win32/seccrypto/signtool)). Sign after building and before checksums.

**SmartScreen.** Unsigned downloads show "Windows protected your PC" and require "More info > Run anyway". A valid OV/EV signature removes "unknown publisher" but **does not guarantee** a clean first run: "Even when signed, a newly created binary could still show a SmartScreen warning until its hash or publisher certificate accumulates sufficient evidence of positive reputation" ([SmartScreen reputation](https://learn.microsoft.com/en-us/windows/apps/package-and-deploy/smartscreen-reputation)). Reputation builds over "several weeks and hundreds of clean installs"; unsigned files start from zero for every version, while signed files can inherit certificate reputation ([same](https://learn.microsoft.com/en-us/windows/apps/package-and-deploy/smartscreen-reputation)). EV no longer bypasses SmartScreen: "EV certificates no longer bypass SmartScreen" ([same](https://learn.microsoft.com/en-us/windows/apps/package-and-deploy/smartscreen-reputation)). Plan user-facing release notes that explain the warning and the expected publisher name.

**Signing options and cost.**

- **Microsoft Artifact Signing (formerly Trusted Signing):** US$9.99/month Basic, 5,000 signatures/month, paid Azure subscription, identity validation, no hardware token, no EV ([Artifact Signing pricing](https://azure.microsoft.com/en-us/pricing/details/artifact-signing/), [Artifact Signing FAQ](https://learn.microsoft.com/en-us/azure/artifact-signing/faq), [SmartScreen reputation](https://learn.microsoft.com/en-us/windows/apps/package-and-deploy/smartscreen-reputation)). It does not issue EV certificates ([FAQ](https://learn.microsoft.com/en-us/azure/artifact-signing/faq)). Eligibility/country and identity validation should be checked before committing; individual validation is documented, but organization validation is stricter.
- **OV certificate:** since 1 June 2023, the CA/Browser Forum requires the subscriber's private key to be generated, stored and used in a suitable hardware crypto module ([CA/B Forum code-signing BR §1.2.2, §6.2.7.4.2](https://cabforum.org/working-groups/code-signing/requirements/)). Budget roughly US$500+/year plus a token/HSM; DigiCert lists code-signing subscriptions starting at US$44/month (~US$528/year) ([DigiCert](https://www.digicert.com/signing/code-signing-certificates)).
- **EV certificate:** more expensive (DigiCert lists US$62/month, ~US$744/year) and **no longer helps SmartScreen** ([same](https://www.digicert.com/signing/code-signing-certificates), [SmartScreen reputation](https://learn.microsoft.com/en-us/windows/apps/package-and-deploy/smartscreen-reputation)). Do not buy EV for SmartScreen.

For a small GPLv3 project, Artifact Signing is the cheapest practical path if eligible; otherwise an OV certificate with a hardware token. Signing every release is important: unsigned files cannot inherit reputation, and a new unsigned hash starts at zero.

**Windows on ARM.** x64 emulation covers most users; a native `windows/arm64` build is optional and cheap with Go, but doubles the Windows artifact/signing matrix. Defer it until there is demand.

### 11.3 Linux

There is no OS-wide Gatekeeper/SmartScreen equivalent for a downloaded user binary. Package repositories sign repository metadata, not each package: `apt-secure` verifies the `Release` file and explicitly "does not review signatures at a package level" ([apt-secure(8)](https://manpages.debian.org/bookworm/apt/apt-secure.8.en.html)). A standalone tarball is therefore trusted via the download site plus checksums, not an OS signature.

**Recommended:** a `.tar.gz` containing the static Go binary, a `calcplus.desktop` file, an icon and a README. The desktop entry follows the freedesktop specification (`[Desktop Entry]`, `Name`, `Exec`, `Icon`, `Type=Application`) ([Desktop Entry Specification](https://specifications.freedesktop.org/desktop-entry-spec/latest/)). The tarball is portable and needs no root; the user can run `./calcplus` directly or install the `.desktop` file into `~/.local/share/applications`.

**Optional:** a `.deb` built with `nfpm` or `dpkg-deb` for users who want a system-wide install and menu entry. It requires root to install and adds packaging maintenance; it is not the primary artifact. If a `.deb` repository is ever hosted, APT repository signing is the relevant trust mechanism, not per-binary signing.

**Not recommended for v1:** AppImage. AppImages require FUSE, and on Ubuntu 24.04 the package is `libfuse2t64`; without FUSE the user must use `--appimage-extract-and-run` ([AppImage FUSE troubleshooting](https://docs.appimage.org/user-guide/troubleshooting/fuse.html)). That is extra failure modes for a single static binary that already runs directly. Add AppImage only if desktop-menu integration without a package manager becomes a requirement.

**Architectures:** x64 is the primary Linux artifact. arm64 is cheap with Go (`GOARCH=arm64`) and useful for Raspberry Pi/ARM laptops; ship it if release automation is already matrixed, otherwise defer.

### 11.4 GitHub Releases, checksums and version stamping

GitHub Releases are tag-based and support up to 1,000 assets per release, each under 2 GiB, with no total size or bandwidth limit ([About releases](https://docs.github.com/en/repositories/releasing-projects-on-github/about-releases)). Upload the per-OS artifacts with `gh release create v1.2.3 dist/* -F release-notes.md` ([gh release create](https://cli.github.com/manual/gh_release_create)).

Publish a `SHA256SUMS` file generated **after** signing:

```sh
sha256sum calcplus-1.2.3-windows-x64.exe \
          calcplus-1.2.3-macos-universal.dmg \
          calcplus-1.2.3-linux-x64.tar.gz > SHA256SUMS
```

Users verify with `sha256sum -c SHA256SUMS` (Linux) or `shasum -a 256 -c SHA256SUMS` (macOS). GitHub also exposes a `digest` field on release assets ([REST release assets](https://docs.github.com/en/rest/releases/assets?apiVersion=2022-11-28)); it is useful as a second check, but a published file works offline and covers every asset. Optionally sign `SHA256SUMS` with minisign/GPG; this is useful only if the public key is distributed through a separate trusted channel.

Version stamping:

- **Go:** `-ldflags "-X main.version=$VERSION -X main.commit=$GIT_SHA"`; `calcplus --version` prints it. Generate `dist/version.json` at build time with `{version, commit, buildDate}` and embed it; the app's About/Privacy panel reads the same file. Use `-trimpath -buildvcs=false` and a pinned Go toolchain for reproducible builds ([go tool link](https://pkg.go.dev/cmd/link), [go build](https://pkg.go.dev/cmd/go#hdr-Compile_packages_and_dependencies)).
- **Bun:** `--define BUILD_VERSION='"1.2.3"'` inlines the version at build time ([Bun executables](https://bun.com/docs/compiler/executables)).
- **Rust:** `env!("CARGO_PKG_VERSION")` or a build script; Cargo sets package version environment variables for the crate ([Cargo environment variables](https://doc.rust-lang.org/cargo/reference/environment-variables.html)).

The tag is the single source of truth: CI derives `VERSION` from `v1.2.3`, stamps the binary and `version.json`, signs, then checksums.

### 11.5 Manual updates

The update story is deliberately boring:

- The user downloads the new artifact from GitHub Releases, verifies `SHA256SUMS`, closes the old launcher, and replaces the file/app. There is no updater, no version check, and no outbound request.
- Windows: replace `calcplus.exe`; macOS: replace `calcplus.app` in Applications; Linux: replace the binary in place (or reinstall the tarball).
- The launcher prints `calcplus <version>` and the app shows the same version in About. If the user skips releases, no migration is needed for files; browser-storage migrations are the app's responsibility.
- With a stable port, browser storage (IndexedDB, remembered handles) survives replacement. With a random port, it does not: another reason to fix the port.
- The release notes should say what changed, the expected signing publisher, and any first-run warning (SmartScreen on Windows; Gatekeeper only if unsigned).

What the launcher should show: version, URL, "close this window to stop", and a one-line privacy statement ("works offline; no telemetry; no update check"). What it should not do: check for updates, fetch release notes, report crashes, resolve a licence server, or contact any host. The hosted path is where "latest" lives; the local launcher is intentionally frozen until replaced.

## 12. Release matrix

| Target | Artifact | Format | Signing | Notes |
|---|---|---|---|---|
| Windows x64 | `calcplus-1.2.3-windows-x64.exe` | Portable PE | Authenticode (Artifact Signing or OV) + RFC 3161 timestamp | No installer/admin; SmartScreen may warn until reputation; x64 emulation covers most ARM |
| Windows arm64 (optional) | `calcplus-1.2.3-windows-arm64.exe` | Portable PE | Same | Defer until demand |
| macOS universal | `calcplus-1.2.3-macos-universal.dmg` containing `calcplus.app` | DMG + minimal `.app` bundle | Developer ID Application + notarize + staple app and DMG | Universal x86_64+arm64; Intel support ends with macOS 26 |
| macOS (future) | `calcplus-1.2.3-macos-arm64.dmg` | DMG + `.app` | Same | Drop x86_64 when macOS 26 is out of support |
| Linux x64 | `calcplus-1.2.3-linux-x64.tar.gz` | Static ELF + `.desktop` + icon + README | None; `SHA256SUMS` | Primary Linux artifact; no root needed |
| Linux arm64 (optional) | `calcplus-1.2.3-linux-arm64.tar.gz` | Static ELF | None; `SHA256SUMS` | Cheap with Go; ship if matrixed |
| Linux x64 `.deb` (optional) | `calcplus_1.2.3_amd64.deb` | Debian package | None per package; repository signing only if a repo is hosted | Root install; convenience, not primary |
| All | `SHA256SUMS` (+ optional `SHA256SUMS.minisig`) | Text | Optional minisign/GPG | Generated after signing; verify offline |
| Source | GitHub tag `v1.2.3` | Git tag + auto source archives | N/A | GPLv3 source availability |

## 13. Signing and notarization costs

| Platform | Requirement | Cost | Operational notes |
|---|---|---|---|
| macOS | Apple Developer Program + Developer ID Application + notarization | **US$99/year** ([Apple](https://developer.apple.com/programs/)) | macOS runner required for `codesign`, `notarytool`, `stapler`, `lipo`; notarization included; 75 submissions/day; typically <15 min ([Apple](https://developer.apple.com/documentation/security/customizing-the-notarization-workflow)) |
| Windows (recommended) | Microsoft Artifact Signing Basic | **US$9.99/month (~US$120/year)** ([Microsoft](https://learn.microsoft.com/en-us/windows/apps/package-and-deploy/smartscreen-reputation)) | Paid Azure subscription; identity validation; 5,000 signatures/month; no hardware token; no EV; SmartScreen reputation still builds ([FAQ](https://learn.microsoft.com/en-us/azure/artifact-signing/faq)) |
| Windows (alternative) | OV code-signing certificate | **~US$528/year** list (DigiCert, from US$44/month) + token/HSM ([DigiCert](https://www.digicert.com/signing/code-signing-certificates)) | Hardware key storage required since June 2023 ([CA/B Forum](https://cabforum.org/working-groups/code-signing/requirements/)); more admin overhead |
| Windows (not useful for SmartScreen) | EV code-signing certificate | **~US$744/year** list (DigiCert, from US$62/month) | EV no longer bypasses SmartScreen ([Microsoft](https://learn.microsoft.com/en-us/windows/apps/package-and-deploy/smartscreen-reputation)) |
| Linux | None | **US$0** | Publish checksums; AppImage/deb do not change the trust model ([apt-secure](https://manpages.debian.org/bookworm/apt/apt-secure.8.en.html)) |
| CI | macOS runner minutes; Windows runner for signing | Depends on GitHub plan | Check current GitHub Actions billing for macOS minute multipliers; budget a few minutes per release plus notarization wait. |

**Baseline budget:** Apple US$99 + Artifact Signing ~US$120 = **~US$220/year**, plus CI. With an OV certificate instead: Apple US$99 + OV ~US$528 = **~US$630/year**, plus a hardware token and CI. Unsigned is US$0 but produces a first-run Gatekeeper/SmartScreen wall on two of three OSes.

## 14. Recommendation and trade-offs

**Use Go.** The launcher is a static file server with a browser hand-off; Go's standard library already has the two security-critical pieces (a path-safe file server and an explicit loopback listener), and `GOOS`/`GOARCH` plus `//go:embed` make the release matrix almost trivial. The binary is 6.9 MB, starts in ~2 ms, uses ~9 MB RSS, is statically linked on Linux, and has no third-party dependency graph. The cost is a second language, but the shell is small and isolated; it does not need to share TypeScript types or app logic.

**Why not Bun.** Bun is the best language fit and the simplest code, and 6.5 ms startup is fine. But 90–116 MB per artifact, a `0.0.0.0` default bind that must be overridden, JIT/unsigned-memory entitlements for macOS signing, no universal macOS target, and a runtime that contains the package manager/test runner/network APIs are all costs that a distribution shell should not pay. Choose Bun only if "one language" is an absolute constraint and those costs are explicitly accepted; then use two macOS artifacts and an exact-match whitelist.

**Why not Rust.** Rust wins on size and memory (1.8 MB, 4.1 MB RSS) and is memory-safe, but cross-compilation needs per-target linkers/SDKs, the dependency graph is much larger than Go's stdlib, and the launcher would be the only Rust code in the project. Choose Rust if a future native/WASM component makes Rust a first-class language, or if binary size becomes a hard constraint.

**What would force a desktop wrapper instead.** If calcplus needs native menus, a tray icon, file associations, drag-and-drop onto an icon, a guaranteed browser engine/API (for example File System Access in Safari/Firefox), app lifecycle/autosave-on-close, native file dialogs/printing, or operation in an environment that blocks loopback HTTP servers, the distribution shell stops being a shell. At that point **Tauri** (OS webview, not Electron) is the candidate, with the understanding that it still needs per-OS packaging/signing and a WebKitGTK dependency on Linux. None of those triggers are in the v1 contract; the current contract explicitly accepts the browser's upload/download fallback and says no desktop app.

## 15. Risks and open questions

- **Stable port vs free port (highest impact).** The contract says "free port"; origin-scoped storage needs a stable port. Ticket 08 must confirm whether v1 uses IndexedDB/autosave/remembered handles. If yes, amend the launcher design to stable-port-first and fail-closed on collision; do not silently randomise the port. This is a contract-level decision, not a Go/Bun choice.
- **macOS `.app` lifecycle.** A minimal `.app` wrapper solves stapling and Finder UX, but a CLI process inside a bundle may have no obvious Quit affordance. Test on macOS 26/27. If a real quit/menu is required, that is a desktop-wrapper trigger.
- **macOS signing cost and eligibility.** US$99/year is unavoidable for notarization; a free account cannot notarize. The Developer ID Sub-CA expires 1 February 2027; plan certificate renewal ([Apple](https://developer.apple.com/news/?id=w4atic4c)).
- **Windows SmartScreen is not solved by signing.** Signed new binaries can still warn; reputation takes weeks and hundreds of installs. Artifact Signing is cheap but has identity-validation/eligibility friction. EV is not worth it for SmartScreen.
- **Unsigned builds are a support burden.** macOS Sequoia+ requires System Settings > Privacy & Security > Open Anyway; managed Macs may block. Windows requires "More info > Run anyway"; enterprise policy may block. Do not ship unsigned as the default.
- **Bun universal is unproven.** Bun documents no universal target. If Bun is chosen, ship two macOS artifacts; do not promise a universal binary without a tested `lipo` + codesign + notarize pipeline.
- **Go directory listing.** `http.FileServer` lists directories without an `index.html`; wrap it. Add a traversal test that checks the body, not just the status code.
- **Windows firewall prompt.** Loopback binding is normally exempt, but this was not tested on Windows. Verify on a clean Windows 11 VM; if a prompt appears, document it. (**thin**)
- **CSP and WASM/workers.** A strict CSP can break a WASM engine or worker; test the real app bundle, not the 1.1 MB fixture. `wasm-unsafe-eval` may be required for WebAssembly in Chromium; verify against the actual engine and browser matrix.
- **Browser storage eviction.** IndexedDB is subject to quota/eviction; stable origin is necessary but not sufficient for durable autosave. Ticket 08 owns this.
- **Reproducibility.** Go with `-trimpath -buildvcs=false` and a pinned toolchain is designed to be reproducible; verify by rebuilding the same tag twice and comparing hashes. Bun embeds a runtime version and downloads target runtimes; reproducibility is weaker.
- **Open:** whether to ship Linux arm64 and Windows arm64 in v1; whether to sign `SHA256SUMS`; whether to publish a `.deb`; whether to include source maps in the served build.

## 16. Sources

**Go**

- `embed` package (compile-time embedding, `http.FileServer(http.FS(content))`): https://pkg.go.dev/embed
- `net/http` `FileServer` / `FileServerFS` / `ServeFile` (`..` rejection): https://pkg.go.dev/net/http#FileServer · https://pkg.go.dev/net/http#ServeFile
- `net/http` source (`path.Clean`, `containsDotDot`, directory listing): https://cs.opensource.google/go/go/+/go1.27.4:src/net/http/fs.go
- `net.Listen` (unspecified host vs loopback, port 0): https://pkg.go.dev/net#Listen
- Go environment variables (`GOOS`, `GOARCH`, `CGO_ENABLED`): https://go.dev/doc/install/source#environment
- `go build` flags (`-trimpath`, `-buildvcs`, `-ldflags`): https://pkg.go.dev/cmd/go#hdr-Compile_packages_and_dependencies
- `go tool link` (`-X`, `-s`, `-w`): https://pkg.go.dev/cmd/link
- Go licence (BSD-3-Clause): https://pkg.go.dev/embed?tab=licenses

**Rust**

- rustup cross-compilation (target std + linker/SDK): https://rust-lang.github.io/rustup/cross-compilation.html
- Rust platform support (tiers, targets): https://doc.rust-lang.org/rustc/platform-support.html
- `rust-embed`: https://docs.rs/rust-embed/latest/rust_embed/
- `axum`: https://docs.rs/axum/latest/axum/
- `tower-http` `ServeDir`: https://docs.rs/tower-http/latest/tower_http/services/struct.ServeDir.html
- Cargo environment variables (package version): https://doc.rust-lang.org/cargo/reference/environment-variables.html
- Rust licence (MIT/Apache-2.0): https://github.com/rust-lang/rust/blob/master/LICENSE-MIT

**Bun**

- Single-file executables, `--target`, `--asset`, `--define`, Windows flags: https://bun.com/docs/compiler/executables
- `Bun.serve` (`hostname` default `0.0.0.0`, `port: 0`, `server.port`): https://bun.com/docs/runtime/http/server
- Smaller executables (`--trim`, "still way too big"): https://bun.com/docs/compiler/size
- Troubleshooting (AOT cannot cross-compile): https://bun.com/docs/compiler/troubleshooting
- macOS codesign guide (JIT entitlements, v1.2.4+): https://bun.com/docs/guides/runtime/codesign-macos-executable
- Bun licence (MIT): https://github.com/oven-sh/bun/blob/main/LICENSE.md

**Tauri**

- What is Tauri? (system webview, not bundled engine): https://v2.tauri.app/start/
- Prerequisites (Linux WebKitGTK, Windows WebView2): https://v2.tauri.app/start/prerequisites/
- Distribution (app bundle/DMG, installers, AppImage/deb/RPM): https://v2.tauri.app/distribute/
- macOS signing/notarization (paid account, ad-hoc limits): https://v2.tauri.app/distribute/sign/macos/
- Tauri licence (MIT/Apache-2.0): https://github.com/tauri-apps/tauri/blob/dev/LICENSE-MIT

**Apple**

- Apple Developer Program (US$99/year): https://developer.apple.com/programs/
- Notarizing macOS software (Developer ID, hardened runtime, notarization requirement): https://developer.apple.com/documentation/security/notarizing-macos-software-before-distribution
- Customizing the notarization workflow (`notarytool`, `stapler`, 75/day, no stapling standalone binaries): https://developer.apple.com/documentation/security/customizing-the-notarization-workflow
- Building a universal macOS binary (`lipo`): https://developer.apple.com/documentation/apple-silicon/building-a-universal-macos-binary
- Safely open apps on your Mac (Gatekeeper, Open Anyway): https://support.apple.com/en-us/102445
- macOS Sequoia Gatekeeper change (no Control-click override): https://developer.apple.com/news/?id=saqachfa
- Rosetta/Intel transition (macOS 26 final Intel release): https://developer.apple.com/news/?id=k1mtkt1k · https://developer.apple.com/news/?id=w5ngl9k2
- Developer ID Sub-CA expiry (1 February 2027): https://developer.apple.com/news/?id=w4atic4c
- Hardened runtime: https://developer.apple.com/documentation/security/hardened-runtime

**Microsoft**

- SignTool (`/fd`, `/tr`, `/td`, Authenticode): https://learn.microsoft.com/en-us/windows/win32/seccrypto/signtool
- SmartScreen overview (reputation checks, signature): https://learn.microsoft.com/en-us/windows/security/operating-system-security/virus-and-threat-protection/microsoft-defender-smartscreen/
- SmartScreen reputation (signed still warns; EV no longer bypasses; weeks/hundreds of installs): https://learn.microsoft.com/en-us/windows/apps/package-and-deploy/smartscreen-reputation
- Artifact Signing overview: https://learn.microsoft.com/en-us/azure/artifact-signing/overview
- Artifact Signing FAQ (no EV, identity validation, paid subscription): https://learn.microsoft.com/en-us/azure/artifact-signing/faq
- Artifact Signing pricing: https://azure.microsoft.com/en-us/pricing/details/artifact-signing/
- CA/Browser Forum code-signing baseline requirements (hardware key storage from 2023-06-01): https://cabforum.org/working-groups/code-signing/requirements/
- DigiCert code-signing list pricing (from US$44/month OV, US$62/month EV): https://www.digicert.com/signing/code-signing-certificates

**Linux**

- AppImage FUSE requirement and extract-and-run fallback: https://docs.appimage.org/user-guide/troubleshooting/fuse.html
- Desktop Entry Specification: https://specifications.freedesktop.org/desktop-entry-spec/latest/
- `apt-secure` (repository Release signatures, not per-package): https://manpages.debian.org/bookworm/apt/apt-secure.8.en.html

**GitHub**

- About releases (1,000 assets, 2 GiB per asset, no total limit): https://docs.github.com/en/repositories/releasing-projects-on-github/about-releases
- Release assets REST API (`digest` field): https://docs.github.com/en/rest/releases/assets?apiVersion=2022-11-28
- `gh release create`: https://cli.github.com/manual/gh_release_create

**Web platform**

- RFC 6454, The Web Origin Concept (origin = scheme/host/port): https://www.rfc-editor.org/rfc/rfc6454
- W3C Secure Contexts (`127.0.0.0/8` potentially trustworthy; port has no effect): https://www.w3.org/TR/secure-contexts/#is-origin-trustworthy

**Test environment**

- Pop!_OS 24.04, Linux 7.1.5, x86_64; Go 1.26.3; Rust 1.96.0; Bun 1.4.4; Node 24.15.0; 2026-11-13. Measurements are single-machine, five-run medians, indicative only.
