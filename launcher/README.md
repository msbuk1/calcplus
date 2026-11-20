# calcplus launcher

This folder builds the small launcher that runs calcplus without Node or any
other runtime. It serves the bundled static build from the loopback address and
opens the default browser.

- **Local only.** The server binds to the literal `127.0.0.1`, never to a LAN
  address. Nothing on your network can reach it.
- **Offline.** The app files are inside the executable. The launcher makes no
  outbound calls, has no updater and sends no telemetry.
- **One stable address.** `http://127.0.0.1:47823/` is kept across restarts, so
  the recent-files list, file handles and the recovery copy stay available
  (`distribution.md` §2).

## Run it

**Windows x64** — double-click `calcplus.exe`. A console window opens, the
browser opens the app, and closing the console stops the server.

**Linux x64** — extract the tarball and run the binary:

```bash
tar -xzf calcplus-<version>-linux-x64.tar.gz
cd calcplus-<version>-linux-x64
./calcplus
```

If the file is not executable, run `chmod +x calcplus` first.

No Node, Python or other runtime is needed. A current Chromium-based browser
(Chrome or Edge) gives open-and-save-in-place; Firefox opens the file and saves
by download.

## The unsigned-build warning (Windows)

The Windows executable is **not code-signed**. Microsoft SmartScreen warns
about any unsigned program that has not yet built a download reputation, so the
first launch may show a blue screen titled **"Windows protected your PC"**.

What you see and what it means:

1. The warning names the app and says the publisher is unknown. That is
   expected: the build carries no Authenticode signature.
2. Choose **More info**, check that the app is `calcplus.exe`, then choose
   **Run anyway**.
3. Windows remembers the choice for this copy of the file.

Why it is unsigned: code signing (Microsoft Artifact Signing and the identity
checks around it) costs money and is not budgeted for v1; it is recorded in
`distribution.md` §8. Signing stays optional, and an unsigned build is the
documented v1 state, not a defect.

You can verify the download without trusting the warning: the release includes
`SHA256SUMS`. Compare it with the value you compute locally:

```powershell
Get-FileHash .\calcplus.exe -Algorithm SHA256
```

```bash
sha256sum calcplus
```

The source is at <https://github.com/msbuk1/calcplus>; builds are reproducible
from the released tag.

## Options

| Flag | Effect |
|---|---|
| `--port N` | Use another loopback port (1-65535). The default is `47823`. |
| `--no-open` | Do not start the default browser. |
| `--dist DIR` | Serve a build from `DIR` instead of the embedded copy (development and tests). |
| `--version` | Print the launcher version and exit. |

A second launch on the same port reuses the running instance instead of
starting a second server. If another program owns the port, the launcher stops
with a message: close that program, or pass `--port N`.

`--port N` uses a different browser origin. Recent files, remembered file
handles and the recovery copy from port `47823` are not visible there, and the
launcher says so when it starts.

## Updating

Replace the executable (or extract the newer tarball over the older one). There
is no background updater. The app shows its version, and `--version` prints the
launcher's.

## Building

From the repository root:

```bash
npm run build                    # produces dist/
node scripts/build-launcher.mjs  # stages dist/ and builds both artifacts
```

The script writes `release/calcplus-<version>-windows-x64/calcplus.exe` and
`release/calcplus-<version>-linux-x64.tar.gz`, each with a `SHA256SUMS`, plus
one `release/SHA256SUMS` covering the artifacts. It cross-compiles the Windows
binary on Linux (`GOOS=windows GOARCH=amd64`) and needs Go 1.18 or newer. When
Go is missing the script prints a skip and leaves the app build untouched; pass
`--require-go` to make that a failure on a release run. Set `GO_BIN` to pick a
specific Go binary: a gccgo toolchain is refused because it cannot cross-compile
the Windows build or link a static Linux binary.

The launcher uses only the Go standard library. The app's `LICENSE`, `NOTICE`,
`THIRD_PARTY_NOTICES.md` and the Go notice travel with it.

## Licence

calcplus is free software under **GPL-3.0-only**. This package includes
`LICENSE`, `NOTICE`, `THIRD_PARTY_NOTICES.md` and the Go standard library's
licence as `GO-LICENSE`. The launcher is a separate program that embeds the
app; it does not lock the app down and does not prevent you from replacing the
embedded build with a modified one (`licensing.md` §3.4).
