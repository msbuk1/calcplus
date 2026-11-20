This directory is the embed target for the calcplus static build.

`scripts/build-launcher.mjs` copies the app's `dist/` here before `go build`,
so the launcher binary carries the app and works offline. Run that script; do
not edit the staged copy by hand.

The placeholder file keeps `go build ./launcher` compiling in a source
checkout that has not built the app yet. A binary built that way refuses to
start and says how to build the launcher.
