# M4A visual baseline

The baseline PNG is generated on the maintainer's pinned Chromium runner; it
is tied to that machine, Chromium version and pinned font set
(testing-conformance §6). The test pins the canvas viewport to 1254 by 614, so
a ribbon or status-bar height change never reshapes the baseline. Generate or
deliberately refresh it with:

```powershell
npx playwright test tests/e2e/formatted-visual.spec.ts --update-snapshots
```

Then commit the PNG this test writes here. Rerun the comparison with:

```powershell
npx playwright test tests/e2e/formatted-visual.spec.ts
```

The default project name makes the expected file
`m4a-formatted-sheet-chromium-win32.png` on Windows, or the matching
Chromium/platform suffix on the pinned runner. Do not hand-edit the PNG.
