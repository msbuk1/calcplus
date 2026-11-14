# Excel oracle harness (scaffold)

The authoritative conformance oracle is **real Excel 365 on a Windows runner**
(testing-conformance §2). This repository does not provision that runner or
license. Until the maintainer does, the `oracle` workflow is skipped by the
`CALCPLUS_ORACLE_ENABLED` repository variable and pull requests use stored
goldens only.

## External setup still required

1. Provision a self-hosted Windows runner with the label
   `calcplus-oracle-windows` and the maintainer's activated Excel 365.
2. Exercise `capture-goldens.ps1` manually against the fixture corpus.
3. Review the generated manifest (Excel version, build, OS, locale, date
   system and generation date) and commit the golden.
4. Set `CALCPLUS_ORACLE_ENABLED=true` only after the capture is reproducible.
5. Add the LibreOffice headless and openpyxl relationship checks at release
   time.

## Script

`capture-goldens.ps1` opens a workbook read-only through COM, calls
`CalculateFullRebuild`, captures each used cell's formula, value and displayed
text, and writes a JSON golden with a manifest. It fails with a clear message
when Excel COM is unavailable. It never executes macros: `AutomationSecurity`
is forced to `msoAutomationSecurityForceDisable`.
