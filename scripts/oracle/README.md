# Excel oracle harness

The authoritative conformance oracle is **real Excel 365 on the maintainer's
Windows PC** (testing-conformance §2). The capture is manual and the test suite
runs locally; no test or check runs on GitHub.

## Manual capture

1. Open Windows PowerShell on the PC with an activated desktop Excel 365.
2. Run `capture-goldens.ps1` against the fixture corpus. If script execution is
   blocked, use a process-scoped execution policy for that command only; do not
   change the machine policy.
3. Review the generated manifest (Excel version, build, OS, locale, date system
   and generation date) and commit the golden.
4. Add the LibreOffice headless and openpyxl relationship checks locally where
   those tools are installed; the Excel check is a manual release step.

The capture must be reproducible before a golden is committed. There is no
runner, scheduled GitHub job or repository variable.

## Script

`capture-goldens.ps1` opens a workbook read-only through COM, calls
`CalculateFullRebuild`, captures each used cell's typed value, error and
formula status, and writes a JSON golden with a manifest. It fails with a clear
message (including the workbook path and COM HRESULT) when Excel COM or the
workbook open fails. It never executes macros: `AutomationSecurity` is forced
to `msoAutomationSecurityForceDisable`.

Cell records carry `ref`, `hasFormula`, `kind` and a typed `value`. `kind` is
one of `text`, `number`, `boolean` or `error`; blank cells are currently
omitted by the script rather than emitted as `kind: blank`. Errors are mapped
from the `Value2` CVErr to the exact names in testing-conformance §2.1;
`Range.Text` is never used for identity because it can be width-dependent
(`####`). `formula` is emitted only when Excel reports `HasFormula = true`, so a
text value beginning with `=` stays a text constant. Formula text is Excel's
exact `Range.Formula` output (including the leading `=`).

The JSON is written as UTF-8 without a BOM by a stable serialiser rather than
Windows PowerShell 5.1 `Set-Content -Encoding utf8`, which adds a BOM, or
`ConvertTo-Json`, whose output differs between PowerShell 5.1 and 7. The
manifest keeps Excel version/build, OS, locale, date system, generation time
and workbook name.

The script targets `fixtures/oracle/m1-values-errors.xlsx`, an Excel-openable
fixture exercising the M1 value and error surface. The M1 chart fixture is a
preservation fixture whose minimal chart package Excel 365 refuses to open.

## Number-format differential oracle

`ssf` is a differential oracle for number formats only (number-formats §1 and
testing-conformance §2). It is not a dependency, is not fetched from the
network and is not part of the local gate. The committed built-in corpus lives
in `fixtures/expected/m4a-number-formats.json`; run it with:

```bash
npm test -- tests/m4a-number-formats.test.ts
```

If `ssf` is already installed on the maintainer's machine, compare the same
cases manually before changing an expected string. The committed corpus stays
the local gate because there is no Excel 365 golden yet.
