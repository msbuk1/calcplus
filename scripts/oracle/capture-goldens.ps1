param(
  [Parameter(Mandatory = $true)]
  [string] $WorkbookPath,

  [Parameter(Mandatory = $true)]
  [string] $OutputPath
)

$ErrorActionPreference = 'Stop'

function Get-ExcelVersion {
  param($Excel)
  return $Excel.Version
}

if (-not (Test-Path -LiteralPath $WorkbookPath)) {
  throw "Oracle workbook not found: $WorkbookPath"
}

try {
  $excel = New-Object -ComObject Excel.Application
} catch {
  throw "Excel 365 COM is unavailable on this runner. Provision Windows, Excel 365 and the maintainer's licence before enabling the oracle workflow. $($_.Exception.Message)"
}

try {
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  # msoAutomationSecurityForceDisable = 3. Macro projects are never executed.
  $excel.AutomationSecurity = 3

  $workbook = $excel.Workbooks.Open((Resolve-Path -LiteralPath $WorkbookPath).Path, 0, $true)
  $workbook.Application.CalculateFullRebuild()

  $sheets = @()
  foreach ($worksheet in $workbook.Worksheets) {
    $used = $worksheet.UsedRange
    $cells = @()
    if ($used.Cells.Count -gt 0) {
      foreach ($cell in $used.Cells) {
        if ($cell.Value2 -ne $null -or $cell.Formula -ne $null) {
          $cells += [ordered]@{
            ref     = $cell.Address($false, $false)
            formula = [string] $cell.Formula
            value   = $cell.Value2
            text    = [string] $cell.Text
          }
        }
      }
    }
    $sheets += [ordered]@{ name = $worksheet.Name; cells = $cells }
  }

  $manifest = [ordered]@{
    excelVersion = (Get-ExcelVersion $excel)
    excelBuild   = [string] $excel.Build
    os           = [System.Environment]::OSVersion.VersionString
    locale       = [System.Globalization.CultureInfo]::CurrentCulture.Name
    date1904     = [bool] $workbook.Date1904
    generatedAt  = (Get-Date).ToUniversalTime().ToString('o')
    workbook     = (Split-Path -Leaf $WorkbookPath)
  }

  $golden = [ordered]@{ manifest = $manifest; sheets = $sheets }
  $directory = Split-Path -Parent $OutputPath
  if ($directory -and -not (Test-Path -LiteralPath $directory)) {
    New-Item -ItemType Directory -Path $directory | Out-Null
  }
  $golden | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $OutputPath -Encoding utf8
  Write-Output "Wrote $OutputPath"
} finally {
  if ($null -ne $workbook) { $workbook.Close($false) | Out-Null }
  if ($null -ne $excel) {
    $excel.Quit() | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
  }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
