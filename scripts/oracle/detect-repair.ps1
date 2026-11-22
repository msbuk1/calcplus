param(
  [Parameter(Mandatory = $true)]
  [string] $Directory
)

# Repair detector (release-gate item 6): a clean workbook opens through Excel
# COM; a workbook Excel would repair fails to open, because the repair dialog
# cannot be answered in automation. OPEN-OK therefore means "Excel opens it
# without a repair prompt". Each file runs in its own Excel instance so one bad
# file cannot poison the rest. Run it on the output of `npm run repair-outputs`
# with the maintainer's Excel 365 installed.

$ErrorActionPreference = 'Stop'

function Test-Open {
  param([string] $FilePath)

  $excel = $null
  $workbook = $null
  $logsBefore = @(Get-ChildItem $env:TEMP -Filter 'error*.xml' -ErrorAction SilentlyContinue | ForEach-Object { $_.Name })
  try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3
    $excel.AskToUpdateLinks = $false
    try {
      $workbook = $excel.Workbooks.Open((Resolve-Path $FilePath).Path, 0, $true)
      $name = $workbook.Name
      $workbook.Close($false) | Out-Null
      [void] [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook)
      $workbook = $null
      Write-Output "OPEN-OK   $name"
    } catch {
      # Excel writes the repair log asynchronously after the refused open.
      Start-Sleep -Milliseconds 1500
      $logsAfter = @(Get-ChildItem $env:TEMP -Filter 'error*.xml' -ErrorAction SilentlyContinue | ForEach-Object { $_.Name })
      $newLogs = @($logsAfter | Where-Object { $logsBefore -notcontains $_ })
      $kind = if ($newLogs.Count -gt 0) { 'REPAIR' } else { 'BLOCKED' }
      Write-Output "OPEN-FAIL ($kind) $(Split-Path $FilePath -Leaf): $($_.Exception.Message)"
      foreach ($logName in $newLogs) {
        Write-Output "--- repair log ${logName} ---"
        Write-Output (Get-Content (Join-Path $env:TEMP $logName) -Raw)
      }
    }
  } finally {
    if ($null -ne $workbook) {
      try { $workbook.Close($false) | Out-Null } catch {}
      [void] [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook)
    }
    if ($null -ne $excel) {
      try { $excel.Quit() | Out-Null } catch {}
      [void] [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    }
  }
}

foreach ($file in Get-ChildItem (Resolve-Path $Directory).Path -Recurse -Include '*.xlsx', '*.xlsm' | Sort-Object FullName) {
  Test-Open $file.FullName
}
