param(
  [Parameter(Mandatory = $true)]
  [string] $OutputPath
)

# D-9 capture (ticket 32, open question 3): what does Excel 365 write for a
# broken internal hyperlink? calcplus writes `location="#REF!"` and rewrites
# locations to track their targets (the maintainer's settled convention); this
# script makes one workbook per case, lets Excel itself save the file, and
# reports the exact hyperlink XML it wrote.
#
# Cases:
#   valid           - an internal link whose target survives (the control)
#   deleted-sheet   - the target worksheet is deleted after the link exists
#   deleted-name    - the link points at a defined name that is then deleted
#   never-existed   - the link points at a location that was never valid
#
# Manual capture rules are those of scripts/oracle/README.md: run on the
# maintainer's PC with desktop Excel 365, use a process-scoped execution
# policy, review the output before committing it.

$ErrorActionPreference = 'Stop'

$excel = $null
$workbook = $null
$captureSucceeded = $false
$cleanupErrors = New-Object System.Collections.Generic.List[string]

function Release-ComObject {
  param([object] $ComObject)

  if ($null -eq $ComObject) {
    return
  }
  try {
    [void] [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject)
  } catch {
    # Already released, or not an RCW. Nothing else to do.
  }
}

function Read-ZipEntryText {
  param(
    [Parameter(Mandatory = $true)] [string] $ZipPath,
    [Parameter(Mandatory = $true)] [string] $EntryName
  )

  $zip = [System.IO.Compression.ZipFile]::OpenRead($ZipPath)
  try {
    $entry = $zip.Entries | Where-Object { $_.FullName -eq $EntryName } | Select-Object -First 1
    if ($null -eq $entry) {
      return $null
    }
    $stream = $entry.Open()
    try {
      $reader = New-Object System.IO.StreamReader($stream)
      try {
        return $reader.ReadToEnd()
      } finally {
        $reader.Dispose()
      }
    } finally {
      $stream.Dispose()
    }
  } finally {
    $zip.Dispose()
  }
}

function Get-HyperlinkBlock {
  param([Parameter(Mandatory = $true)] [string] $WorksheetXml)

  $match = [regex]::Match($WorksheetXml, '<hyperlinks>.*?</hyperlinks>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
  if ($match.Success) {
    return $match.Value
  }
  # Excel writes a self-closing <hyperlinks/> only when the list is empty;
  # when no link survived, the element may be absent entirely.
  $empty = [regex]::Match($WorksheetXml, '<hyperlinks\s*/>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
  if ($empty.Success) {
    return $empty.Value
  }
  return '(no <hyperlinks> element)'
}

function Save-CaseWorkbook {
  param(
    [Parameter(Mandatory = $true)] [string] $CaseName,
    [Parameter(Mandatory = $true)] [scriptblock] $BuildAndBreak
  )

  $caseWorkbook = $null
  try {
    $caseWorkbook = $excel.Workbooks.Add()
    & $BuildAndBreak $caseWorkbook
    $savePath = Join-Path $OutputDirectory "$CaseName.xlsx"
    if (Test-Path $savePath) {
      Remove-Item $savePath -Force
    }
    $caseWorkbook.SaveAs($savePath, 51) # xlOpenXMLWorkbook
    return $savePath
  } finally {
    if ($null -ne $caseWorkbook) {
      try {
        $caseWorkbook.Close($false) | Out-Null
      } catch {
        $cleanupErrors.Add("Close($CaseName) failed: $($_.Exception.Message)")
      }
      Release-ComObject $caseWorkbook
    }
  }
}

try {
  Add-Type -AssemblyName System.IO.Compression.FileSystem

  if ([System.IO.Path]::IsPathRooted($OutputPath)) {
    $OutputDirectory = $OutputPath
  } else {
    $OutputDirectory = Join-Path (Get-Location).Path $OutputPath
  }
  [void] [System.IO.Directory]::CreateDirectory($OutputDirectory)

  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $excel.AutomationSecurity = 3 # msoAutomationSecurityForceDisable

  $results = New-Object System.Collections.Generic.List[object]

  # --- valid: the control, a link whose target survives ---------------------
  $validPath = Save-CaseWorkbook -CaseName 'case-valid' -BuildAndBreak {
    param($wb)
    $target = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
    $target.Name = 'Target'
    $data = $wb.Worksheets.Item(1)
    $data.Name = 'Data'
    $data.Range('A1').Value2 = 'Go to target'
    [void] $data.Hyperlinks.Add($data.Range('A1'), '', 'Target!A1', 'the target', 'Go to target')
    $data.Activate()
  }

  # --- deleted-sheet: the target worksheet is deleted ----------------------
  $deletedSheetPath = Save-CaseWorkbook -CaseName 'case-deleted-sheet' -BuildAndBreak {
    param($wb)
    $target = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
    $target.Name = 'Target'
    $data = $wb.Worksheets.Item(1)
    $data.Name = 'Data'
    $data.Range('A1').Value2 = 'Go to target'
    [void] $data.Hyperlinks.Add($data.Range('A1'), '', 'Target!A1', 'the target', 'Go to target')
    $data.Activate()
    $target.Delete() | Out-Null
  }

  # --- deleted-name: the link points at a defined name that is deleted -----
  $deletedNamePath = Save-CaseWorkbook -CaseName 'case-deleted-name' -BuildAndBreak {
    param($wb)
    $data = $wb.Worksheets.Item(1)
    $data.Name = 'Data'
    $data.Range('B2').Value2 = 42
    [void] $wb.Names.Add('GoneName', '=Data!$B$2')
    $data.Range('A1').Value2 = 'Go to name'
    [void] $data.Hyperlinks.Add($data.Range('A1'), '', 'GoneName', 'the name', 'Go to name')
    $data.Activate()
    $wb.Names.Item('GoneName').Delete() | Out-Null
  }

  # --- never-existed: the location was never valid -------------------------
  $neverExistedPath = Save-CaseWorkbook -CaseName 'case-never-existed' -BuildAndBreak {
    param($wb)
    $data = $wb.Worksheets.Item(1)
    $data.Name = 'Data'
    $data.Range('A1').Value2 = 'Go nowhere'
    [void] $data.Hyperlinks.Add($data.Range('A1'), '', 'NoSuchSheet!A1', 'nowhere', 'Go nowhere')
    $data.Activate()
  }

  # --- deleted-row: a structural edit removes the target row ---------------
  $deletedRowPath = Save-CaseWorkbook -CaseName 'case-deleted-row' -BuildAndBreak {
    param($wb)
    $target = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
    $target.Name = 'Target'
    $data = $wb.Worksheets.Item(1)
    $data.Name = 'Data'
    $target.Range('A5').Value2 = 42
    $data.Range('A1').Value2 = 'Go to target'
    [void] $data.Hyperlinks.Add($data.Range('A1'), '', 'Target!A5', 'the target', 'Go to target')
    $data.Activate()
    $target.Range('5:5').Delete() | Out-Null
  }

  # --- inserted-rows: a structural edit above the target moves it ---------
  $insertedRowsPath = Save-CaseWorkbook -CaseName 'case-inserted-rows' -BuildAndBreak {
    param($wb)
    $target = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
    $target.Name = 'Target'
    $data = $wb.Worksheets.Item(1)
    $data.Name = 'Data'
    $target.Range('A5').Value2 = 42
    $data.Range('A1').Value2 = 'Go to target'
    [void] $data.Hyperlinks.Add($data.Range('A1'), '', 'Target!A5', 'the target', 'Go to target')
    $data.Activate()
    $target.Range('2:3').Insert() | Out-Null
  }

  # --- deleted-cells-shift-up: the target cell is deleted, cells shift -----
  $deletedCellsPath = Save-CaseWorkbook -CaseName 'case-deleted-cells-shift-up' -BuildAndBreak {
    param($wb)
    $target = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
    $target.Name = 'Target'
    $data = $wb.Worksheets.Item(1)
    $data.Name = 'Data'
    $target.Range('A5').Value2 = 42
    $data.Range('A1').Value2 = 'Go to target'
    [void] $data.Hyperlinks.Add($data.Range('A1'), '', 'Target!A5', 'the target', 'Go to target')
    $data.Activate()
    $target.Range('A5').Delete(-4162) | Out-Null # xlShiftUp
  }

  foreach ($case in @(
      @{ Name = 'valid'; Path = $validPath },
      @{ Name = 'deleted-sheet'; Path = $deletedSheetPath },
      @{ Name = 'deleted-name'; Path = $deletedNamePath },
      @{ Name = 'never-existed'; Path = $neverExistedPath },
      @{ Name = 'deleted-row'; Path = $deletedRowPath },
      @{ Name = 'inserted-rows'; Path = $insertedRowsPath },
      @{ Name = 'deleted-cells-shift-up'; Path = $deletedCellsPath }
    )) {
    $worksheetXml = Read-ZipEntryText -ZipPath $case.Path -EntryName 'xl/worksheets/sheet1.xml'
    $relsXml = Read-ZipEntryText -ZipPath $case.Path -EntryName 'xl/worksheets/_rels/sheet1.xml.rels'
    $block = Get-HyperlinkBlock -WorksheetXml $worksheetXml
    $results.Add([pscustomobject]@{
        case = $case.Name
        file = [System.IO.Path]::GetFileName($case.Path)
        hyperlinksXml = $block
        sheetRelsXml = if ($null -eq $relsXml) { '(none)' } else { $relsXml }
      })
  }

  $manifest = [pscustomobject]@{
    schemaVersion = 1
    purpose = 'D-9 broken internal hyperlink capture (ticket 32, open question 3)'
    oracle = [pscustomobject]@{
      product = 'Microsoft Excel 365'
      version = $excel.Version
      build = $excel.Build
      os = [System.Environment]::OSVersion.VersionString
      locale = [System.Globalization.CultureInfo]::CurrentCulture.Name
    }
    generatedAt = [DateTime]::UtcNow.ToString('o')
    cases = $results
  }

  $reportPath = Join-Path $OutputDirectory 'broken-hyperlink-report.json'
  $json = $manifest | ConvertTo-Json -Depth 6
  $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
  [System.IO.File]::WriteAllText($reportPath, $json + "`n", $utf8NoBom)

  foreach ($result in $results) {
    Write-Output "=== $($result.case) ($($result.file)) ==="
    Write-Output $result.hyperlinksXml
    Write-Output '--- sheet rels ---'
    Write-Output $result.sheetRelsXml
    Write-Output ''
  }
  Write-Output "Wrote $reportPath"
  $captureSucceeded = $true
} finally {
  if ($null -ne $workbook) {
    try {
      $workbook.Close($false) | Out-Null
    } catch {
      $cleanupErrors.Add("Close failed: $($_.Exception.Message)")
    }
    Release-ComObject $workbook
  }
  if ($null -ne $excel) {
    try {
      $excel.Quit() | Out-Null
    } catch {
      $cleanupErrors.Add("Quit failed: $($_.Exception.Message)")
    }
    Release-ComObject $excel
  }
  if ($captureSucceeded -and $cleanupErrors.Count -gt 0) {
    Write-Warning "Capture succeeded but cleanup reported problems: $($cleanupErrors -join '; ')"
  }
  if ((-not $captureSucceeded) -and $cleanupErrors.Count -gt 0) {
    Write-Warning "Cleanup problems while the capture was already failing: $($cleanupErrors -join '; ')"
  }
}
