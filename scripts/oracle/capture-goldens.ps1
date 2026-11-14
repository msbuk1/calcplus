param(
  [Parameter(Mandatory = $true)]
  [string] $WorkbookPath,

  [Parameter(Mandatory = $true)]
  [string] $OutputPath
)

$ErrorActionPreference = 'Stop'

# Initialise every COM reference before the try/finally so cleanup is safe
# even when opening or capturing fails halfway through.
$excel = $null
$workbooks = $null
$workbook = $null
$worksheets = $null
$worksheet = $null
$usedRange = $null
$usedCells = $null
$cell = $null
$captureSucceeded = $false
$cleanupErrors = New-Object System.Collections.Generic.List[string]

function Get-ComHResultHex {
  param([System.Exception] $Exception)

  $current = $Exception
  while ($null -ne $current) {
    if ($current -is [System.Runtime.InteropServices.COMException]) {
      return ([int] $current.HResult).ToString('X8')
    }
    $current = $current.InnerException
  }
  return $null
}

function Invoke-ComActionWithRetry {
  param(
    [Parameter(Mandatory = $true)]
    [scriptblock] $Action,

    [int] $Attempts = 5,

    [int] $InitialDelayMilliseconds = 250
  )

  # Excel can reject a COM call while it is busy: VBA_E_IGNORE (0x800AC472),
  # RPC_E_CALL_REJECTED (0x80010001) or RPC_E_SERVERCALL_RETRYLATER
  # (0x8001010A). Retry those only, with a bounded backoff.
  $transientHResults = @('800AC472', '80010001', '8001010A')
  $delayMilliseconds = $InitialDelayMilliseconds

  for ($attempt = 1; $attempt -le $Attempts; $attempt++) {
    try {
      & $Action
      return
    } catch {
      $hresult = Get-ComHResultHex -Exception $_.Exception
      $isTransient = ($null -ne $hresult) -and ($transientHResults -contains $hresult)
      if ((-not $isTransient) -or ($attempt -eq $Attempts)) {
        throw
      }

      Start-Sleep -Milliseconds $delayMilliseconds
      $delayMilliseconds = [Math]::Min($delayMilliseconds * 2, 2000)
    }
  }
}

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

# Excel exposes a cell error as an Int32 CVErr in Value2 (often with the
# 0x800A0000 VBA error marker set), never as a string. Map it to the exact
# error names in testing-conformance §2.1 without reading Range.Text, which
# can be width-dependent (####).
$excelErrorNames = @{
  2000 = '#NULL!'
  2007 = '#DIV/0!'
  2015 = '#VALUE!'
  2023 = '#REF!'
  2029 = '#NAME?'
  2036 = '#NUM!'
  2042 = '#N/A'
  2043 = '#GETTING_DATA'
  2045 = '#SPILL!'
  2046 = '#CONNECT!'
  2047 = '#BLOCKED!'
  2048 = '#UNKNOWN!'
  2049 = '#FIELD!'
  2050 = '#CALC!'
}

function Format-JsonNumber {
  param([Parameter(Mandatory = $true)] [object] $Value)

  if (($Value -is [double]) -or ($Value -is [single])) {
    $number = [double] $Value
    if ([double]::IsNaN($number) -or [double]::IsInfinity($number)) {
      throw "Cannot write a non-finite number to JSON: $number"
    }
    # G17 gives the same 17-significant-digit representation on the .NET
    # Framework used by Windows PowerShell 5.1 and the .NET runtime used by
    # PowerShell 7; 'R' changed its output between those runtimes.
    return $number.ToString('G17', [System.Globalization.CultureInfo]::InvariantCulture)
  }
  if ($Value -is [decimal]) {
    return ([decimal] $Value).ToString([System.Globalization.CultureInfo]::InvariantCulture)
  }
  if (
    ($Value -is [byte]) -or ($Value -is [sbyte]) -or
    ($Value -is [int16]) -or ($Value -is [uint16]) -or
    ($Value -is [int32]) -or ($Value -is [uint32]) -or
    ($Value -is [int64]) -or ($Value -is [uint64])
  ) {
    return ([long] $Value).ToString([System.Globalization.CultureInfo]::InvariantCulture)
  }

  throw "Unsupported JSON number type: $($Value.GetType().FullName)"
}

function Add-JsonEscapedString {
  param(
    [Parameter(Mandatory = $true)] [System.Text.StringBuilder] $Builder,
    [Parameter(Mandatory = $true)] [AllowEmptyString()] [string] $Text
  )

  [void] $Builder.Append('"')
  for ($index = 0; $index -lt $Text.Length; $index++) {
    $character = $Text[$index]
    $code = [int] $character
    if ($character -eq '"') {
      [void] $Builder.Append('\"')
    } elseif ($character -eq '\') {
      [void] $Builder.Append('\\')
    } elseif ($code -eq 8) {
      [void] $Builder.Append('\b')
    } elseif ($code -eq 9) {
      [void] $Builder.Append('\t')
    } elseif ($code -eq 10) {
      [void] $Builder.Append('\n')
    } elseif ($code -eq 12) {
      [void] $Builder.Append('\f')
    } elseif ($code -eq 13) {
      [void] $Builder.Append('\r')
    } elseif ($code -lt 0x20) {
      [void] $Builder.Append('\u')
      [void] $Builder.Append($code.ToString('x4', [System.Globalization.CultureInfo]::InvariantCulture))
    } else {
      [void] $Builder.Append($character)
    }
  }
  [void] $Builder.Append('"')
}

function Add-JsonValue {
  param(
    [Parameter(Mandatory = $true)] [System.Text.StringBuilder] $Builder,
    [AllowNull()] [object] $Value,
    [Parameter(Mandatory = $true)] [int] $Depth
  )

  if ($null -eq $Value) {
    [void] $Builder.Append('null')
    return
  }
  if ($Value -is [bool]) {
    [void] $Builder.Append($(if ($Value) { 'true' } else { 'false' }))
    return
  }
  if ($Value -is [string]) {
    Add-JsonEscapedString -Builder $Builder -Text $Value
    return
  }
  if ($Value -is [System.Collections.IDictionary]) {
    $keys = @($Value.Keys)
    if ($keys.Count -eq 0) {
      [void] $Builder.Append('{}')
      return
    }
    [void] $Builder.Append("{`n")
    for ($index = 0; $index -lt $keys.Count; $index++) {
      [void] $Builder.Append((' ' * (($Depth + 1) * 2)))
      Add-JsonEscapedString -Builder $Builder -Text ([string] $keys[$index])
      [void] $Builder.Append(': ')
      Add-JsonValue -Builder $Builder -Value $Value[$keys[$index]] -Depth ($Depth + 1)
      if ($index -lt ($keys.Count - 1)) {
        [void] $Builder.Append(',')
      }
      [void] $Builder.Append("`n")
    }
    [void] $Builder.Append((' ' * ($Depth * 2)))
    [void] $Builder.Append('}')
    return
  }
  if ($Value -is [System.Collections.IEnumerable]) {
    $items = @($Value)
    if ($items.Count -eq 0) {
      [void] $Builder.Append('[]')
      return
    }
    [void] $Builder.Append("[`n")
    for ($index = 0; $index -lt $items.Count; $index++) {
      [void] $Builder.Append((' ' * (($Depth + 1) * 2)))
      Add-JsonValue -Builder $Builder -Value $items[$index] -Depth ($Depth + 1)
      if ($index -lt ($items.Count - 1)) {
        [void] $Builder.Append(',')
      }
      [void] $Builder.Append("`n")
    }
    [void] $Builder.Append((' ' * ($Depth * 2)))
    [void] $Builder.Append(']')
    return
  }
  if (
    ($Value -is [double]) -or ($Value -is [single]) -or ($Value -is [decimal]) -or
    ($Value -is [byte]) -or ($Value -is [sbyte]) -or
    ($Value -is [int16]) -or ($Value -is [uint16]) -or
    ($Value -is [int32]) -or ($Value -is [uint32]) -or
    ($Value -is [int64]) -or ($Value -is [uint64])
  ) {
    [void] $Builder.Append((Format-JsonNumber -Value $Value))
    return
  }

  throw "Unsupported JSON value type: $($Value.GetType().FullName)"
}

function ConvertTo-StableJson {
  param([Parameter(Mandatory = $true)] [AllowNull()] [object] $Value)

  $builder = New-Object System.Text.StringBuilder
  Add-JsonValue -Builder $builder -Value $Value -Depth 0
  return $builder.ToString()
}

if (-not (Test-Path -LiteralPath $WorkbookPath)) {
  throw "Oracle workbook not found: $WorkbookPath"
}

try {
  try {
    $excel = New-Object -ComObject Excel.Application
  } catch {
    throw "Excel 365 COM is unavailable on this PC. Install and activate desktop Excel 365 on this PC before capturing goldens. $($_.Exception.Message)"
  }

  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  # msoAutomationSecurityForceDisable = 3. Macro projects are never executed.
  $excel.AutomationSecurity = 3

  # Keep each COM collection in its own variable and use indexed loops.
  # `foreach` over an Excel collection leaves RCWs alive, and chained access
  # such as `$workbook.Application` leaks an extra Application RCW.
  $workbooks = $excel.Workbooks
  $resolvedWorkbookPath = (Resolve-Path -LiteralPath $WorkbookPath).Path
  try {
    $workbook = $workbooks.Open($resolvedWorkbookPath, 0, $true)
  } catch {
    $hresult = Get-ComHResultHex -Exception $_.Exception
    $suffix = if ($null -ne $hresult) { " (0x$hresult)" } else { '' }
    throw "Failed to open oracle workbook '$resolvedWorkbookPath' read-only$suffix`: $($_.Exception.Message)"
  }
  $excel.CalculateFullRebuild()

  $sheets = @()
  $worksheets = $workbook.Worksheets
  $worksheetCount = $worksheets.Count
  for ($worksheetIndex = 1; $worksheetIndex -le $worksheetCount; $worksheetIndex++) {
    $worksheet = $null
    $usedRange = $null
    $usedCells = $null
    try {
      $worksheet = $worksheets.Item($worksheetIndex)
      $usedRange = $worksheet.UsedRange
      $usedCells = $usedRange.Cells
      $cellCount = $usedCells.Count
      $cells = @()
      for ($cellIndex = 1; $cellIndex -le $cellCount; $cellIndex++) {
        $cell = $null
        try {
          $cell = $usedCells.Item($cellIndex)
          $hasFormula = [bool] $cell.HasFormula
          $value = $cell.Value2
          if (($null -eq $value) -and (-not $hasFormula)) {
            continue
          }

          $entry = [ordered]@{
            ref        = [string] $cell.Address($false, $false)
            hasFormula = $hasFormula
          }
          if ($hasFormula) {
            # Range.Formula is the exact Excel formula text. It is only
            # recorded when Excel says the cell is a formula, so a text cell
            # that happens to start with '=' stays a text value.
            $entry['formula'] = [string] $cell.Formula
          }

          if ($null -eq $value) {
            $entry['kind'] = 'blank'
            $entry['value'] = $null
          } elseif ($value -is [bool]) {
            $entry['kind'] = 'boolean'
            $entry['value'] = [bool] $value
          } elseif ($value -is [string]) {
            $entry['kind'] = 'text'
            $entry['value'] = $value
          } elseif ($value -is [int32]) {
            # Value2 encodes a CVErr as an Int32; mask the VBA error marker
            # (0x800A0000) to recover the XlCVError number.
            $errorCode = ([int] $value) -band 0xFFFF
            if (-not $excelErrorNames.ContainsKey($errorCode)) {
              throw "Unmapped Excel cell error code $errorCode at $($entry['ref'])."
            }
            $entry['kind'] = 'error'
            $entry['value'] = $excelErrorNames[$errorCode]
          } elseif (
            ($value -is [double]) -or ($value -is [single]) -or ($value -is [decimal]) -or
            ($value -is [byte]) -or ($value -is [sbyte]) -or
            ($value -is [int16]) -or ($value -is [uint16]) -or
            ($value -is [uint32]) -or ($value -is [int64]) -or ($value -is [uint64])
          ) {
            $entry['kind'] = 'number'
            $entry['value'] = [double] $value
          } else {
            throw "Unsupported Excel cell value type $($value.GetType().FullName) at $($entry['ref'])."
          }

          $cells += $entry
        } finally {
          Release-ComObject $cell
          $cell = $null
        }
      }
      $sheets += [ordered]@{ name = $worksheet.Name; cells = $cells }
    } finally {
      Release-ComObject $usedCells
      $usedCells = $null
      Release-ComObject $usedRange
      $usedRange = $null
      Release-ComObject $worksheet
      $worksheet = $null
    }
  }

  $manifest = [ordered]@{
    excelVersion = [string] $excel.Version
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
  if ([System.IO.Path]::IsPathRooted($OutputPath)) {
    $fullOutputPath = $OutputPath
  } else {
    $fullOutputPath = Join-Path (Get-Location).Path $OutputPath
  }
  # Write UTF-8 without a BOM and with our own stable serialiser: Windows
  # PowerShell 5.1 Set-Content -Encoding utf8 emits a BOM and ConvertTo-Json
  # differs from PowerShell 7 for some values.
  $json = ConvertTo-StableJson -Value $golden
  $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
  [System.IO.File]::WriteAllText($fullOutputPath, $json + "`n", $utf8NoBom)
  $captureSucceeded = $true
  Write-Output "Wrote $OutputPath"
} finally {
  # Release leaf objects first. Excel is only asked to quit after its child
  # RCWs are gone, so a rejected Quit is not caused by our own references.
  Release-ComObject $cell
  $cell = $null
  Release-ComObject $usedCells
  $usedCells = $null
  Release-ComObject $usedRange
  $usedRange = $null
  Release-ComObject $worksheet
  $worksheet = $null
  Release-ComObject $worksheets
  $worksheets = $null

  if ($null -ne $workbook) {
    try {
      Invoke-ComActionWithRetry -Action {
        $workbook.Close($false) | Out-Null
      }
    } catch {
      $hresult = Get-ComHResultHex -Exception $_.Exception
      $suffix = if ($null -ne $hresult) { " (0x$hresult)" } else { '' }
      $cleanupErrors.Add("Close failed$suffix`: $($_.Exception.Message)")
    }
  }

  Release-ComObject $workbook
  $workbook = $null
  Release-ComObject $workbooks
  $workbooks = $null

  if ($null -ne $excel) {
    try {
      Invoke-ComActionWithRetry -Action {
        $excel.Quit() | Out-Null
      }
    } catch {
      $hresult = Get-ComHResultHex -Exception $_.Exception
      $suffix = if ($null -ne $hresult) { " (0x$hresult)" } else { '' }
      $cleanupErrors.Add("Quit failed$suffix`: $($_.Exception.Message)")
    }
  }

  Release-ComObject $excel
  $excel = $null

  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()

  if ($cleanupErrors.Count -gt 0) {
    $message = "Excel COM cleanup failed after the capture: $($cleanupErrors -join '; ')"
    if ($captureSucceeded) {
      # The golden was written, but a persistent cleanup failure must not be
      # reported as a clean pass.
      throw $message
    }
    Write-Warning $message
  }
}
