param(
  [Parameter(Mandatory = $true)]
  [string] $EvidencePath,

  # Optional probe filter, matched with -like against the probe ids in
  # scripts/oracle/clipboard-spike.md. Omit it to run every probe.
  #   -Probe S03,T07
  #   -Probe 'S*'
  #   -Probe 'T0*','T10'
  [string[]] $Probe = @()
)

$ErrorActionPreference = 'Stop'

# =====================================================================
# calcplus M5 risk spike: Excel clipboard interop
#
# Answers the Excel-side open questions in
# .scratch/calcplus/research/06-clipboard-interop.md section 10 (1-5, 9, 10)
# against real Excel 365 on the maintainer's Windows PC, and captures the
# evidence that the clipboard contract (spec/clipboard.md) must rest on.
#
# Two probe families:
#   excel-source  Build a probe workbook in Excel, call Range.Copy, then read
#                 the Windows clipboard's registered "HTML Format" (CF_HTML)
#                 and CF_UNICODETEXT payloads with Win32 P/Invoke. Each raw
#                 payload is written to disk byte for byte.
#   excel-target  Write a CF_HTML payload in the calcplus dialect (and, where
#                 useful, a TSV fallback) to the clipboard, let Excel paste it
#                 into a clean sheet, then record a typed cell inventory of
#                 what Excel actually imported.
#
# This script cannot run anywhere but Windows with desktop Excel 365. It never
# executes macros (AutomationSecurity is forced off) and never touches the
# network. It fails loudly, with the COM HRESULT, when Excel COM or a clipboard
# format is unavailable; a payload that is present but empty (the single-cell
# Excel copy case) is an observation, not a failure.
#
# Runbook: scripts/oracle/clipboard-spike.md
# =====================================================================

if ([System.Environment]::OSVersion.Platform -ne [System.PlatformID]::Win32NT) {
  throw 'clipboard-spike.ps1 captures Win32 clipboard formats and drives desktop Excel 365 through COM. Run it on the maintainer''s Windows PC, not here.'
}

# Win32 clipboard entry points. Everything the script reads or writes goes
# through these; the wrapper functions below keep the handle lifetime rules
# explicit at the call sites.
$nativeDefinitions = @'
using System;
using System.Runtime.InteropServices;
using System.Text;

public static class NativeClipboard
{
  public const uint CF_TEXT = 1;
  public const uint CF_UNICODETEXT = 13;
  public const uint GMEM_MOVEABLE = 0x0002;

  [DllImport("user32.dll", SetLastError = true)]
  public static extern bool OpenClipboard(IntPtr hWndNewOwner);

  [DllImport("user32.dll", SetLastError = true)]
  public static extern bool CloseClipboard();

  [DllImport("user32.dll", SetLastError = true)]
  public static extern bool EmptyClipboard();

  [DllImport("user32.dll", SetLastError = true)]
  public static extern IntPtr GetClipboardData(uint uFormat);

  [DllImport("user32.dll", SetLastError = true)]
  public static extern IntPtr SetClipboardData(uint uFormat, IntPtr hMem);

  [DllImport("user32.dll", SetLastError = true)]
  public static extern bool IsClipboardFormatAvailable(uint format);

  [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
  public static extern uint RegisterClipboardFormat(string lpszFormat);

  [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
  public static extern int GetClipboardFormatName(uint format, StringBuilder lpszFormatName, int cchMaxCount);

  [DllImport("user32.dll", SetLastError = true)]
  public static extern uint EnumClipboardFormats(uint format);

  [DllImport("kernel32.dll", SetLastError = true)]
  public static extern IntPtr GlobalAlloc(uint uFlags, UIntPtr dwBytes);

  [DllImport("kernel32.dll", SetLastError = true)]
  public static extern IntPtr GlobalLock(IntPtr hMem);

  [DllImport("kernel32.dll", SetLastError = true)]
  public static extern bool GlobalUnlock(IntPtr hMem);

  [DllImport("kernel32.dll", SetLastError = true)]
  public static extern UIntPtr GlobalSize(IntPtr hMem);

  [DllImport("kernel32.dll", SetLastError = true)]
  public static extern IntPtr GlobalFree(IntPtr hMem);

  // PowerShell cannot build the UIntPtr size argument for GlobalAlloc without
  // an overload guess, so the single allocation call lives in C#.
  public static IntPtr AllocateMoveable(int byteCount)
  {
    return GlobalAlloc(GMEM_MOVEABLE, new UIntPtr((uint)byteCount));
  }
}
'@

$nativeTypeLoaded = $false
try {
  $nativeTypeLoaded = ($null -ne ([System.Management.Automation.PSTypeName]'NativeClipboard').Type)
} catch {
  $nativeTypeLoaded = $false
}
if (-not $nativeTypeLoaded) {
  try {
    Add-Type -TypeDefinition $nativeDefinitions -ErrorAction Stop
  } catch {
    # The script is normally run once per window, but tolerating "already
    # exists" keeps a dot-source or a second run in the same session working.
    if ($_.Exception.Message -notmatch 'already exists') {
      throw "Could not compile the Win32 clipboard P/Invoke block. Run this script in Windows PowerShell 5.1 on Windows. $($_.Exception.Message)"
    }
  }
}

# Initialise every long-lived COM reference before the try/finally so cleanup is
# safe even when a probe fails halfway through. Leaf objects (fonts, borders,
# cells, auxiliary ranges) are created and released inside the helper functions;
# these are the references the main blocks hold across a probe.
$excel = $null
$workbooks = $null
$workbook = $null
$worksheet = $null
$usedRange = $null
$range = $null
$destination = $null
$evidenceWritten = $false
$cleanupErrors = New-Object System.Collections.Generic.List[string]
$records = New-Object System.Collections.Generic.List[object]

# Windows PowerShell 5.1 Set-Content -Encoding utf8 writes a BOM, and Excel
# refuses a BOM-prefixed UTF-8 clipboard payload in some builds. Encode once.
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
# A throw-on-invalid encoder: the CF_HTML spec says UTF-8, so a payload that
# does not decode is itself an observation worth recording.
$strictUtf8 = New-Object System.Text.UTF8Encoding($false, $true)

# Xl constant names. The evidence JSON is read by a human writing a report, so
# enum values are translated; an unknown value falls back to its integer text.
$standardClipboardFormats = @{
  '1'  = 'CF_TEXT'
  '2'  = 'CF_BITMAP'
  '3'  = 'CF_METAFILEPICT'
  '4'  = 'CF_SYLK'
  '5'  = 'CF_DIF'
  '6'  = 'CF_TIFF'
  '7'  = 'CF_OEMTEXT'
  '8'  = 'CF_DIB'
  '9'  = 'CF_PALETTE'
  '10' = 'CF_PENDATA'
  '11' = 'CF_RIFF'
  '12' = 'CF_WAVE'
  '13' = 'CF_UNICODETEXT'
  '14' = 'CF_ENHMETAFILE'
  '15' = 'CF_HDROP'
  '16' = 'CF_LOCALE'
  '17' = 'CF_DIBV5'
}
$xlHorizontalAlignment = @{
  '1'     = 'general'
  '5'     = 'fill'
  '7'     = 'centerAcrossSelection'
  '-4108' = 'center'
  '-4117' = 'distributed'
  '-4130' = 'justify'
  '-4131' = 'left'
  '-4152' = 'right'
}
$xlVerticalAlignment = @{
  '-4107' = 'bottom'
  '-4108' = 'center'
  '-4117' = 'distributed'
  '-4130' = 'justify'
  '-4160' = 'top'
}
$xlLineStyle = @{
  '1'     = 'continuous'
  '4'     = 'dashDot'
  '5'     = 'dashDotDot'
  '13'    = 'slantDashDot'
  '-4115' = 'dash'
  '-4118' = 'dot'
  '-4119' = 'double'
  '-4142' = 'none'
}
$xlBorderWeight = @{
  '1'     = 'hairline'
  '2'     = 'thin'
  '4'     = 'thick'
  '-4138' = 'medium'
}
$xlPattern = @{
  '1'     = 'solid'
  '2'     = 'gray125'
  '-4105' = 'automatic'
  '-4124' = 'gray25'
  '-4125' = 'gray50'
  '-4126' = 'gray75'
  '-4142' = 'none'
}
$xlUnderline = @{
  '2'     = 'single'
  '4'     = 'singleAccounting'
  '5'     = 'doubleAccounting'
  '-4119' = 'double'
  '-4142' = 'none'
}
# XlBordersIndex: 7-10 are the four outer edges.
$borderIndexByName = @{
  EdgeLeft   = 7
  EdgeTop    = 8
  EdgeBottom = 9
  EdgeRight  = 10
}

# Excel exposes a cell error as an Int32 CVErr in Value2 (often with the
# 0x800A0000 VBA error marker set), never as a string. The same map as
# capture-goldens.ps1; Range.Text is never used for identity because it is
# width-dependent (####).
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

function New-ComFailureMessage {
  param(
    [Parameter(Mandatory = $true)] [string] $Context,
    [Parameter(Mandatory = $true)] [System.Exception] $Exception
  )

  $hresult = Get-ComHResultHex -Exception $Exception
  if ($null -ne $hresult) {
    return "$Context failed (COM HRESULT 0x$hresult): $($Exception.Message)"
  }
  return "$Context failed: $($Exception.Message)"
}

function Invoke-ComActionWithRetry {
  param(
    [Parameter(Mandatory = $true)]
    [scriptblock] $Action,

    [string] $Context = 'Excel COM call',

    [int] $Attempts = 5,

    [int] $InitialDelayMilliseconds = 250
  )

  # Excel can reject a COM call while it is busy: VBA_E_IGNORE (0x800AC472),
  # RPC_E_CALL_REJECTED (0x80010001) or RPC_E_SERVERCALL_RETRYLATER
  # (0x8001010A). Retry those only, with a bounded backoff. Everything else is
  # reported with its HRESULT so a failure names the call and the code.
  $transientHResults = @('800AC472', '80010001', '8001010A')
  $delayMilliseconds = $InitialDelayMilliseconds

  for ($attempt = 1; $attempt -le $Attempts; $attempt++) {
    try {
      # The action's output is returned so a caller can record what Excel said
      # (Worksheet.Paste returns True or False); callers that do not care
      # discard it with $null =.
      return (& $Action)
    } catch {
      $hresult = Get-ComHResultHex -Exception $_.Exception
      $isTransient = ($null -ne $hresult) -and ($transientHResults -contains $hresult)
      if ((-not $isTransient) -or ($attempt -eq $Attempts)) {
        throw (New-ComFailureMessage -Context $Context -Exception $_.Exception)
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

function ConvertTo-Utf8Bytes {
  param([Parameter(Mandatory = $true)] [AllowEmptyString()] [string] $Text)

  return $utf8NoBom.GetBytes($Text)
}

function ConvertTo-UnicodeTextBytes {
  param([Parameter(Mandatory = $true)] [AllowEmptyString()] [string] $Text)

  # CF_UNICODETEXT is UTF-16LE with a single terminating NUL code unit. Windows
  # applications (Excel included) read to the terminator, so a missing NUL
  # truncates the payload at whatever follows it in the block.
  return [System.Text.Encoding]::Unicode.GetBytes($Text + [string] [char] 0)
}

function ConvertFrom-UnicodeTextBytes {
  param([AllowNull()] [byte[]] $Bytes)

  if (($null -eq $Bytes) -or ($Bytes.Length -eq 0)) {
    return $null
  }

  # GlobalSize returns the allocation size, which can exceed the data length;
  # the first UTF-16 NUL code unit marks the end of the string.
  $length = $Bytes.Length
  for ($index = 0; $index -lt ($Bytes.Length - 1); $index += 2) {
    if (($Bytes[$index] -eq 0) -and ($Bytes[$index + 1] -eq 0)) {
      $length = $index
      break
    }
  }
  if ($length -eq 0) {
    return ''
  }
  if (($length % 2) -ne 0) {
    $length = $length - 1
  }
  return [System.Text.Encoding]::Unicode.GetString($Bytes, 0, $length)
}

function ConvertFrom-HtmlPayloadBytes {
  param([AllowNull()] [byte[]] $Bytes)

  $result = [ordered]@{
    text             = $null
    byteLength       = 0
    validUtf8        = $false
    suspectedUtf16Le = $false
    trimmedAtNul     = $false
  }
  if (($null -eq $Bytes) -or ($Bytes.Length -eq 0)) {
    return $result
  }
  $result['byteLength'] = $Bytes.Length

  # CF_HTML is specified as UTF-8 [Doc: HTML Clipboard Format]. If a build ever
  # writes UTF-16 the browser reads mojibake, so the hint is recorded rather
  # than silently decoded as if it were UTF-8. In UTF-16LE an ASCII character
  # has a zero high byte, so bytes 1 and 3 being zero is a reliable tell. The
  # raw file on disk is the authority in every case.
  if (($Bytes.Length -ge 4) -and ($Bytes[1] -eq 0) -and ($Bytes[3] -eq 0)) {
    $result['suspectedUtf16Le'] = $true
    $result['text'] = ([System.Text.Encoding]::Unicode.GetString($Bytes)).TrimEnd([char] 0)
    return $result
  }

  # Excel NUL-terminates the CF_HTML block and GlobalSize can return the larger
  # allocation size, so the first NUL byte ends the data.
  $length = $Bytes.Length
  for ($index = 0; $index -lt $Bytes.Length; $index++) {
    if ($Bytes[$index] -eq 0) {
      $length = $index
      $result['trimmedAtNul'] = $true
      break
    }
  }
  if ($length -eq 0) {
    return $result
  }

  try {
    $result['text'] = $strictUtf8.GetString($Bytes, 0, $length)
    $result['validUtf8'] = $true
  } catch {
    # A build that writes ANSI bytes is exactly the kind of finding this spike
    # exists to record. The raw file is still written verbatim; this decode is
    # only for the observation text.
    $result['text'] = [System.Text.Encoding]::Default.GetString($Bytes, 0, $length)
    $result['validUtf8'] = $false
  }
  return $result
}

function Format-CfHtmlOffset {
  param([Parameter(Mandatory = $true)] [int] $Value)

  # The header field is ten digits wide; the guard keeps the value inside both
  # that field and Int32, because a payload that big could not be on the
  # clipboard anyway.
  if ($Value -lt 0 -or $Value -gt 999999999) {
    throw "CF_HTML byte offset $Value does not fit the ten-digit header field."
  }
  return $Value.ToString('D10', [System.Globalization.CultureInfo]::InvariantCulture)
}

function ConvertTo-CfHtmlPayload {
  param([Parameter(Mandatory = $true)] [AllowEmptyString()] [string] $Html)

  # CF_HTML is UTF-8 with a plain-text ASCII header. Key mechanics:
  #  - StartHTML/EndHTML delimit the whole HTML document, StartFragment and
  #    EndFragment the pasteable slice between the <!--StartFragment--> markers.
  #  - Every offset is a BYTE offset from the first byte of the payload
  #    ("Version:" included), not a character offset.
  #  - Header lines are CRLF-terminated; some consumers also require CRLF
  #    between the wrapper's html/body tags.
  #  - The header length depends on the offsets, so it is first built with
  #    ten-digit zero placeholders. Swapping a placeholder for another
  #    ten-digit number cannot change the header length.
  #
  # Chromium wraps whatever a page wrote with text/html in exactly this
  # envelope, so a fragment that is itself a complete document (as calcplus
  # writes it) really does end up nested inside the wrapper. That is the shape
  # Excel receives from calcplus; the target probes test it deliberately.
  $htmlBytes = ConvertTo-Utf8Bytes -Text $Html
  $documentOpen = '<html><body>' + "`r`n" + '<!--StartFragment-->'
  $documentClose = '<!--EndFragment-->' + "`r`n" + '</body></html>'
  $openBytes = ConvertTo-Utf8Bytes -Text $documentOpen
  $closeBytes = ConvertTo-Utf8Bytes -Text $documentClose

  $placeholder = '0000000000'
  $templateHeader = 'Version:1.0' + "`r`n" +
    'StartHTML:' + $placeholder + "`r`n" +
    'EndHTML:' + $placeholder + "`r`n" +
    'StartFragment:' + $placeholder + "`r`n" +
    'EndFragment:' + $placeholder + "`r`n" +
    'SourceURL:about:blank' + "`r`n"
  $headerByteLength = (ConvertTo-Utf8Bytes -Text $templateHeader).Length

  $startHtml = $headerByteLength
  $startFragment = $startHtml + $openBytes.Length
  $endFragment = $startFragment + $htmlBytes.Length
  $endHtml = $endFragment + $closeBytes.Length

  $header = 'Version:1.0' + "`r`n" +
    'StartHTML:' + (Format-CfHtmlOffset -Value $startHtml) + "`r`n" +
    'EndHTML:' + (Format-CfHtmlOffset -Value $endHtml) + "`r`n" +
    'StartFragment:' + (Format-CfHtmlOffset -Value $startFragment) + "`r`n" +
    'EndFragment:' + (Format-CfHtmlOffset -Value $endFragment) + "`r`n" +
    'SourceURL:about:blank' + "`r`n"
  $prefixBytes = ConvertTo-Utf8Bytes -Text ($header + $documentOpen)
  if ($prefixBytes.Length -ne ($headerByteLength + $openBytes.Length)) {
    throw 'CF_HTML header length changed when the offsets were written; the template and the final header disagree.'
  }

  $payloadBytes = New-Object byte[] ($prefixBytes.Length + $htmlBytes.Length + $closeBytes.Length)
  [System.Array]::Copy($prefixBytes, 0, $payloadBytes, 0, $prefixBytes.Length)
  [System.Array]::Copy($htmlBytes, 0, $payloadBytes, $prefixBytes.Length, $htmlBytes.Length)
  [System.Array]::Copy($closeBytes, 0, $payloadBytes, $prefixBytes.Length + $htmlBytes.Length, $closeBytes.Length)

  return [ordered]@{
    bytes  = $payloadBytes
    header = [ordered]@{
      version       = '1.0'
      startHtml     = $startHtml
      endHtml       = $endHtml
      startFragment = $startFragment
      endFragment   = $endFragment
      headerBytes   = $headerByteLength
    }
  }
}

function Get-CfHtmlHeader {
  param([AllowNull()] [byte[]] $Bytes)

  $result = [ordered]@{
    present             = $false
    version             = $null
    startHtml           = $null
    endHtml             = $null
    startFragment       = $null
    endFragment         = $null
    sourceUrl           = $null
    offsetsConsistent   = $false
    markersPresent      = $false
    fragmentByteLength  = 0
    fragment            = $null
    fragmentValidUtf8   = $false
  }
  if (($null -eq $Bytes) -or ($Bytes.Length -eq 0)) {
    return $result
  }

  # The header is ASCII, so only the first 2 KB are decoded: multi-byte
  # fragment content must not be able to shift the parse.
  $prefixLength = [Math]::Min(2048, $Bytes.Length)
  $prefix = [System.Text.Encoding]::ASCII.GetString($Bytes, 0, $prefixLength)
  if (-not ($prefix -match '^\s*Version:')) {
    return $result
  }
  $result['present'] = $true
  $result['markersPresent'] = $prefix.Contains('<!--StartFragment-->') -and $prefix.Contains('<!--EndFragment-->')

  $pattern = '(?m)^(?<name>Version|StartHTML|EndHTML|StartFragment|EndFragment|StartSelection|EndSelection|SourceURL)\s*:\s*(?<value>[^\r\n]*)'
  $fields = @{}
  foreach ($match in [regex]::Matches($prefix, $pattern)) {
    $fields[$match.Groups['name'].Value] = $match.Groups['value'].Value.Trim()
  }
  if ($fields.ContainsKey('Version')) {
    $result['version'] = $fields['Version']
  }
  if ($fields.ContainsKey('SourceURL')) {
    $result['sourceUrl'] = $fields['SourceURL']
  }

  $offsetNames = @(
    @{ Field = 'StartHTML'; Key = 'startHtml' },
    @{ Field = 'EndHTML'; Key = 'endHtml' },
    @{ Field = 'StartFragment'; Key = 'startFragment' },
    @{ Field = 'EndFragment'; Key = 'endFragment' }
  )
  foreach ($offsetName in $offsetNames) {
    if (-not $fields.ContainsKey($offsetName.Field)) {
      continue
    }
    $parsed = 0
    if ([int]::TryParse($fields[$offsetName.Field], [ref] $parsed)) {
      $result[$offsetName.Key] = $parsed
    }
  }

  if (
    ($null -ne $result['startHtml']) -and ($null -ne $result['endHtml']) -and
    ($null -ne $result['startFragment']) -and ($null -ne $result['endFragment'])
  ) {
    $envelopeValid = (
      ($result['startHtml'] -ge 0) -and
      ($result['startHtml'] -le $result['startFragment']) -and
      ($result['startFragment'] -le $result['endFragment']) -and
      ($result['endFragment'] -le $result['endHtml']) -and
      ($result['endHtml'] -le $Bytes.Length)
    )
    $result['offsetsConsistent'] = $envelopeValid

    if ($envelopeValid) {
      $fragmentLength = $result['endFragment'] - $result['startFragment']
      $result['fragmentByteLength'] = $fragmentLength
      if ($fragmentLength -gt 0) {
        $fragmentBytes = New-Object byte[] $fragmentLength
        [System.Array]::Copy($Bytes, $result['startFragment'], $fragmentBytes, 0, $fragmentLength)
        try {
          $result['fragment'] = $strictUtf8.GetString($fragmentBytes)
          $result['fragmentValidUtf8'] = $true
        } catch {
          # Offsets that do not land on a UTF-8 boundary are themselves the
          # observation; leave the fragment text null and keep the offsets.
          $result['fragment'] = $null
          $result['fragmentValidUtf8'] = $false
        }
      }
    }
  }
  return $result
}

function Add-DistinctMatches {
  param(
    [Parameter(Mandatory = $true)] [System.Collections.Generic.List[string]] $Values,
    [Parameter(Mandatory = $true)] [AllowEmptyString()] [string] $Text,
    [Parameter(Mandatory = $true)] [string] $Pattern,
    [int] $Limit = 12
  )

  foreach ($match in [regex]::Matches($Text, $Pattern)) {
    $value = $match.Groups[1].Value
    if (-not $Values.Contains($value)) {
      $Values.Add($value)
      if ($Values.Count -ge $Limit) {
        break
      }
    }
  }
}

function Get-HtmlPayloadObservation {
  param(
    [AllowNull()] [string] $Fragment,
    [int] $MaxDetailLength = 2000000
  )

  if ($null -eq $Fragment) {
    $Fragment = ''
  }
  $isLarge = ($Fragment.Length -gt $MaxDetailLength)

  $observation = [ordered]@{
    characterLength          = $Fragment.Length
    detailLimitedToCounts    = $isLarge
    hasHtmlElement           = $Fragment.Contains('<html')
    hasProgIdMeta            = $Fragment.Contains('ProgId')
    hasExcelNamespace        = $Fragment.Contains('urn:schemas-microsoft-com:office:excel')
    hasOfficeNamespace       = $Fragment.Contains('urn:schemas-microsoft-com:office:office')
    hasStyleElement          = $Fragment.Contains('<style')
    hasTable                 = $Fragment.Contains('<table')
    hasStartFragmentMarker   = $Fragment.Contains('<!--StartFragment-->')
    hasHeightSourceStyle     = $Fragment.Contains('mso-height-source')
    hasSameCellDataPlacement = $Fragment.Contains('mso-data-placement')
    hasScriptElement         = [regex]::IsMatch($Fragment, '<script', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    hasEventHandlerAttribute = [regex]::IsMatch($Fragment, '\son[a-z]+\s*=', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    programId                = $null
    generator                = $null
    styleBlockCharacterLength = 0
    tableCellCount           = 0
    formulaAttributeCount    = 0
    numericAttributeCount    = 0
    textAttributeCount       = 0
    booleanAttributeCount    = 0
    errorAttributeCount      = 0
    formulaExamples          = @()
    numericExamples          = @()
    errorValues              = @()
    numberFormats            = @()
    colSpanValues            = @()
    rowSpanValues            = @()
    heightValues             = @()
    anchorHrefs              = @()
  }

  # Counts are cheap enough even for the 50,000-cell payload; the per-value
  # collections are not, so a large payload records counts only.
  $observation['tableCellCount'] = ([regex]::Matches($Fragment, '<td')).Count
  $observation['formulaAttributeCount'] = ([regex]::Matches($Fragment, 'x:fmla=')).Count

  if (-not $isLarge) {
    $observation['numericAttributeCount'] = ([regex]::Matches($Fragment, 'x:num')).Count
    $observation['textAttributeCount'] = ([regex]::Matches($Fragment, 'x:str')).Count
    $observation['booleanAttributeCount'] = ([regex]::Matches($Fragment, 'x:bool')).Count
    $observation['errorAttributeCount'] = ([regex]::Matches($Fragment, 'x:err')).Count

    $progIdMatch = [regex]::Match($Fragment, 'ProgId\s+content\s*=\s*"?([^"\s>]+)')
    if ($progIdMatch.Success) {
      $observation['programId'] = $progIdMatch.Groups[1].Value
    }
    $generatorMatch = [regex]::Match($Fragment, 'Generator\s+content\s*=\s*"?([^"\s>]+(?:\s+[^"\s>]+)*)')
    if ($generatorMatch.Success) {
      $observation['generator'] = $generatorMatch.Groups[1].Value
    }
    $styleMatch = [regex]::Match($Fragment, '(?s)<style[^>]*>(.*?)</style>')
    if ($styleMatch.Success) {
      $observation['styleBlockCharacterLength'] = $styleMatch.Groups[1].Value.Length
    }

    $formulaExamples = New-Object System.Collections.Generic.List[string]
    Add-DistinctMatches -Values $formulaExamples -Text $Fragment -Pattern 'x:fmla="([^"]*)"' -Limit 8
    $observation['formulaExamples'] = $formulaExamples.ToArray()

    $numericExamples = New-Object System.Collections.Generic.List[string]
    Add-DistinctMatches -Values $numericExamples -Text $Fragment -Pattern 'x:num="([^"]*)"' -Limit 5
    $observation['numericExamples'] = $numericExamples.ToArray()

    $errorValues = New-Object System.Collections.Generic.List[string]
    Add-DistinctMatches -Values $errorValues -Text $Fragment -Pattern 'x:err="([^"]*)"' -Limit 8
    $observation['errorValues'] = $errorValues.ToArray()

    # mso-number-format rides in CSS, either as a class rule or inline. The
    # format code itself can contain escaped commas, hashes and dots, so the
    # capture runs to the next CSS separator or quote rather than to a token.
    $numberFormats = New-Object System.Collections.Generic.List[string]
    Add-DistinctMatches -Values $numberFormats -Text $Fragment -Pattern 'mso-number-format\s*:\s*([^;}]+)' -Limit 12
    $trimmedNumberFormats = New-Object System.Collections.Generic.List[string]
    foreach ($numberFormat in $numberFormats) {
      $trimmedNumberFormats.Add($numberFormat.Trim().Trim('"'))
    }
    $observation['numberFormats'] = $trimmedNumberFormats.ToArray()

    $colSpans = New-Object System.Collections.Generic.List[string]
    Add-DistinctMatches -Values $colSpans -Text $Fragment -Pattern 'colspan\s*=\s*"?([0-9]+)' -Limit 8
    $observation['colSpanValues'] = $colSpans.ToArray()

    $rowSpans = New-Object System.Collections.Generic.List[string]
    Add-DistinctMatches -Values $rowSpans -Text $Fragment -Pattern 'rowspan\s*=\s*"?([0-9]+)' -Limit 8
    $observation['rowSpanValues'] = $rowSpans.ToArray()

    $heights = New-Object System.Collections.Generic.List[string]
    Add-DistinctMatches -Values $heights -Text $Fragment -Pattern 'height\s*[:=]\s*"?([0-9.]+)' -Limit 8
    $observation['heightValues'] = $heights.ToArray()

    $anchors = New-Object System.Collections.Generic.List[string]
    Add-DistinctMatches -Values $anchors -Text $Fragment -Pattern 'href\s*=\s*"([^"]*)"' -Limit 5
    $observation['anchorHrefs'] = $anchors.ToArray()
  }

  return $observation
}

function Get-ClipboardFormatPayload {
  param([Parameter(Mandatory = $true)] [uint32] $Format)

  # The caller must already hold the clipboard open: the HGLOBAL that
  # GetClipboardData returns belongs to the clipboard and is valid only until
  # CloseClipboard, so the bytes are copied out immediately.
  $payload = [ordered]@{
    available  = $false
    bytes      = $null
    byteLength = 0
  }

  $handle = [NativeClipboard]::GetClipboardData($Format)
  if ($handle -eq [System.IntPtr]::Zero) {
    # Absent is not an error: a single-cell Excel copy is allowed to omit the
    # HTML flavour, and that absence is exactly what the probe records.
    return $payload
  }
  $payload['available'] = $true

  $pointer = [NativeClipboard]::GlobalLock($handle)
  if ($pointer -eq [System.IntPtr]::Zero) {
    $win32Error = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
    throw "GlobalLock failed for clipboard format $Format (Win32 error $win32Error)."
  }
  try {
    $size = [int64] ([NativeClipboard]::GlobalSize($handle)).ToUInt64()
    if ($size -le 0) {
      $payload['bytes'] = New-Object byte[] 0
      return $payload
    }
    $buffer = New-Object byte[] ([int] $size)
    [System.Runtime.InteropServices.Marshal]::Copy($pointer, $buffer, 0, [int] $size)
    $payload['bytes'] = $buffer
    $payload['byteLength'] = [int] $size
  } finally {
    # GlobalUnlock returns false when the last lock is released; that is
    # success, so the return value is deliberately ignored.
    [void] [NativeClipboard]::GlobalUnlock($handle)
  }
  return $payload
}

function Get-ClipboardFormatList {
  # The enumeration order is the order the clipboard owner placed the formats,
  # which is roughly the priority Excel applies when it pastes into itself.
  # Bounded because clipboard contents are untrusted input.
  $formats = @()
  $current = [uint32] 0
  for ($index = 0; $index -lt 256; $index++) {
    $current = [NativeClipboard]::EnumClipboardFormats($current)
    if ($current -eq 0) {
      break
    }
    $name = $null
    if ($current -lt 0xC000) {
      $key = ([int] $current).ToString([System.Globalization.CultureInfo]::InvariantCulture)
      if ($standardClipboardFormats.ContainsKey($key)) {
        $name = [string] $standardClipboardFormats[$key]
      }
    } else {
      $builder = New-Object System.Text.StringBuilder 256
      $length = [NativeClipboard]::GetClipboardFormatName($current, $builder, 256)
      if ($length -gt 0) {
        $name = $builder.ToString()
      }
    }
    if ([string]::IsNullOrEmpty($name)) {
      $name = '(unnamed)'
    }
    $formats += [ordered]@{
      id    = [int] $current
      idHex = '0x' + ([int] $current).ToString('X4', [System.Globalization.CultureInfo]::InvariantCulture)
      name  = $name
    }
  }
  return , $formats
}

function Invoke-WithClipboardOpen {
  param(
    [Parameter(Mandatory = $true)] [scriptblock] $Action,
    [AllowNull()] [object] $Argument,
    [int] $Attempts = 20,
    [int] $InitialDelayMilliseconds = 50
  )

  # OpenClipboard fails with ERROR_ACCESS_DENIED (5) while Excel, a clipboard
  # manager or the shell holds the clipboard open. Excel releases it within
  # milliseconds of Range.Copy returning, so a short bounded retry is enough; a
  # permanent failure means something else owns the clipboard.
  $delayMilliseconds = $InitialDelayMilliseconds
  $opened = $false
  for ($attempt = 1; $attempt -le $Attempts; $attempt++) {
    if ([NativeClipboard]::OpenClipboard([System.IntPtr]::Zero)) {
      $opened = $true
      break
    }
    $win32Error = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
    if ($win32Error -ne 5) {
      throw "OpenClipboard failed with Win32 error $win32Error."
    }
    if ($attempt -eq $Attempts) {
      throw "OpenClipboard failed with ERROR_ACCESS_DENIED after $Attempts attempts: another process is holding the Windows clipboard open. Close clipboard managers and rerun."
    }
    Start-Sleep -Milliseconds $delayMilliseconds
    $delayMilliseconds = [Math]::Min($delayMilliseconds * 2, 500)
  }
  if (-not $opened) {
    throw 'OpenClipboard failed for an unknown reason.'
  }

  try {
    return (& $Action $Argument)
  } finally {
    [void] [NativeClipboard]::CloseClipboard()
  }
}

function Get-ClipboardSnapshot {
  param(
    [switch] $IncludeFormatList,
    [Parameter(Mandatory = $true)] [uint32] $HtmlFormat
  )

  # One open for the whole snapshot: both HGLOBALs must be read before the
  # clipboard is closed.
  $request = [ordered]@{
    includeFormatList = [bool] $IncludeFormatList
    htmlFormat         = $HtmlFormat
  }
  return (Invoke-WithClipboardOpen -Argument $request -Action {
    param($request)

    $requestedFormat = [uint32] ($request['htmlFormat'])
    $snapshot = [ordered]@{
      htmlFormatId         = [int] $requestedFormat
      htmlAvailable        = [bool] [NativeClipboard]::IsClipboardFormatAvailable($requestedFormat)
      html                 = $null
      unicodeTextAvailable = [bool] [NativeClipboard]::IsClipboardFormatAvailable([uint32] 13)
      unicodeText          = $null
      formats              = $null
    }
    $snapshot['html'] = Get-ClipboardFormatPayload -Format $requestedFormat
    $snapshot['unicodeText'] = Get-ClipboardFormatPayload -Format ([uint32] 13)
    if ($request['includeFormatList']) {
      $snapshot['formats'] = Get-ClipboardFormatList
    }
    return $snapshot
  })
}

function Set-ClipboardFormatData {
  param(
    [Parameter(Mandatory = $true)] [uint32] $Format,
    [Parameter(Mandatory = $true)] [byte[]] $Bytes
  )

  if ($Bytes.Length -eq 0) {
    throw "Refusing to publish an empty clipboard payload for format $Format."
  }

  # The clipboard takes ownership of the HGLOBAL when SetClipboardData
  # succeeds; on failure the caller must free it, so success and failure are
  # handled separately here.
  $handle = [NativeClipboard]::AllocateMoveable($Bytes.Length)
  if ($handle -eq [System.IntPtr]::Zero) {
    $win32Error = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
    throw "GlobalAlloc failed for clipboard format $Format (Win32 error $win32Error)."
  }

  $pointer = [NativeClipboard]::GlobalLock($handle)
  if ($pointer -eq [System.IntPtr]::Zero) {
    $win32Error = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
    [void] [NativeClipboard]::GlobalFree($handle)
    throw "GlobalLock failed for clipboard format $Format (Win32 error $win32Error)."
  }
  try {
    [System.Runtime.InteropServices.Marshal]::Copy($Bytes, 0, $pointer, $Bytes.Length)
  } finally {
    [void] [NativeClipboard]::GlobalUnlock($handle)
  }

  $published = [NativeClipboard]::SetClipboardData($Format, $handle)
  if ($published -eq [System.IntPtr]::Zero) {
    $win32Error = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
    [void] [NativeClipboard]::GlobalFree($handle)
    throw "SetClipboardData failed for clipboard format $Format (Win32 error $win32Error)."
  }
  # On success the system owns the handle; freeing it here would corrupt the
  # clipboard, so it is deliberately not freed.
}

function Set-ClipboardPayload {
  param(
    # [object] rather than [string]: PowerShell would coerce a $null [string]
    # argument to '', and "no HTML flavour" must stay distinguishable from "an
    # empty HTML flavour" for the text-only probes.
    [AllowNull()] [object] $Html,
    [AllowNull()] [object] $Text,
    [Parameter(Mandatory = $true)] [uint32] $HtmlFormat
  )

  if (($null -eq $Html) -and ($null -eq $Text)) {
    throw 'Refusing to set an empty clipboard payload.'
  }

  $htmlPayloadBytes = $null
  $htmlHeader = $null
  if ($null -ne $Html) {
    $cfHtml = ConvertTo-CfHtmlPayload -Html ([string] $Html)
    $htmlPayloadBytes = $cfHtml['bytes']
    $htmlHeader = $cfHtml['header']
  }
  $textValue = $null
  $textBytes = $null
  if ($null -ne $Text) {
    $textValue = [string] $Text
    $textBytes = ConvertTo-UnicodeTextBytes -Text $textValue
  }

  $request = [ordered]@{
    htmlBytes  = $htmlPayloadBytes
    htmlFormat = $HtmlFormat
    textBytes  = $textBytes
  }
  $null = Invoke-WithClipboardOpen -Argument $request -Action {
    param($request)

    # EmptyClipboard must run with the clipboard open; it takes ownership from
    # whoever held it (Excel, in the target loop) and discards the old formats,
    # so all new formats must be published before CloseClipboard.
    if (-not [NativeClipboard]::EmptyClipboard()) {
      $win32Error = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
      throw "EmptyClipboard failed (Win32 error $win32Error)."
    }
    if ($null -ne $request['htmlBytes']) {
      Set-ClipboardFormatData -Format ([uint32] $request['htmlFormat']) -Bytes $request['htmlBytes']
    }
    if ($null -ne $request['textBytes']) {
      Set-ClipboardFormatData -Format ([uint32] 13) -Bytes $request['textBytes']
    }
    return $true
  }

  # Read back what we just published. A silent clipboard write would make every
  # target probe meaningless, so a mismatch is a hard failure.
  $verification = Get-ClipboardSnapshot -HtmlFormat $HtmlFormat
  $verified = $false
  if (($null -ne $htmlPayloadBytes) -and $verification['html']['available']) {
    $publishedBytes = $verification['html']['bytes']
    $verified = ($publishedBytes.Length -eq $htmlPayloadBytes.Length)
    if ($verified) {
      for ($index = 0; $index -lt $htmlPayloadBytes.Length; $index++) {
        if ($publishedBytes[$index] -ne $htmlPayloadBytes[$index]) {
          $verified = $false
          break
        }
      }
    }
  } elseif (($null -eq $htmlPayloadBytes) -and ($null -ne $textBytes)) {
    $publishedText = ConvertFrom-UnicodeTextBytes -Bytes ($verification['unicodeText']['bytes'])
    $verified = ($publishedText -eq $textValue)
  }
  if (-not $verified) {
    throw 'The clipboard does not contain the payload that was just written, so a paste probe would measure the wrong thing.'
  }

  $formatNames = @()
  if ($null -ne $htmlPayloadBytes) {
    $formatNames += 'HTML Format'
  }
  if ($null -ne $textBytes) {
    $formatNames += 'CF_UNICODETEXT'
  }

  return [ordered]@{
    formats      = $formatNames
    htmlByteLength = $(if ($null -eq $htmlPayloadBytes) { 0 } else { $htmlPayloadBytes.Length })
    textByteLength = $(if ($null -eq $textBytes) { 0 } else { $textBytes.Length })
    cfHtmlHeader = $htmlHeader
    verified     = $verified
  }
}

function Get-ComPropertyValue {
  param(
    [AllowNull()] [object] $Object,
    [Parameter(Mandatory = $true)] [string] $PropertyName
  )

  if ($null -eq $Object) {
    return $null
  }
  try {
    return $Object.$PropertyName
  } catch {
    # A property Excel refuses for this range (mixed values, a merged-covered
    # cell) is recorded as null rather than aborting a whole probe.
    return $null
  }
}

function ConvertTo-TriStateBool {
  param([AllowNull()] [object] $Value)

  if ($null -eq $Value) {
    return $null
  }
  if ($Value -is [bool]) {
    return [bool] $Value
  }
  $text = [string] $Value
  if ($text -eq '') {
    return $null
  }
  if ($text -eq 'True') {
    return $true
  }
  if ($text -eq 'False') {
    return $false
  }
  $number = 0.0
  if ([double]::TryParse($text, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref] $number)) {
    return ($number -ne 0)
  }
  return $null
}

function ConvertTo-NullableNumber {
  param([AllowNull()] [object] $Value)

  if ($null -eq $Value) {
    return $null
  }
  if ($Value -is [bool]) {
    return $null
  }
  try {
    return [double] $Value
  } catch {
    return $null
  }
}

function Get-XlConstantName {
  param(
    [Parameter(Mandatory = $true)] [hashtable] $Table,
    [AllowNull()] [object] $Value
  )

  if ($null -eq $Value) {
    return $null
  }
  $number = 0.0
  if (-not [double]::TryParse([string] $Value, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref] $number)) {
    return [string] $Value
  }
  $key = ([int] $number).ToString([System.Globalization.CultureInfo]::InvariantCulture)
  if ($Table.ContainsKey($key)) {
    return [string] $Table[$key]
  }
  return $key
}

function Add-CellValueFields {
  param(
    [Parameter(Mandatory = $true)] [object] $Cell,
    [Parameter(Mandatory = $true)] [System.Collections.IDictionary] $Record
  )

  $hasFormula = [bool] (Get-ComPropertyValue -Object $Cell -PropertyName 'HasFormula')
  $value = Get-ComPropertyValue -Object $Cell -PropertyName 'Value2'

  $Record['hasFormula'] = $hasFormula
  if ($hasFormula) {
    $formula = Get-ComPropertyValue -Object $Cell -PropertyName 'Formula'
    if ($null -eq $formula) {
      $Record['formula'] = $null
    } else {
      $Record['formula'] = [string] $formula
    }
  } else {
    $Record['formula'] = $null
  }

  if ($null -eq $value) {
    $Record['kind'] = 'blank'
    $Record['value'] = $null
  } elseif ($value -is [bool]) {
    $Record['kind'] = 'boolean'
    $Record['value'] = [bool] $value
  } elseif ($value -is [string]) {
    $Record['kind'] = 'text'
    $Record['value'] = [string] $value
  } elseif ($value -is [int32]) {
    # Value2 encodes a CVErr as an Int32 with the VBA error marker set. Mask it
    # and fall back to a number when the code is not a known XlCVError.
    $errorCode = ([int] $value) -band 0xFFFF
    if ($excelErrorNames.ContainsKey($errorCode)) {
      $Record['kind'] = 'error'
      $Record['value'] = [string] $excelErrorNames[$errorCode]
    } else {
      $Record['kind'] = 'number'
      $Record['value'] = [double] $value
    }
  } elseif (
    ($value -is [double]) -or ($value -is [single]) -or ($value -is [decimal]) -or
    ($value -is [byte]) -or ($value -is [sbyte]) -or
    ($value -is [int16]) -or ($value -is [uint16]) -or
    ($value -is [uint32]) -or ($value -is [int64]) -or ($value -is [uint64])
  ) {
    $Record['kind'] = 'number'
    $Record['value'] = [double] $value
  } else {
    $Record['kind'] = 'unknown'
    $Record['value'] = [string] $value
  }
}

function Get-CellRecord {
  param([Parameter(Mandatory = $true)] [object] $Cell)

  $record = [ordered]@{
    ref = [string] ($Cell.Address($false, $false))
  }
  Add-CellValueFields -Cell $Cell -Record $record

  $numberFormat = Get-ComPropertyValue -Object $Cell -PropertyName 'NumberFormat'
  if ($null -eq $numberFormat) {
    $record['numberFormat'] = $null
  } else {
    $record['numberFormat'] = [string] $numberFormat
  }

  $font = $null
  try {
    $font = $Cell.Font
    $record['font'] = [ordered]@{
      name       = [string] (Get-ComPropertyValue -Object $font -PropertyName 'Name')
      size       = ConvertTo-NullableNumber -Value (Get-ComPropertyValue -Object $font -PropertyName 'Size')
      bold       = ConvertTo-TriStateBool -Value (Get-ComPropertyValue -Object $font -PropertyName 'Bold')
      italic     = ConvertTo-TriStateBool -Value (Get-ComPropertyValue -Object $font -PropertyName 'Italic')
      underline  = Get-XlConstantName -Table $xlUnderline -Value (Get-ComPropertyValue -Object $font -PropertyName 'Underline')
      colorIndex = ConvertTo-NullableNumber -Value (Get-ComPropertyValue -Object $font -PropertyName 'ColorIndex')
    }
  } catch {
    $record['font'] = $null
  } finally {
    Release-ComObject $font
    $font = $null
  }

  $interior = $null
  try {
    $interior = $Cell.Interior
    $record['fill'] = [ordered]@{
      pattern           = Get-XlConstantName -Table $xlPattern -Value (Get-ComPropertyValue -Object $interior -PropertyName 'Pattern')
      colorIndex        = ConvertTo-NullableNumber -Value (Get-ComPropertyValue -Object $interior -PropertyName 'ColorIndex')
      patternColorIndex = ConvertTo-NullableNumber -Value (Get-ComPropertyValue -Object $interior -PropertyName 'PatternColorIndex')
    }
  } catch {
    $record['fill'] = $null
  } finally {
    Release-ComObject $interior
    $interior = $null
  }

  $borderRecords = @()
  $borderCollection = $null
  $borderObject = $null
  try {
    $borderCollection = $Cell.Borders
    foreach ($edgeName in @('EdgeLeft', 'EdgeTop', 'EdgeBottom', 'EdgeRight')) {
      $borderObject = $borderCollection.Item($borderIndexByName[$edgeName])
      $borderRecords += [ordered]@{
        edge       = $edgeName
        lineStyle  = Get-XlConstantName -Table $xlLineStyle -Value (Get-ComPropertyValue -Object $borderObject -PropertyName 'LineStyle')
        weight     = Get-XlConstantName -Table $xlBorderWeight -Value (Get-ComPropertyValue -Object $borderObject -PropertyName 'Weight')
        colorIndex = ConvertTo-NullableNumber -Value (Get-ComPropertyValue -Object $borderObject -PropertyName 'ColorIndex')
      }
      Release-ComObject $borderObject
      $borderObject = $null
    }
  } catch {
    $borderRecords += [ordered]@{ edge = 'unavailable'; lineStyle = $null; weight = $null; colorIndex = $null }
  } finally {
    Release-ComObject $borderObject
    $borderObject = $null
    Release-ComObject $borderCollection
    $borderCollection = $null
  }
  $record['borders'] = $borderRecords

  $record['alignment'] = [ordered]@{
    horizontal  = Get-XlConstantName -Table $xlHorizontalAlignment -Value (Get-ComPropertyValue -Object $Cell -PropertyName 'HorizontalAlignment')
    vertical    = Get-XlConstantName -Table $xlVerticalAlignment -Value (Get-ComPropertyValue -Object $Cell -PropertyName 'VerticalAlignment')
    wrapText    = ConvertTo-TriStateBool -Value (Get-ComPropertyValue -Object $Cell -PropertyName 'WrapText')
    indentLevel = ConvertTo-NullableNumber -Value (Get-ComPropertyValue -Object $Cell -PropertyName 'IndentLevel')
    orientation = ConvertTo-NullableNumber -Value (Get-ComPropertyValue -Object $Cell -PropertyName 'Orientation')
  }

  $merged = ConvertTo-TriStateBool -Value (Get-ComPropertyValue -Object $Cell -PropertyName 'MergeCells')
  $mergeArea = $null
  if ($merged -eq $true) {
    $mergeRange = $null
    try {
      $mergeRange = $Cell.MergeArea
      $mergeArea = [string] ($mergeRange.Address($false, $false))
    } finally {
      Release-ComObject $mergeRange
      $mergeRange = $null
    }
  }
  $record['merge'] = [ordered]@{
    merged = $merged
    area   = $mergeArea
  }

  $record['rowHeight'] = ConvertTo-NullableNumber -Value (Get-ComPropertyValue -Object $Cell -PropertyName 'RowHeight')

  $hyperlinkRecord = [ordered]@{ count = 0 }
  $hyperlinkCollection = $null
  $hyperlinkObject = $null
  try {
    $hyperlinkCollection = $Cell.Hyperlinks
    $hyperlinkCount = [int] (Get-ComPropertyValue -Object $hyperlinkCollection -PropertyName 'Count')
    $hyperlinkRecord['count'] = $hyperlinkCount
    if ($hyperlinkCount -gt 0) {
      $hyperlinkObject = $hyperlinkCollection.Item(1)
      $hyperlinkRecord['address'] = [string] (Get-ComPropertyValue -Object $hyperlinkObject -PropertyName 'Address')
      $hyperlinkRecord['subAddress'] = [string] (Get-ComPropertyValue -Object $hyperlinkObject -PropertyName 'SubAddress')
      $hyperlinkRecord['textToDisplay'] = [string] (Get-ComPropertyValue -Object $hyperlinkObject -PropertyName 'TextToDisplay')
    }
  } catch {
    $hyperlinkRecord['error'] = 'hyperlink properties unavailable'
  } finally {
    Release-ComObject $hyperlinkObject
    $hyperlinkObject = $null
    Release-ComObject $hyperlinkCollection
    $hyperlinkCollection = $null
  }
  $record['hyperlink'] = $hyperlinkRecord

  return $record
}

function Get-RangeDimension {
  param(
    [Parameter(Mandatory = $true)] [object] $Range,
    [Parameter(Mandatory = $true)] [ValidateSet('Rows', 'Columns')] [string] $Kind
  )

  $dimension = $null
  try {
    if ($Kind -eq 'Rows') {
      $dimension = $Range.Rows
    } else {
      $dimension = $Range.Columns
    }
    return [int] $dimension.Count
  } finally {
    Release-ComObject $dimension
    $dimension = $null
  }
}

function Get-CellInventory {
  param(
    [Parameter(Mandatory = $true)] [object] $Range,
    [int] $MaxCells = 200
  )

  $rowCount = Get-RangeDimension -Range $Range -Kind Rows
  $columnCount = Get-RangeDimension -Range $Range -Kind Columns
  $inventory = @()
  $truncated = $false

  for ($rowIndex = 1; $rowIndex -le $rowCount; $rowIndex++) {
    for ($columnIndex = 1; $columnIndex -le $columnCount; $columnIndex++) {
      if ($inventory.Count -ge $MaxCells) {
        $truncated = $true
        break
      }
      $inventoryCell = $null
      try {
        $inventoryCell = $Range.Item($rowIndex, $columnIndex)
        $inventory += (Get-CellRecord -Cell $inventoryCell)
      } finally {
        Release-ComObject $inventoryCell
        $inventoryCell = $null
      }
    }
    if ($truncated) {
      break
    }
  }

  return [ordered]@{
    cells       = $inventory
    truncated   = $truncated
    rowCount    = $rowCount
    columnCount = $columnCount
  }
}

function Get-RowHeightMap {
  param([Parameter(Mandatory = $true)] [object] $Range)

  $rowCount = Get-RangeDimension -Range $Range -Kind Rows
  $firstRow = [int] (Get-ComPropertyValue -Object $Range -PropertyName 'Row')
  $heights = [ordered]@{}
  $limit = [Math]::Min($rowCount, 40)
  for ($index = 0; $index -lt $limit; $index++) {
    $heightCell = $null
    $heightRow = $null
    try {
      $heightCell = $Range.Item(($index + 1), 1)
      $heightRow = $heightCell.EntireRow
      $rowKey = [string] ($firstRow + $index)
      $heights[$rowKey] = ConvertTo-NullableNumber -Value (Get-ComPropertyValue -Object $heightRow -PropertyName 'RowHeight')
    } finally {
      Release-ComObject $heightRow
      $heightRow = $null
      Release-ComObject $heightCell
      $heightCell = $null
    }
  }
  return $heights
}

function Get-UsedRangeCellSummary {
  param(
    [Parameter(Mandatory = $true)] [object] $Range,
    [Parameter(Mandatory = $true)] [int] $Row,
    [Parameter(Mandatory = $true)] [int] $Column
  )

  $summaryCell = $null
  try {
    $summaryCell = $Range.Item($Row, $Column)
    $summary = [ordered]@{
      ref = [string] ($summaryCell.Address($false, $false))
    }
    Add-CellValueFields -Cell $summaryCell -Record $summary
    return $summary
  } finally {
    Release-ComObject $summaryCell
    $summaryCell = $null
  }
}

function New-ProbeSheet {
  param(
    [Parameter(Mandatory = $true)] [object] $Workbook,
    [Parameter(Mandatory = $true)] [string] $Name
  )

  $sheetCollection = $null
  try {
    $sheetCollection = $Workbook.Worksheets
    $newSheet = $sheetCollection.Add()
    $newSheet.Name = $Name
    return $newSheet
  } finally {
    Release-ComObject $sheetCollection
    $sheetCollection = $null
  }
}

function Set-ProbeCellValue {
  param(
    [Parameter(Mandatory = $true)] [object] $Sheet,
    [Parameter(Mandatory = $true)] [string] $Address,
    [AllowNull()] [object] $Value
  )

  $cellRange = $null
  try {
    $cellRange = $Sheet.Range($Address)
    $cellRange.Value2 = $Value
  } finally {
    Release-ComObject $cellRange
    $cellRange = $null
  }
}

function Set-ProbeCellFormula {
  param(
    [Parameter(Mandatory = $true)] [object] $Sheet,
    [Parameter(Mandatory = $true)] [string] $Address,
    [Parameter(Mandatory = $true)] [string] $Formula
  )

  # Range.Formula is the English, A1-style formula text; assigning it is
  # unambiguous, whereas assigning "=..." to Value2 relies on Excel's parse of
  # a value string.
  $cellRange = $null
  try {
    $cellRange = $Sheet.Range($Address)
    $cellRange.Formula = $Formula
  } finally {
    Release-ComObject $cellRange
    $cellRange = $null
  }
}

function Set-ProbeCellNumberFormat {
  param(
    [Parameter(Mandatory = $true)] [object] $Sheet,
    [Parameter(Mandatory = $true)] [string] $Address,
    [Parameter(Mandatory = $true)] [string] $Format
  )

  $cellRange = $null
  try {
    $cellRange = $Sheet.Range($Address)
    $cellRange.NumberFormat = $Format
  } finally {
    Release-ComObject $cellRange
    $cellRange = $null
  }
}

function Set-ProbeCellFont {
  param(
    [Parameter(Mandatory = $true)] [object] $Sheet,
    [Parameter(Mandatory = $true)] [string] $Address,
    [Parameter(Mandatory = $true)] [hashtable] $Properties
  )

  $cellRange = $null
  $cellFont = $null
  try {
    $cellRange = $Sheet.Range($Address)
    $cellFont = $cellRange.Font
    foreach ($propertyName in $Properties.Keys) {
      $cellFont.$propertyName = $Properties[$propertyName]
    }
  } finally {
    Release-ComObject $cellFont
    $cellFont = $null
    Release-ComObject $cellRange
    $cellRange = $null
  }
}

function Set-ProbeCellFill {
  param(
    [Parameter(Mandatory = $true)] [object] $Sheet,
    [Parameter(Mandatory = $true)] [string] $Address,
    [Parameter(Mandatory = $true)] [hashtable] $Properties
  )

  $cellRange = $null
  $cellInterior = $null
  try {
    $cellRange = $Sheet.Range($Address)
    $cellInterior = $cellRange.Interior
    foreach ($propertyName in $Properties.Keys) {
      $cellInterior.$propertyName = $Properties[$propertyName]
    }
  } finally {
    Release-ComObject $cellInterior
    $cellInterior = $null
    Release-ComObject $cellRange
    $cellRange = $null
  }
}

function Set-ProbeCellBorder {
  param(
    [Parameter(Mandatory = $true)] [object] $Sheet,
    [Parameter(Mandatory = $true)] [string] $Address,
    [Parameter(Mandatory = $true)] [string[]] $Edges,
    [Parameter(Mandatory = $true)] [hashtable] $Properties
  )

  $cellRange = $null
  $cellBorders = $null
  $cellBorder = $null
  try {
    $cellRange = $Sheet.Range($Address)
    $cellBorders = $cellRange.Borders
    foreach ($edgeName in $Edges) {
      $cellBorder = $cellBorders.Item($borderIndexByName[$edgeName])
      foreach ($propertyName in $Properties.Keys) {
        $cellBorder.$propertyName = $Properties[$propertyName]
      }
      Release-ComObject $cellBorder
      $cellBorder = $null
    }
  } finally {
    Release-ComObject $cellBorder
    $cellBorder = $null
    Release-ComObject $cellBorders
    $cellBorders = $null
    Release-ComObject $cellRange
    $cellRange = $null
  }
}

function Set-ProbeCellAlignment {
  param(
    [Parameter(Mandatory = $true)] [object] $Sheet,
    [Parameter(Mandatory = $true)] [string] $Address,
    [Parameter(Mandatory = $true)] [hashtable] $Properties
  )

  # Alignment properties live directly on Range (not on a sub-object), so the
  # properties are set with a dynamic name loop rather than chained access.
  $cellRange = $null
  try {
    $cellRange = $Sheet.Range($Address)
    foreach ($propertyName in $Properties.Keys) {
      $cellRange.$propertyName = $Properties[$propertyName]
    }
  } finally {
    Release-ComObject $cellRange
    $cellRange = $null
  }
}

function Join-ProbeRange {
  param(
    [Parameter(Mandatory = $true)] [object] $Sheet,
    [Parameter(Mandatory = $true)] [string] $Address
  )

  $cellRange = $null
  try {
    $cellRange = $Sheet.Range($Address)
    $cellRange.Merge()
  } finally {
    Release-ComObject $cellRange
    $cellRange = $null
  }
}

function Add-ProbeHyperlink {
  param(
    [Parameter(Mandatory = $true)] [object] $Sheet,
    [Parameter(Mandatory = $true)] [string] $Address,
    [Parameter(Mandatory = $true)] [string] $Url,
    [Parameter(Mandatory = $true)] [string] $TextToDisplay
  )

  $anchor = $null
  $hyperlinkObject = $null
  try {
    $anchor = $Sheet.Range($Address)
    $hyperlinkObject = $Sheet.Hyperlinks.Add($anchor, $Url, '', 'calcplus clipboard spike', $TextToDisplay)
  } finally {
    Release-ComObject $hyperlinkObject
    $hyperlinkObject = $null
    Release-ComObject $anchor
    $anchor = $null
  }
}

function Set-ProbeRowHeight {
  param(
    [Parameter(Mandatory = $true)] [object] $Sheet,
    [Parameter(Mandatory = $true)] [int] $RowIndex,
    [Parameter(Mandatory = $true)] [double] $Height
  )

  $rowCollection = $null
  $rowObject = $null
  try {
    $rowCollection = $Sheet.Rows
    $rowObject = $rowCollection.Item($RowIndex)
    $rowObject.RowHeight = $Height
  } finally {
    Release-ComObject $rowObject
    $rowObject = $null
    Release-ComObject $rowCollection
    $rowCollection = $null
  }
}

function New-ExcelHtmlPayload {
  param(
    [Parameter(Mandatory = $true)] [AllowEmptyString()] [string] $Body,
    [AllowEmptyString()] [string] $Style = ''
  )

  # The calcplus write dialect (spec/clipboard.md section 4): the Excel
  # namespaces plus ProgId make Excel take its native HTML path instead of the
  # reduced generic web-HTML path, which is what keeps mso-* styling.
  $parts = @()
  $parts += '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">'
  $parts += '<head>'
  $parts += '<meta http-equiv="Content-Type" content="text/html; charset=utf-8">'
  $parts += '<meta name="ProgId" content="Excel.Sheet">'
  $parts += '<meta name="Generator" content="Microsoft Excel 15">'
  $parts += '<style><!--'
  $parts += 'td { mso-number-format:General; font-size:11.0pt; font-family:Calibri, sans-serif; vertical-align:bottom; white-space:nowrap; mso-rotate:0; }'
  if ($Style -ne '') {
    $parts += $Style
  }
  $parts += '--></style>'
  $parts += '</head>'
  $parts += '<body>'
  $parts += $Body
  $parts += '</body>'
  $parts += '</html>'
  return ($parts -join "`r`n")
}

function New-GenericHtmlPayload {
  param(
    [Parameter(Mandatory = $true)] [AllowEmptyString()] [string] $Body,
    [AllowEmptyString()] [string] $Style = ''
  )

  # The control payload for the ProgId question: no Excel namespaces, no
  # ProgId, no Generator, and x:num left as an undeclared prefix.
  $parts = @()
  $parts += '<html>'
  $parts += '<head>'
  $parts += '<meta http-equiv="Content-Type" content="text/html; charset=utf-8">'
  $parts += '<style><!--'
  $parts += 'td { font-size:11.0pt; font-family:Calibri, sans-serif; }'
  if ($Style -ne '') {
    $parts += $Style
  }
  $parts += '--></style>'
  $parts += '</head>'
  $parts += '<body>'
  $parts += $Body
  $parts += '</body>'
  $parts += '</html>'
  return ($parts -join "`r`n")
}

function Write-PayloadFile {
  param(
    [Parameter(Mandatory = $true)] [string] $Directory,
    [Parameter(Mandatory = $true)] [string] $FileName,
    [Parameter(Mandatory = $true)] [byte[]] $Bytes
  )

  $path = Join-Path $Directory $FileName
  [System.IO.File]::WriteAllBytes($path, $Bytes)
  return $path
}

function Write-TextPayloadFile {
  param(
    [Parameter(Mandatory = $true)] [string] $Directory,
    [Parameter(Mandatory = $true)] [string] $FileName,
    [AllowNull()] [string] $Text
  )

  $path = Join-Path $Directory $FileName
  if ($null -eq $Text) {
    $Text = ''
  }
  [System.IO.File]::WriteAllText($path, $Text, $utf8NoBom)
  return $path
}

function Test-ProbeSelected {
  param([Parameter(Mandatory = $true)] [string] $ProbeId)

  $patterns = @($script:Probe)
  if ($patterns.Count -eq 0) {
    return $true
  }
  foreach ($pattern in $patterns) {
    if ([string]::IsNullOrWhiteSpace($pattern)) {
      continue
    }
    if ($ProbeId -like $pattern) {
      return $true
    }
  }
  return $false
}

# ---------------------------------------------------------------------
# Excel-source probes (Excel -> clipboard -> this script)
#
# Each probe builds one sheet, copies one range, and the main loop captures the
# clipboard. `question` names the research 06 section 10 question the probe
# answers, or says "context" when it only supplies evidence for a section 3/5
# claim that no open question turns on.
# ---------------------------------------------------------------------
$sourceProbes = @(
  [ordered]@{
    id         = 'S01'
    family     = 'excel-source'
    sheetName  = 'S01 formats'
    title      = 'Clipboard format inventory for a four-cell copy'
    question   = 'Context for research 06 section 3.1 (not a section 10 question)'
    lookFor    = 'The registered-format names Excel publishes and their enumeration order. Expect Biff12, XML Spreadsheet, HTML Format, Rich Text Format, Csv, Hyperlink and the picture formats; only HTML Format and CF_UNICODETEXT are reachable from a browser.'
    decisionRule = 'Context only. The list cannot change the contract because no browser can request the other formats; record it in the report as the baseline for what is thrown away.'
    includeFormatList = $true
    setup = {
      param($sheet, $workbook)
      Set-ProbeCellValue -Sheet $sheet -Address 'A1' -Value 'left'
      Set-ProbeCellValue -Sheet $sheet -Address 'B1' -Value 42
      Set-ProbeCellValue -Sheet $sheet -Address 'A2' -Value 'right'
      Set-ProbeCellValue -Sheet $sheet -Address 'B2' -Value 42.5
    }
    selection = {
      param($sheet, $workbook)
      return $sheet.Range('A1:B2')
    }
  },
  [ordered]@{
    id         = 'S02'
    family     = 'excel-source'
    sheetName  = 'S02 values'
    title      = 'Values and cell typing'
    question   = 'Context for research 06 sections 3.3 and 5 (values row); not a section 10 question'
    lookFor    = 'x:num on the number with the full-precision value in the attribute, x:str on text that looks numeric and on the text constant beginning with =, x:bool for the boolean, x:err for the error, and how the in-cell line break is encoded (br plus mso-data-placement:same-cell).'
    decisionRule = 'If a typed value arrives without its x: attribute, calcplus must stop trusting element text for typing and re-infer from the format code instead. Capture the exact attribute spellings as conformance fixture strings.'
    setup = {
      param($sheet, $workbook)
      Set-ProbeCellValue -Sheet $sheet -Address 'A1' -Value 1234.5678
      Set-ProbeCellValue -Sheet $sheet -Address 'A2' -Value 'hello world'
      Set-ProbeCellValue -Sheet $sheet -Address 'A3' -Value '00123'
      Set-ProbeCellNumberFormat -Sheet $sheet -Address 'A3' -Format '@'
      Set-ProbeCellValue -Sheet $sheet -Address 'A4' -Value $true
      Set-ProbeCellFormula -Sheet $sheet -Address 'A5' -Formula '=1/0'
      Set-ProbeCellValue -Sheet $sheet -Address 'A6' -Value '=1+1'
      Set-ProbeCellValue -Sheet $sheet -Address 'A7' -Value ("line1" + [char] 10 + "line2")
      Set-ProbeCellAlignment -Sheet $sheet -Address 'A7' -Properties @{ WrapText = $true }
    }
    selection = {
      param($sheet, $workbook)
      return $sheet.Range('A1:A7')
    }
  },
  [ordered]@{
    id         = 'S03'
    family     = 'excel-source'
    sheetName  = 'S03 formulas'
    title      = 'Multi-cell formula range, relative, absolute, mixed and cross-sheet'
    question   = 'Section 10 question 1: does Excel 365 put x:fmla in the clipboard HTML for a multi-cell formula range?'
    lookFor    = 'One x:fmla attribute per formula cell, English A1 text with a leading =, the cross-sheet reference kept intact, and the element text holding the computed value. C3 is the exact =B2*2 cell question 2 asks about.'
    decisionRule = 'If any formula cell is missing x:fmla, Excel to calcplus formulas are demoted to not supported via the clipboard (clipboard.md section 3, testing-conformance clipboard row) and the paste dialog says so. If every cell has it, add a conformance fixture built from this payload and keep formulas best-effort.'
    setup = {
      param($sheet, $workbook)
      Set-ProbeCellValue -Sheet $sheet -Address 'B2' -Value 10
      Set-ProbeCellValue -Sheet $sheet -Address 'B3' -Value 20
      Set-ProbeCellValue -Sheet $sheet -Address 'B4' -Value 30
      Set-ProbeCellValue -Sheet $sheet -Address 'B5' -Value 40
      Set-ProbeCellFormula -Sheet $sheet -Address 'C3' -Formula '=B2*2'
      Set-ProbeCellFormula -Sheet $sheet -Address 'C4' -Formula '=$B$2*2'
      Set-ProbeCellFormula -Sheet $sheet -Address 'C5' -Formula '=B$2*2'
      Set-ProbeCellFormula -Sheet $sheet -Address 'C6' -Formula '=$B2*2'
      Set-ProbeCellFormula -Sheet $sheet -Address 'C7' -Formula '=RefData!A1*2'
      Set-ProbeCellFormula -Sheet $sheet -Address 'C8' -Formula "='RefData'!`$A`$1*2"
      Set-ProbeCellFormula -Sheet $sheet -Address 'C9' -Formula '=SUM(B2:B5)'
    }
    selection = {
      param($sheet, $workbook)
      return $sheet.Range('C3:C9')
    }
  },
  [ordered]@{
    id         = 'S04'
    family     = 'excel-source'
    sheetName  = 'S04 formats'
    title      = 'Number formats including currency, date, percent, scientific and a custom code'
    question   = 'Context for research 06 sections 3.3 and 5 (number-format row); not a section 10 question'
    lookFor    = 'The exact mso-number-format strings Excel writes, including the escaped separators in \#\,\#\#0 and the negative section of a custom code, plus which codes come back as named formats (Short Date) instead of explicit codes.'
    decisionRule = 'Emit verbatim the escaping Excel writes (backslash-escaped # , . and -). If a code arrives unescaped, emit unescaped. Store the captured strings as the conformance fixture strings for the HTML writer.'
    setup = {
      param($sheet, $workbook)
      Set-ProbeCellValue -Sheet $sheet -Address 'A1' -Value 1234.5
      Set-ProbeCellNumberFormat -Sheet $sheet -Address 'A1' -Format '#,##0.00'
      Set-ProbeCellValue -Sheet $sheet -Address 'A2' -Value 0.125
      Set-ProbeCellNumberFormat -Sheet $sheet -Address 'A2' -Format '0.00%'
      Set-ProbeCellValue -Sheet $sheet -Address 'A3' -Value (Get-Date -Year 2026 -Month 9 -Day 1)
      Set-ProbeCellNumberFormat -Sheet $sheet -Address 'A3' -Format 'yyyy-mm-dd'
      Set-ProbeCellValue -Sheet $sheet -Address 'A4' -Value (-1234.5678)
      Set-ProbeCellNumberFormat -Sheet $sheet -Address 'A4' -Format '#,##0.000;[Red]-#,##0.000'
      Set-ProbeCellValue -Sheet $sheet -Address 'A5' -Value 12345.678
      Set-ProbeCellNumberFormat -Sheet $sheet -Address 'A5' -Format '0.0000E+00'
      Set-ProbeCellValue -Sheet $sheet -Address 'A6' -Value 1234.5
      Set-ProbeCellNumberFormat -Sheet $sheet -Address 'A6' -Format '$#,##0.00'
      Set-ProbeCellValue -Sheet $sheet -Address 'A7' -Value '00123'
      Set-ProbeCellNumberFormat -Sheet $sheet -Address 'A7' -Format '@'
    }
    selection = {
      param($sheet, $workbook)
      return $sheet.Range('A1:A7')
    }
  },
  [ordered]@{
    id         = 'S05'
    family     = 'excel-source'
    sheetName  = 'S05 styles'
    title      = 'Fonts, fills, borders and alignment'
    question   = 'Context for research 06 sections 3.4 and 5 (font, fill, border, alignment rows); not a section 10 question'
    lookFor    = 'CSS properties per cell: font-family/font-size/font-weight/font-style/text-decoration/color, background plus mso-pattern for the fills, per-edge border properties (and whether diagonal borders use mso-border-diagonal-up), text-align/vertical-align/white-space/mso-rotate/indent.'
    decisionRule = 'The captured CSS is the write contract. Any property Excel writes that the list in clipboard.md section 4 does not cover is either added with its observed spelling or dropped from the promise. If fills need both background and mso-pattern, calcplus must emit both.'
    setup = {
      param($sheet, $workbook)
      Set-ProbeCellValue -Sheet $sheet -Address 'A1' -Value 'bold'
      Set-ProbeCellFont -Sheet $sheet -Address 'A1' -Properties @{ Bold = $true }
      Set-ProbeCellValue -Sheet $sheet -Address 'A2' -Value 'italic'
      Set-ProbeCellFont -Sheet $sheet -Address 'A2' -Properties @{ Italic = $true }
      Set-ProbeCellValue -Sheet $sheet -Address 'A3' -Value 'size 14'
      Set-ProbeCellFont -Sheet $sheet -Address 'A3' -Properties @{ Size = 14 }
      Set-ProbeCellValue -Sheet $sheet -Address 'A4' -Value 'arial'
      Set-ProbeCellFont -Sheet $sheet -Address 'A4' -Properties @{ Name = 'Arial' }
      Set-ProbeCellValue -Sheet $sheet -Address 'A5' -Value 'red'
      Set-ProbeCellFont -Sheet $sheet -Address 'A5' -Properties @{ ColorIndex = 3 }
      Set-ProbeCellValue -Sheet $sheet -Address 'A6' -Value 'underline'
      Set-ProbeCellFont -Sheet $sheet -Address 'A6' -Properties @{ Underline = 2 }
      Set-ProbeCellValue -Sheet $sheet -Address 'B1' -Value 'fill solid'
      Set-ProbeCellFill -Sheet $sheet -Address 'B1' -Properties @{ ColorIndex = 6; Pattern = 1 }
      Set-ProbeCellValue -Sheet $sheet -Address 'B2' -Value 'fill pattern'
      Set-ProbeCellFill -Sheet $sheet -Address 'B2' -Properties @{ ColorIndex = 6; Pattern = -4125; PatternColorIndex = 1 }
      Set-ProbeCellValue -Sheet $sheet -Address 'C1' -Value 'border box'
      Set-ProbeCellBorder -Sheet $sheet -Address 'C1' -Edges @('EdgeLeft', 'EdgeTop', 'EdgeBottom', 'EdgeRight') -Properties @{ LineStyle = 1; Weight = 4; ColorIndex = 1 }
      Set-ProbeCellValue -Sheet $sheet -Address 'C2' -Value 'border dash'
      Set-ProbeCellBorder -Sheet $sheet -Address 'C2' -Edges @('EdgeLeft', 'EdgeTop', 'EdgeBottom', 'EdgeRight') -Properties @{ LineStyle = -4115; Weight = 2; ColorIndex = 5 }
      Set-ProbeCellValue -Sheet $sheet -Address 'C3' -Value 'align center'
      Set-ProbeCellAlignment -Sheet $sheet -Address 'C3' -Properties @{ HorizontalAlignment = -4108 }
      Set-ProbeCellValue -Sheet $sheet -Address 'C4' -Value 'wrap text in a cell'
      Set-ProbeCellAlignment -Sheet $sheet -Address 'C4' -Properties @{ WrapText = $true }
      Set-ProbeCellValue -Sheet $sheet -Address 'C5' -Value 'indent 2'
      Set-ProbeCellAlignment -Sheet $sheet -Address 'C5' -Properties @{ IndentLevel = 2 }
      Set-ProbeCellValue -Sheet $sheet -Address 'C6' -Value 'rotate 45'
      Set-ProbeCellAlignment -Sheet $sheet -Address 'C6' -Properties @{ Orientation = 45 }
      Set-ProbeCellValue -Sheet $sheet -Address 'C7' -Value 'valign top'
      Set-ProbeCellAlignment -Sheet $sheet -Address 'C7' -Properties @{ VerticalAlignment = -4160 }
    }
    selection = {
      param($sheet, $workbook)
      return $sheet.Range('A1:C7')
    }
  },
  [ordered]@{
    id         = 'S06'
    family     = 'excel-source'
    sheetName  = 'S06 merges'
    title      = 'Merged ranges (2x3 block, 1x4 column merge, 2x2 block) with row heights'
    question   = 'Context for research 06 section 5 (merged cells and row heights); feeds the target-side section 10 question 4'
    lookFor    = 'colspan and rowspan on the top-left cell of each merge, the covered cells omitted, the merged cell value present once, and height plus mso-height-source on the explicitly sized rows.'
    decisionRule = 'If Excel writes colspan/rowspan, the calcplus parser maps spans to merges exactly as Excel does and the writer emits the same shape. If spans are missing, merges do not travel through the clipboard and clipboard.md section 3 loses the merge promise.'
    setup = {
      param($sheet, $workbook)
      Join-ProbeRange -Sheet $sheet -Address 'A1:C2'
      Set-ProbeCellValue -Sheet $sheet -Address 'A1' -Value 'merged 2 rows x 3 columns'
      Join-ProbeRange -Sheet $sheet -Address 'A4:A7'
      Set-ProbeCellValue -Sheet $sheet -Address 'A4' -Value 'vertical merge 4 rows'
      Join-ProbeRange -Sheet $sheet -Address 'B4:C5'
      Set-ProbeCellValue -Sheet $sheet -Address 'B4' -Value 'block 2x2'
      Set-ProbeCellValue -Sheet $sheet -Address 'A9' -Value 'after the merges'
      Set-ProbeRowHeight -Sheet $sheet -RowIndex 1 -Height 30
      Set-ProbeRowHeight -Sheet $sheet -RowIndex 2 -Height 30
      Set-ProbeRowHeight -Sheet $sheet -RowIndex 6 -Height 42
    }
    selection = {
      param($sheet, $workbook)
      return $sheet.Range('A1:C9')
    }
  },
  [ordered]@{
    id         = 'S07'
    family     = 'excel-source'
    sheetName  = 'S07 hyperlink'
    title      = 'Cells carrying a hyperlink'
    question   = 'Section 10 question 3, source half: does Excel write an a href for a hyperlink cell?'
    lookFor    = 'An a href around the cell text, the URL as attribute value, and whether the display text is the link text or the URL. Note any change to the cell typing attributes.'
    decisionRule = 'If the source payload has no anchor at all, Excel to calcplus can never recover a hyperlink from HTML and the promise stays at best-effort. The target probe T06 decides the other direction.'
    setup = {
      param($sheet, $workbook)
      Add-ProbeHyperlink -Sheet $sheet -Address 'A1' -Url 'https://example.com/calcplus-one' -TextToDisplay 'calcplus link'
      Add-ProbeHyperlink -Sheet $sheet -Address 'A2' -Url 'https://example.com/calcplus-two' -TextToDisplay 'https://example.com/calcplus-two'
    }
    selection = {
      param($sheet, $workbook)
      return $sheet.Range('A1:A2')
    }
  },
  [ordered]@{
    id         = 'S08'
    family     = 'excel-source'
    sheetName  = 'S08 single text'
    title      = 'Single-cell copy of a text cell'
    question   = 'Section 10 question 9: does a single-cell copy degrade to a bare text run with no table?'
    lookFor    = 'Whether the CF_HTML payload contains table and td at all, whether the style block and ProgId survive, and whether the fragment is only the value text. An absent or empty HTML flavour is an observation to record, not a failure.'
    decisionRule = 'If there is no table, the parser must take the no-table path (one cell from the body text) before clipboard paste ships, and a fixture from this payload is committed to prove it.'
    setup = {
      param($sheet, $workbook)
      Set-ProbeCellValue -Sheet $sheet -Address 'C3' -Value 'single cell text'
    }
    selection = {
      param($sheet, $workbook)
      return $sheet.Range('C3')
    }
  },
  [ordered]@{
    id         = 'S09'
    family     = 'excel-source'
    sheetName  = 'S09 single number'
    title      = 'Single-cell copy of a formatted number'
    question   = 'Section 10 question 9: single-cell copy shape for a numeric cell'
    lookFor    = 'Whether x:num and mso-number-format survive when the copy is a single cell, and whether the payload still has a table.'
    decisionRule = 'If mso-number-format disappears on a single-cell copy, single-cell pastes fall back to plain typing and the UI must not claim format fidelity for a one-cell copy.'
    setup = {
      param($sheet, $workbook)
      Set-ProbeCellValue -Sheet $sheet -Address 'C3' -Value 1234.5
      Set-ProbeCellNumberFormat -Sheet $sheet -Address 'C3' -Format '#,##0.00'
    }
    selection = {
      param($sheet, $workbook)
      return $sheet.Range('C3')
    }
  },
  [ordered]@{
    id         = 'S10'
    family     = 'excel-source'
    sheetName  = 'S10 single formula'
    title      = 'Single-cell copy of a formula cell'
    question   = 'Section 10 questions 9 and 1: single-cell copy shape for a formula cell'
    lookFor    = 'Whether x:fmla appears for a one-cell copy, which decides whether a single formula paste can ever carry the formula rather than its value.'
    decisionRule = 'If x:fmla is absent for a single-cell copy but present for a range, calcplus must warn on single-formula copies or fall back to paste-all only. Record both spellings in the conformance fixtures.'
    setup = {
      param($sheet, $workbook)
      Set-ProbeCellValue -Sheet $sheet -Address 'B2' -Value 10
      Set-ProbeCellFormula -Sheet $sheet -Address 'C3' -Formula '=B2*2'
    }
    selection = {
      param($sheet, $workbook)
      return $sheet.Range('C3')
    }
  },
  [ordered]@{
    id         = 'S11'
    family     = 'excel-source'
    sheetName  = 'S11 large'
    title      = 'Large range: 10,000 rows x 5 columns (50,000 cells)'
    question   = 'Section 10 question 10: large-range payload size, copy time and paste time'
    lookFor    = 'CF_HTML byte length, CF_UNICODETEXT byte length, the Copy call time, the clipboard read time (Excel renders CF_HTML on first read, so the read is where most of the cost appears) and the time for Excel to paste its own payload into a clean sheet.'
    decisionRule = 'Take the measured HTML size as the payload budget input. If 50,000 cells exceed the Windows clipboard history limit (about 4 MB) the clipboard writer caps the payload at the measured size and cell count and warns, instead of truncating silently.'
    measurePaste = $true
    setup = {
      param($sheet, $workbook)
      # One Value2 assignment per 1,000-row chunk: two orders of magnitude
      # faster than 50,000 single-cell COM writes.
      $chunkRows = 1000
      for ($chunkStart = 1; $chunkStart -le 10000; $chunkStart += $chunkRows) {
        $values = New-Object 'object[,]' $chunkRows, 5
        for ($rowOffset = 0; $rowOffset -lt $chunkRows; $rowOffset++) {
          for ($columnOffset = 0; $columnOffset -lt 5; $columnOffset++) {
            $values[$rowOffset, $columnOffset] = (($chunkStart + $rowOffset - 1) * 5) + $columnOffset + 1
          }
        }
        $chunkRange = $null
        try {
          $chunkRange = $sheet.Range(('A{0}:E{1}' -f $chunkStart, ($chunkStart + $chunkRows - 1)))
          $chunkRange.Value2 = $values
        } finally {
          Release-ComObject $chunkRange
          $chunkRange = $null
        }
      }
    }
    selection = {
      param($sheet, $workbook)
      return $sheet.Range('A1:E10000')
    }
  }
)

# ---------------------------------------------------------------------
# Excel-target probes (this script -> clipboard -> Excel)
#
# Each probe publishes one payload and lets Excel paste it into a clean sheet.
# HTML-only payloads deliberately leave CF_UNICODETEXT off the clipboard so the
# observed result can only have come from the HTML path; T11 publishes both to
# record which flavour wins.
# ---------------------------------------------------------------------
$targetProbes = @(
  [ordered]@{
    id         = 'T01'
    family     = 'excel-target'
    sheetName  = 'T01 types'
    title      = 'Typed cells: x:num, x:str, x:bool, x:err, plain text and a date-looking number'
    question   = 'Context for research 06 sections 3.4 and 6.1 rule 3; not a section 10 question'
    lookFor    = 'A1 number 1234.5, B1 text 00123, C1 boolean TRUE, D1 error #DIV/0!, E1 text, F1 the number 46266 rather than a date. Element text is deliberately the formatted text, so the typed value must come from the attribute.'
    decisionRule = 'If any cell lands with the wrong kind, calcplus must add or change a type attribute for that case. If 46266 lands as a number, the no-guessing claim in clipboard.md section 4 holds for date-looking numbers.'
    build = {
      $body = @'
<table border=0 cellpadding=0 cellspacing=0>
<tr height=20>
<td x:num="1234.5">1,234.50</td>
<td x:str>00123</td>
<td x:bool="TRUE">TRUE</td>
<td x:err="#DIV/0!">#DIV/0!</td>
<td>plain text</td>
<td x:num="46266">2026-09-01</td>
</tr>
</table>
'@
      return [ordered]@{ html = (New-ExcelHtmlPayload -Body $body); text = $null }
    }
  },
  [ordered]@{
    id         = 'T02'
    family     = 'excel-target'
    sheetName  = 'T02 formats'
    title      = 'Number formats via mso-number-format, escaped and named'
    question   = 'Context for research 06 sections 3.4 and 6.3 (number-format row); not a section 10 question'
    lookFor    = 'The NumberFormat read back per cell against the code sent. Watch the escaped separators in \#\,\#\#0\.00, whether Short Date resolves to a real code (for example m/d/yyyy), and whether the negative section survives.'
    decisionRule = 'Emit codes in exactly the spelling this probe shows Excel accepts. If escaped separators are normalised on paste, emit them normalised (and say so in clipboard.md section 4); if a named format like Short Date is rejected, never emit named formats.'
    build = {
      $style = @'
.xl1 { mso-number-format:"\#\,\#\#0\.00"; }
.xl2 { mso-number-format:"0\.00%"; }
.xl3 { mso-number-format:"yyyy\-mm\-dd"; }
.xl4 { mso-number-format:"Short Date"; }
.xl5 { mso-number-format:"0\.0000E+00"; }
.xl6 { mso-number-format:"\#\,\#\#0\.000;[Red]\-\#\,\#\#0\.000"; }
.xl7 { mso-number-format:"@"; }
'@
      $body = @'
<table border=0 cellpadding=0 cellspacing=0>
<tr height=20>
<td class=xl1 x:num="1234.5">1,234.50</td>
<td class=xl2 x:num="0.125">12.50%</td>
<td class=xl3 x:num="46266">2026-09-01</td>
<td class=xl4 x:num="46266">9/1/2026</td>
<td class=xl5 x:num="12345.678">1.2346E+04</td>
<td class=xl6 x:num="-1234.5678">-1,234.568</td>
<td class=xl7 x:str>00123</td>
</tr>
</table>
'@
      return [ordered]@{ html = (New-ExcelHtmlPayload -Body $body -Style $style); text = $null }
    }
  },
  [ordered]@{
    id         = 'T03'
    family     = 'excel-target'
    sheetName  = 'T03 styles'
    title      = 'Font, fill, border and alignment classes'
    question   = 'Context for research 06 sections 3.4 and 6.3 (font, fill, border, alignment rows); not a section 10 question'
    lookFor    = 'Font weight, italic, size and colour; fill colorIndex with the pattern; per-edge border line style, weight and colour; horizontal and vertical alignment, wrap and rotation. Compare the inventory against the CSS that was sent.'
    decisionRule = 'Every property that does not survive is removed from the calcplus to Excel promise in clipboard.md section 4, or its CSS spelling is changed to one Excel accepts. Properties that do survive become the fixture expectations.'
    build = {
      $style = @'
.xl1 { font-weight:700; }
.xl2 { font-style:italic; color:#C00000; }
.xl3 { font-size:14.0pt; }
.xl4 { font-family:"Courier New", monospace; }
.xl5 { background:#FFF2CC; mso-pattern:black none; }
.xl6 { background:#FFF2CC; mso-pattern:gray50; }
.xl7 { border:2.0pt solid #000000; }
.xl8 { border-style:dashed; border-width:1.0pt; border-color:#0000FF; }
.xl9 { text-align:center; vertical-align:top; }
.xl10 { white-space:normal; }
.xl11 { mso-rotate:45; }
'@
      $body = @'
<table border=0 cellpadding=0 cellspacing=0>
<tr height=20>
<td class=xl1 x:str>bold</td>
<td class=xl2 x:str>italic red</td>
<td class=xl3 x:str>size 14</td>
<td class=xl4 x:str>courier</td>
</tr>
<tr height=20>
<td class=xl5 x:str>fill solid</td>
<td class=xl6 x:str>fill gray50</td>
<td class=xl7 x:str>border thick</td>
<td class=xl8 x:str>border dashed</td>
</tr>
<tr height=20>
<td class=xl9 x:str>center top</td>
<td class=xl10 x:str>wrap</td>
<td class=xl11 x:str>rotate 45</td>
<td x:str>plain</td>
</tr>
</table>
'@
      return [ordered]@{ html = (New-ExcelHtmlPayload -Body $body -Style $style); text = $null }
    }
  },
  [ordered]@{
    id         = 'T04'
    family     = 'excel-target'
    sheetName  = 'T04 merges'
    title      = 'colspan and rowspan merges'
    question   = 'Context for research 06 sections 3.4 and 6.3 (merged cells), and the rowspan half of section 10 question 4'
    lookFor    = 'MergeCells true and the MergeArea address for A1:C1 (colspan 3), A2:A3 (rowspan 2) and A4:B5 (2x2). Unmerged neighbours must stay unmerged. Used range is A1:C5.'
    decisionRule = 'If every span becomes the right merge, colspan/rowspan stays guaranteed. If a colspan works but a rowspan does not, merges are only promised for single-row merges and multi-row merges are reported as unsupported.'
    build = {
      $body = @'
<table border=0 cellpadding=0 cellspacing=0>
<tr height=20><td colspan=3 x:str>wide 3 columns</td></tr>
<tr height=20><td rowspan=2 x:str>tall 2 rows</td><td x:num="1">1</td><td x:num="2">2</td></tr>
<tr height=20><td x:num="3">3</td><td x:num="4">4</td></tr>
<tr height=20><td colspan=2 rowspan=2 x:str>block 2x2</td><td x:num="5">5</td></tr>
<tr height=20><td x:num="6">6</td></tr>
</table>
'@
      return [ordered]@{ html = (New-ExcelHtmlPayload -Body $body); text = $null }
    }
  },
  [ordered]@{
    id         = 'T05'
    family     = 'excel-target'
    sheetName  = 'T05 heights'
    title      = 'Row heights in pixels, in points, and default'
    question   = 'Section 10 question 4: does Excel honour height and rowspan on paste, or auto-fit?'
    lookFor    = 'The rowHeights map for rows 1 to 4. Row 1 is tr height=40 (pixels in HTML), row 2 is style height:60pt plus mso-height-source:userset, row 3 is tr height=20, row 4 carries no height attribute. Default for 11pt Calibri is about 15 points (20 pixels).'
    decisionRule = 'If pixel heights land, calcplus converts points to pixels (1pt = 96/72px). If only the point form lands, emit points. If Excel auto-fits everything, row heights are dropped from the calcplus to Excel promise and the paste UI stops offering them.'
    build = {
      $body = @'
<table border=0 cellpadding=0 cellspacing=0>
<tr height=40><td x:str>height 40 pixels</td></tr>
<tr><td style="mso-height-source:userset;height:60pt" x:str>height 60 points</td></tr>
<tr height=20><td x:str>height 20 pixels</td></tr>
<tr><td x:str>no height attribute</td></tr>
</table>
'@
      return [ordered]@{ html = (New-ExcelHtmlPayload -Body $body); text = $null }
    }
  },
  [ordered]@{
    id         = 'T06'
    family     = 'excel-target'
    sheetName  = 'T06 hyperlink'
    title      = 'a href hyperlink import'
    question   = 'Section 10 question 3: does Excel import an a href as a hyperlink on HTML paste, or only as text?'
    lookFor    = 'Per cell: hyperlink.count and hyperlink.address. A1 and C1 should carry the URL, B1 the mailto address. Also record whether the visible text is the link text or the URL.'
    decisionRule = 'If Hyperlinks.Count is 1 with the right address, a href stays in the write contract as best-effort and a conformance fixture pins it. If the cell only holds text, the a href promise is removed from clipboard.md section 4 and the cell text (not the URL) is what calcplus emits.'
    build = {
      $body = @'
<table border=0 cellpadding=0 cellspacing=0>
<tr height=20>
<td><a href="https://example.com/calcplus-one">calcplus link</a></td>
<td><a href="mailto:someone@example.com">mail us</a></td>
<td><a href="https://example.com/calcplus-two">https://example.com/calcplus-two</a></td>
</tr>
</table>
'@
      return [ordered]@{ html = (New-ExcelHtmlPayload -Body $body); text = $null }
    }
  },
  [ordered]@{
    id         = 'T07'
    family     = 'excel-target'
    sheetName  = 'T07 anchoring'
    title      = 'x:fmla reference anchoring (payload pasted at F10)'
    question   = 'Section 10 question 2: how does Excel anchor x:fmla references on paste?'
    lookFor    = 'The formula read back in F10 to F14. Verbatim keeps =B2*2; a spreadsheet-style rebase from A1 to F10 shifts relative references by 5 columns and 9 rows, giving =G11*2 and =SUM(G11:G14). Absolute and mixed references show which part moves.'
    decisionRule = 'If verbatim, calcplus must write the formula already anchored for the destination cell (rebase relative references itself) or default to paste as values whenever a formula is present. If rebased, emit the source text verbatim and document that pasting at an offset shifts relative references.'
    expectations = @(
      'verbatim: F10 =B2*2, F11 =$B$2*2, F12 =B$2*2, F13 =$B2*2, F14 =SUM(B2:B5)',
      'rebased from A1 to F10 (+5 columns, +9 rows): F10 =G11*2, F11 =$B$2*2, F12 =G$2*2, F13 =$B11*2, F14 =SUM(G11:G14)'
    )
    destination = 'F10'
    build = {
      $body = @'
<table border=0 cellpadding=0 cellspacing=0>
<tr height=20><td x:fmla="=B2*2" x:num>20</td></tr>
<tr height=20><td x:fmla="=$B$2*2" x:num>20</td></tr>
<tr height=20><td x:fmla="=B$2*2" x:num>20</td></tr>
<tr height=20><td x:fmla="=$B2*2" x:num>20</td></tr>
<tr height=20><td x:fmla="=SUM(B2:B5)" x:num>20</td></tr>
</table>
'@
      return [ordered]@{ html = (New-ExcelHtmlPayload -Body $body); text = $null }
    }
  },
  [ordered]@{
    id         = 'T08'
    family     = 'excel-target'
    sheetName  = 'T08 tsv quotes'
    title      = 'TSV fallback with CSV-style quoting, embedded tab, embedded newline and doubled quotes'
    question   = 'Section 10 question 5: does Excel''s TSV parser honour quoting for embedded tabs and newlines on paste?'
    lookFor    = 'The used range shape. Quoting honoured gives A1:D2 with B1 = quoted "inner" quotes, C1 a two-line cell and D1 a cell containing a tab. Quoting ignored splits D1 and row 1 into extra columns or rows; count the used cells and read each one.'
    decisionRule = 'If quoting is honoured, keep the CSV-style quoting in clipboard.md section 4 and in the TSV writer. If not, the TSV writer must refuse embedded tabs and newlines (replace with a space or a special marker) and rely on the HTML flavour for multi-line cells.'
    build = {
      $tab = [string] [char] 9
      $lineFeed = [string] [char] 10
      $rowOne = 'plain' + $tab + '"quoted ""inner"" quotes"' + $tab + '"line1' + $lineFeed + 'line2"' + $tab + '"tab' + $tab + 'inside"'
      $rowTwo = '2' + $tab + '3' + $tab + '4' + $tab + '5'
      $text = $rowOne + "`r`n" + $rowTwo
      return [ordered]@{ html = $null; text = $text }
    }
  },
  [ordered]@{
    id         = 'T09'
    family     = 'excel-target'
    sheetName  = 'T09 line endings'
    title      = 'TSV fallback with LF-only row separators'
    question   = 'Section 10 question 5, line-ending half: does Excel split rows on a bare LF?'
    lookFor    = 'Whether the used range is A1:B2 (LF splits rows) or A1:B1 with one text cell (only CRLF splits rows).'
    decisionRule = 'CRLF is what calcplus writes. If Excel needs CRLF and a lone LF does not split, keep CRLF as a hard rule in the TSV writer and add a unit test that asserts no bare LF reaches the clipboard.'
    build = {
      $tab = [string] [char] 9
      $lineFeed = [string] [char] 10
      $text = 'a' + $tab + 'b' + $lineFeed + 'c' + $tab + 'd'
      return [ordered]@{ html = $null; text = $text }
    }
  },
  [ordered]@{
    id         = 'T10'
    family     = 'excel-target'
    sheetName  = 'T10 no progid'
    title      = 'Control: the same typed payload without xmlns:x, ProgId or Generator'
    question   = 'Context for research 06 section 6.1 rule 2; not a section 10 question'
    lookFor    = 'Whether the number format and the fill still arrive when the Excel namespaces and ProgId are missing and x:num is an undeclared prefix. Compare with T02 and T03, which carry them.'
    decisionRule = 'If the control loses formatting, ProgId plus the x namespace is load-bearing and the writer must emit both, with a conformance test that asserts their presence. If the control keeps everything, the meta tags stay as defence in depth for older builds.'
    build = {
      $style = @'
.xl1 { mso-number-format:"\#\,\#\#0\.00"; background:#FFF2CC; mso-pattern:black none; }
'@
      $body = @'
<table border=0 cellpadding=0 cellspacing=0>
<tr height=20>
<td class=xl1 x:num="1234.5">1,234.50</td>
<td class=xl1 x:num="0.5">50.00%</td>
</tr>
</table>
'@
      return [ordered]@{ html = (New-GenericHtmlPayload -Body $body -Style $style); text = $null }
    }
  },
  [ordered]@{
    id         = 'T11'
    family     = 'excel-target'
    sheetName  = 'T11 both flavours'
    title      = 'Both flavours on the clipboard: which does Excel prefer?'
    question   = 'Context for the paste-priority gap in research 06 section 3.2; not a section 10 question'
    lookFor    = 'A1: 111 (HTML won) or 999 (text won). The number format shows whether the HTML path ran at all. This is the real calcplus clipboard shape, HTML plus TSV.'
    decisionRule = 'If text wins, Excel never sees the formats or merges when both flavours are present, so the write path needs an HTML-only mode or a documented Paste Special hint for Excel targets.'
    build = {
      $style = @'
.xl1 { mso-number-format:"\#\,\#\#0\.00"; }
'@
      $body = @'
<table border=0 cellpadding=0 cellspacing=0>
<tr height=20>
<td class=xl1 x:num="111">111.00</td>
<td x:str>html</td>
</tr>
</table>
'@
      $tab = [string] [char] 9
      $text = '999' + $tab + 'text'
      return [ordered]@{ html = (New-ExcelHtmlPayload -Body $body -Style $style); text = $text }
    }
  },
  [ordered]@{
    id         = 'T12'
    family     = 'excel-target'
    sheetName  = 'T12 large'
    title      = 'Large calcplus payload: 10,000 rows x 5 columns (50,000 cells)'
    question   = 'Section 10 question 10: large-range payload size, clipboard write time and Excel paste time'
    lookFor    = 'Generated HTML size, clipboard write time, Excel paste time, and the used cell count. Payload generation time is PowerShell building a string, not the browser, so only the clipboard and paste numbers are comparable with the browser measurement.'
    decisionRule = 'Use this pair (S11 and T12) as the two ends of the payload budget. Excel paste time is the floor a browser paste can never beat; if calcplus needs more than the S11 measured size for the same cell count, the writer is emitting too much per cell.'
    inventory = 'none'
    build = {
      $builder = New-Object System.Text.StringBuilder
      [void] $builder.Append('<table border=0 cellpadding=0 cellspacing=0>')
      for ($row = 1; $row -le 10000; $row++) {
        [void] $builder.Append('<tr height=20>')
        for ($column = 1; $column -le 5; $column++) {
          $value = (($row - 1) * 5) + $column
          [void] $builder.Append('<td class=xl1 x:num="' + $value.ToString([System.Globalization.CultureInfo]::InvariantCulture) + '">' + $value.ToString([System.Globalization.CultureInfo]::InvariantCulture) + '</td>')
        }
        [void] $builder.Append('</tr>')
      }
      [void] $builder.Append('</table>')
      $style = '.xl1 { mso-number-format:"\#\,\#\#0"; }'
      return [ordered]@{ html = (New-ExcelHtmlPayload -Body $builder.ToString() -Style $style); text = $null }
    }
  }
)

# ---------------------------------------------------------------------
# Evidence directory and probe selection
# ---------------------------------------------------------------------
$evidenceDirectory = $EvidencePath
if (-not [System.IO.Path]::IsPathRooted($evidenceDirectory)) {
  $evidenceDirectory = Join-Path (Get-Location).Path $evidenceDirectory
}
if (-not (Test-Path -LiteralPath $evidenceDirectory)) {
  New-Item -ItemType Directory -Path $evidenceDirectory -Force | Out-Null
}

$selectedSourceProbes = @()
foreach ($probeDefinition in $sourceProbes) {
  if (Test-ProbeSelected -ProbeId ([string] $probeDefinition.id)) {
    $selectedSourceProbes += $probeDefinition
  }
}
$selectedTargetProbes = @()
foreach ($probeDefinition in $targetProbes) {
  if (Test-ProbeSelected -ProbeId ([string] $probeDefinition.id)) {
    $selectedTargetProbes += $probeDefinition
  }
}
if (($selectedSourceProbes.Count + $selectedTargetProbes.Count) -eq 0) {
  $availableIds = @()
  foreach ($probeDefinition in $sourceProbes) { $availableIds += [string] $probeDefinition.id }
  foreach ($probeDefinition in $targetProbes) { $availableIds += [string] $probeDefinition.id }
  throw ("No probe matched -Probe '$($Probe -join ', ')'.
Available probe ids: $($availableIds -join ', ')")
}

try {
  try {
    $excel = New-Object -ComObject Excel.Application
  } catch {
    throw "Excel 365 COM is unavailable on this PC. Install and activate desktop Excel 365 before running the clipboard spike. $($_.Exception.Message)"
  }

  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $excel.ScreenUpdating = $false
  # msoAutomationSecurityForceDisable = 3. Macro projects are never executed.
  $excel.AutomationSecurity = 3

  # The registered "HTML Format" id is what turns text/html into CF_HTML on
  # Windows. A zero id means this Windows build cannot carry the format at all.
  $htmlFormatId = [NativeClipboard]::RegisterClipboardFormat('HTML Format')
  if ($htmlFormatId -eq 0) {
    throw 'RegisterClipboardFormat("HTML Format") returned 0, so this Windows build cannot carry CF_HTML. The clipboard spike cannot run here.'
  }

  # Keep each COM collection in its own variable and use indexed loops, as
  # capture-goldens.ps1 does: a chained access such as $workbook.Application
  # leaks an extra Application RCW.
  $workbooks = $excel.Workbooks
  try {
    $workbook = $workbooks.Add()
  } catch {
    throw (New-ComFailureMessage -Context 'Creating the probe workbook' -Exception $_.Exception)
  }
  # Pin the 1900 date system so the date serial in the fixtures (46266 is
  # 2026-09-01 in the 1900 system) is reproducible regardless of the PC's
  # "Use 1904 date system" option.
  try {
    $workbook.Date1904 = $false
  } catch {
    throw (New-ComFailureMessage -Context 'Setting the probe workbook date system' -Exception $_.Exception)
  }

  $refDataSheet = $null
  try {
    $refDataSheet = New-ProbeSheet -Workbook $workbook -Name 'RefData'
    Set-ProbeCellValue -Sheet $refDataSheet -Address 'A1' -Value 5
    Set-ProbeCellValue -Sheet $refDataSheet -Address 'A2' -Value 7
  } catch {
    throw (New-ComFailureMessage -Context 'Building the cross-sheet reference sheet' -Exception $_.Exception)
  } finally {
    Release-ComObject $refDataSheet
    $refDataSheet = $null
  }

  $manifest = [ordered]@{
    spike           = 'M5 risk spike: Excel clipboard interop'
    script          = 'scripts/oracle/clipboard-spike.ps1'
    research        = '.scratch/calcplus/research/06-clipboard-interop.md section 10'
    contract        = '.scratch/calcplus/spec/clipboard.md'
    excelVersion    = [string] $excel.Version
    excelBuild      = [string] $excel.Build
    os              = [System.Environment]::OSVersion.VersionString
    locale          = [System.Globalization.CultureInfo]::CurrentCulture.Name
    date1904        = [bool] $workbook.Date1904
    generatedAt     = (Get-Date).ToUniversalTime().ToString('o')
    probeFilter     = @($Probe)
    probesRequested = ($selectedSourceProbes.Count + $selectedTargetProbes.Count)
    clipboardFormatIds = [ordered]@{
      htmlFormat      = [int] $htmlFormatId
      unicodeText     = 13
      text            = 1
    }
  }

  # ---------------- Excel-source probes ----------------
  foreach ($probeDefinition in $selectedSourceProbes) {
    $probeId = [string] $probeDefinition.id
    Write-Output "Probe $probeId - $($probeDefinition.title)"

    $worksheet = $null
    $range = $null
    $payloadFiles = @()

    try {
      try {
        $worksheet = New-ProbeSheet -Workbook $workbook -Name ([string] $probeDefinition.sheetName)
      } catch {
        throw (New-ComFailureMessage -Context "Creating the probe sheet for $probeId" -Exception $_.Exception)
      }

      try {
        & $probeDefinition.setup $worksheet $workbook
      } catch {
        throw (New-ComFailureMessage -Context "Building the probe workbook for $probeId" -Exception $_.Exception)
      }

      try {
        $range = & $probeDefinition.selection $worksheet $workbook
      } catch {
        throw (New-ComFailureMessage -Context "Selecting the copy range for $probeId" -Exception $_.Exception)
      }

      $copyWatch = [System.Diagnostics.Stopwatch]::StartNew()
      $null = Invoke-ComActionWithRetry -Context "Range.Copy for $probeId" -Action { $range.Copy() }
      $copyWatch.Stop()

      # Excel renders the clipboard formats on demand: Range.Copy returns when
      # the data object is registered, and the first GetClipboardData inside
      # this snapshot triggers the actual render. Both timings are recorded
      # because that split is invisible to the browser that reads text/html.
      $readWatch = [System.Diagnostics.Stopwatch]::StartNew()
      $includeFormatList = [bool] $probeDefinition.includeFormatList
      $snapshot = Get-ClipboardSnapshot -IncludeFormatList:$includeFormatList -HtmlFormat $htmlFormatId
      $readWatch.Stop()

      $htmlDecoded = ConvertFrom-HtmlPayloadBytes -Bytes ($snapshot['html']['bytes'])
      $cfHtmlHeader = Get-CfHtmlHeader -Bytes ($snapshot['html']['bytes'])
      $fragment = $cfHtmlHeader['fragment']
      $fragmentSource = 'none'
      if ($null -ne $fragment) {
        $fragmentSource = 'cfHtmlHeader'
      } elseif ($null -ne $htmlDecoded['text']) {
        # No usable CF_HTML header: fall back to the whole payload. Whether that
        # text is trustworthy is what validUtf8 and suspectedUtf16Le record.
        $fragment = $htmlDecoded['text']
        $fragmentSource = 'wholePayload'
      }
      $htmlObservation = Get-HtmlPayloadObservation -Fragment $fragment

      $textString = ConvertFrom-UnicodeTextBytes -Bytes $snapshot['unicodeText']['bytes']
      $textLines = @()
      $textLineCount = 0
      if ($null -ne $textString) {
        $splitLines = $textString -split "`r`n|`n"
        $textLineCount = $splitLines.Count
        $lineLimit = [Math]::Min($splitLines.Count, 5)
        for ($lineIndex = 0; $lineIndex -lt $lineLimit; $lineIndex++) {
          $line = [string] $splitLines[$lineIndex]
          if ($line.Length -gt 200) {
            $line = $line.Substring(0, 200) + '...'
          }
          $textLines += $line
        }
      }

      if ($snapshot['html']['available']) {
        $htmlPath = Write-PayloadFile -Directory $evidenceDirectory -FileName ($probeId + '-html.txt') -Bytes ($snapshot['html']['bytes'])
        $payloadFiles += [ordered]@{
          role       = 'clipboard-html'
          format     = 'HTML Format (CF_HTML)'
          file       = (Split-Path -Leaf $htmlPath)
          byteLength = [int] ($snapshot['html']['byteLength'])
        }
      }
      if ($snapshot['unicodeText']['available']) {
        # The raw CF_UNICODETEXT block is UTF-16LE; the file is the decoded
        # text written as UTF-8 without a BOM so it diffs and reads cleanly.
        # The original byte length is recorded beside it.
        $textPath = Write-TextPayloadFile -Directory $evidenceDirectory -FileName ($probeId + '-text.txt') -Text $textString
        $payloadFiles += [ordered]@{
          role       = 'clipboard-text'
          format     = 'CF_UNICODETEXT (decoded to UTF-8 without BOM)'
          file       = (Split-Path -Leaf $textPath)
          byteLength = [int] ($snapshot['unicodeText']['byteLength'])
        }
      }

      # The extracted fragment is the whole HTML document; it is already on
      # disk as the raw payload, so only the parsed facts go into the JSON.
      $cfHtmlRecord = [ordered]@{}
      foreach ($cfHtmlKey in $cfHtmlHeader.Keys) {
        if ($cfHtmlKey -ne 'fragment') {
          $cfHtmlRecord[$cfHtmlKey] = $cfHtmlHeader[$cfHtmlKey]
        }
      }
      $cfHtmlRecord['fragmentExtracted'] = ($null -ne $cfHtmlHeader['fragment'])

      $observation = [ordered]@{
        htmlFormatAvailable  = [bool] ($snapshot['html']['available'])
        htmlByteLength       = [int] ($snapshot['html']['byteLength'])
        htmlValidUtf8        = [bool] ($htmlDecoded['validUtf8'])
        htmlSuspectedUtf16Le = [bool] ($htmlDecoded['suspectedUtf16Le'])
        fragmentSource       = $fragmentSource
        cfHtmlHeaderPresent  = [bool] ($cfHtmlHeader['present'])
        cfHtml               = $cfHtmlRecord
        html                 = $htmlObservation
        textAvailable        = [bool] ($snapshot['unicodeText']['available'])
        textByteLength       = [int] ($snapshot['unicodeText']['byteLength'])
        textLineCount        = $textLineCount
        textFirstLines       = $textLines
      }
      if ($null -ne $snapshot['formats']) {
        $observation['clipboardFormats'] = $snapshot['formats']
      }

      $timing = [ordered]@{
        copyCallMs       = [Math]::Round($copyWatch.Elapsed.TotalMilliseconds, 1)
        clipboardReadMs  = [Math]::Round($readWatch.Elapsed.TotalMilliseconds, 1)
      }

      if ([bool] $probeDefinition.measurePaste) {
        # Paste Excel's own payload back into a clean sheet: the fastest paste
        # the format allows, which bounds what a browser payload can cost.
        $pasteSheet = $null
        $pasteTarget = $null
        try {
          $pasteSheet = New-ProbeSheet -Workbook $workbook -Name ($probeId + ' paste')
          $pasteTarget = $pasteSheet.Range('A1')
          $pasteWatch = [System.Diagnostics.Stopwatch]::StartNew()
          $excelPasteReturned = Invoke-ComActionWithRetry -Context "Pasting the $probeId payload into a clean sheet" -Action { $pasteSheet.Paste($pasteTarget) }
          $pasteWatch.Stop()
          $timing['excelPasteMs'] = [Math]::Round($pasteWatch.Elapsed.TotalMilliseconds, 1)
          $observation['excelPasteReturned'] = $excelPasteReturned

          $pastedRange = $null
          try {
            $pastedRange = $pasteSheet.UsedRange
            $pastedRows = Get-RangeDimension -Range $pastedRange -Kind Rows
            $pastedColumns = Get-RangeDimension -Range $pastedRange -Kind Columns
            $observation['pastedCellCount'] = ($pastedRows * $pastedColumns)
            $observation['pastedRangeAddress'] = [string] ($pastedRange.Address($false, $false))
          } finally {
            Release-ComObject $pastedRange
            $pastedRange = $null
          }
        } finally {
          Release-ComObject $pasteTarget
          $pasteTarget = $null
          Release-ComObject $pasteSheet
          $pasteSheet = $null
        }
      }

      $records.Add([ordered]@{
        id           = $probeId
        family       = [string] $probeDefinition.family
        title        = [string] $probeDefinition.title
        question     = [string] $probeDefinition.question
        lookFor      = [string] $probeDefinition.lookFor
        decisionRule = [string] $probeDefinition.decisionRule
        expectations = $probeDefinition.expectations
        payloads     = $payloadFiles
        timing       = $timing
        observation  = $observation
      })

      Write-Output ("  html {0} bytes, text {1} bytes, table cells {2}, x:fmla {3}" -f $observation['htmlByteLength'], $observation['textByteLength'], $htmlObservation['tableCellCount'], $htmlObservation['formulaAttributeCount'])
    } finally {
      try {
        $null = Invoke-ComActionWithRetry -Context 'Clearing the Excel copy marquee' -Action { $excel.CutCopyMode = 0 }
      } catch {
        $cleanupErrors.Add($_.Exception.Message)
      }
      Release-ComObject $range
      $range = $null
      Release-ComObject $worksheet
      $worksheet = $null
    }
  }

  # ---------------- Excel-target probes ----------------
  foreach ($probeDefinition in $selectedTargetProbes) {
    $probeId = [string] $probeDefinition.id
    Write-Output "Probe $probeId - $($probeDefinition.title)"

    $worksheet = $null
    $destination = $null
    $payloadFiles = @()

    try {
      $buildWatch = [System.Diagnostics.Stopwatch]::StartNew()
      try {
        $payload = & $probeDefinition.build
      } catch {
        throw (New-ComFailureMessage -Context "Building the payload for $probeId" -Exception $_.Exception)
      }
      $buildWatch.Stop()

      if ($null -ne $payload['html']) {
        $sentHtmlPath = Write-TextPayloadFile -Directory $evidenceDirectory -FileName ($probeId + '-sent-html.txt') -Text $payload['html']
        $payloadFiles += [ordered]@{
          role       = 'sent-html-document'
          format     = 'text/html wrapped in CF_HTML by the harness'
          file       = (Split-Path -Leaf $sentHtmlPath)
          byteLength = (ConvertTo-Utf8Bytes -Text $payload['html']).Length
        }
      }
      if ($null -ne $payload['text']) {
        $sentTextPath = Write-TextPayloadFile -Directory $evidenceDirectory -FileName ($probeId + '-sent-text.txt') -Text $payload['text']
        $payloadFiles += [ordered]@{
          role       = 'sent-tsv'
          format     = 'CF_UNICODETEXT'
          file       = (Split-Path -Leaf $sentTextPath)
          byteLength = (ConvertTo-UnicodeTextBytes -Text $payload['text']).Length
        }
      }

      $writeWatch = [System.Diagnostics.Stopwatch]::StartNew()
      $written = Set-ClipboardPayload -Html $payload['html'] -Text $payload['text'] -HtmlFormat $htmlFormatId
      $writeWatch.Stop()

      try {
        $worksheet = New-ProbeSheet -Workbook $workbook -Name ([string] $probeDefinition.sheetName)
      } catch {
        throw (New-ComFailureMessage -Context "Creating the paste sheet for $probeId" -Exception $_.Exception)
      }

      $destinationAddress = 'A1'
      if ($null -ne $probeDefinition.destination) {
        $destinationAddress = [string] $probeDefinition.destination
      }
      $destination = $worksheet.Range($destinationAddress)

      $pasteWatch = [System.Diagnostics.Stopwatch]::StartNew()
      # Worksheet.Paste answers True or False; a False is recorded rather than
      # thrown, because "Excel refused this payload" is itself the finding.
      $pasteReturned = Invoke-ComActionWithRetry -Context "Worksheet.Paste for $probeId" -Action { $worksheet.Paste($destination) }
      $pasteWatch.Stop()

      $usedRange = $worksheet.UsedRange
      $usedCellCount = (Get-RangeDimension -Range $usedRange -Kind Rows) * (Get-RangeDimension -Range $usedRange -Kind Columns)
      $usedRangeAddress = [string] ($usedRange.Address($false, $false))

      $inventoryPolicy = 'cells'
      if ($null -ne $probeDefinition.inventory) {
        $inventoryPolicy = [string] $probeDefinition.inventory
      }

      $inventory = $null
      if ($inventoryPolicy -ne 'none') {
        $inventory = Get-CellInventory -Range $usedRange -MaxCells 200
      }

      $observation = [ordered]@{
        pastedRangeAddress = $usedRangeAddress
        pastedCellCount    = $usedCellCount
        pasteReturned      = $pasteReturned
        rowHeights         = (Get-RowHeightMap -Range $usedRange)
        inventory          = $inventory
        samples            = $null
      }
      if ($inventoryPolicy -eq 'none') {
        $usedRowCount = Get-RangeDimension -Range $usedRange -Kind Rows
        $usedColumnCount = Get-RangeDimension -Range $usedRange -Kind Columns
        $observation['samples'] = [ordered]@{
          topLeft     = (Get-UsedRangeCellSummary -Range $usedRange -Row 1 -Column 1)
          middle      = (Get-UsedRangeCellSummary -Range $usedRange -Row ([int] (($usedRowCount + 1) / 2)) -Column 1)
          bottomRight = (Get-UsedRangeCellSummary -Range $usedRange -Row $usedRowCount -Column $usedColumnCount)
        }
      }

      $records.Add([ordered]@{
        id           = $probeId
        family       = [string] $probeDefinition.family
        title        = [string] $probeDefinition.title
        question     = [string] $probeDefinition.question
        lookFor      = [string] $probeDefinition.lookFor
        decisionRule = [string] $probeDefinition.decisionRule
        expectations = $probeDefinition.expectations
        payloads     = $payloadFiles
        timing       = [ordered]@{
          payloadBuildMs  = [Math]::Round($buildWatch.Elapsed.TotalMilliseconds, 1)
          clipboardWriteMs = [Math]::Round($writeWatch.Elapsed.TotalMilliseconds, 1)
          pasteMs         = [Math]::Round($pasteWatch.Elapsed.TotalMilliseconds, 1)
        }
        clipboard    = [ordered]@{
          formats        = $written['formats']
          htmlByteLength = $written['htmlByteLength']
          textByteLength = $written['textByteLength']
          cfHtmlHeader   = $written['cfHtmlHeader']
          verified       = [bool] ($written['verified'])
        }
        observation  = $observation
      })

      Write-Output ("  pasted {0} at {1}, {2} ms" -f $usedCellCount, $usedRangeAddress, $records[$records.Count - 1]['timing']['pasteMs'])
    } finally {
      Release-ComObject $destination
      $destination = $null
      Release-ComObject $usedRange
      $usedRange = $null
      Release-ComObject $worksheet
      $worksheet = $null
    }
  }

  $evidence = [ordered]@{
    manifest = $manifest
    probes   = $records
  }
  $jsonPath = Join-Path $evidenceDirectory 'clipboard-evidence.json'
  # UTF-8 without a BOM, through the same stable serialiser as
  # capture-goldens.ps1: Windows PowerShell 5.1 Set-Content -Encoding utf8 adds
  # a BOM and ConvertTo-Json differs between PowerShell 5.1 and 7.
  $json = ConvertTo-StableJson -Value $evidence
  [System.IO.File]::WriteAllText($jsonPath, $json + "`n", $utf8NoBom)
  $evidenceWritten = $true

  Write-Output "Wrote $jsonPath"
  Write-Output "Raw payloads are in $evidenceDirectory."
  # The target probes publish their own payloads, so whatever the maintainer had
  # on the clipboard before the run is already gone by this point.
  Write-Output 'The Windows clipboard was emptied by the target probes and cannot be restored; Windows also discards clipboard data owned by this process when it exits.'
} finally {
  # Release leaf objects first. Excel is only asked to quit after its child
  # RCWs are gone, so a rejected Quit is not caused by our own references.
  Release-ComObject $usedRange
  $usedRange = $null
  Release-ComObject $destination
  $destination = $null
  Release-ComObject $range
  $range = $null
  Release-ComObject $worksheet
  $worksheet = $null

  if ($null -ne $workbook) {
    try {
      # Never save: the probe workbook is scratch state, and a save would
      # create a file next to the evidence directory.
      Invoke-ComActionWithRetry -Context 'Closing the probe workbook' -Action {
        $workbook.Close($false) | Out-Null
      }
    } catch {
      $cleanupErrors.Add($_.Exception.Message)
    }
  }

  Release-ComObject $workbook
  $workbook = $null
  Release-ComObject $workbooks
  $workbooks = $null

  if ($null -ne $excel) {
    try {
      Invoke-ComActionWithRetry -Context 'Quitting Excel' -Action {
        $excel.Quit() | Out-Null
      }
    } catch {
      $cleanupErrors.Add($_.Exception.Message)
    }
  }

  Release-ComObject $excel
  $excel = $null

  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()

  if ($cleanupErrors.Count -gt 0) {
    $message = "Excel COM cleanup failed after the clipboard spike: $($cleanupErrors -join '; ')"
    if ($evidenceWritten) {
      # The evidence was written, but a persistent cleanup failure must not be
      # reported as a clean pass.
      throw $message
    }
    Write-Warning $message
  }
}
