<#
.SYNOPSIS
    Diagnose a CSV file: encoding, delimiter, header tokens, and what
    Import-Csv would actually see. Run this when New-DocFromTemplate.ps1
    -CsvPath misbehaves.

.DESCRIPTION
    Reports six things in one PSCustomObject, all of which point at a
    specific kind of trouble when they look wrong:

      FileSizeBytes      sanity check the file is what you think
      FirstFourBytes     the BOM (or absence of one), in hex
      DetectedEncoding   UTF-8 (no BOM) | UTF-8 BOM | UTF-16 LE | UTF-16 BE
      Delimiter          the delimiter that was probed (',' / ';' / TAB)
      RawLineCount       how many non-empty lines are in the file
      HeaderLine         the literal first line as a string
      HeaderTokenCount   how many tokens the line parsed into
      HeaderTokens       the parsed tokens as a string[]
      DuplicateHeaders   any token name that appears more than once
      ImportCsvRows      what Import-Csv -Delimiter <d> actually built
      ImportCsvError     Import-Csv's own error message, if it threw

    The header tokeniser is CSV-aware (respects quoted fields, doubled ""
    escapes, the supplied delimiter), so what shows up in HeaderTokens is
    exactly what Import-Csv would have seen. If HeaderTokens disagrees
    with what you typed into the file, the file's contents are different
    from what you think.

    Open in Excel or VS Code while the script runs - it opens with
    FileShare.ReadWrite, so an open editor will not block the probe.

.PARAMETER Path
    Path to the .csv file to inspect.

.PARAMETER Delimiter
    Field delimiter to test the file with. Defaults to ','. Pass ';' for
    semicolon-delimited files (common in non-English Excel locales), or
    "`t" for tab-delimited files (Excel's "Unicode Text" save).

.EXAMPLE
    .\tools\Test-CsvFile.ps1 -Path .\data.csv | Format-List

.EXAMPLE
    # Same file, three delimiters - the one with non-zero ImportCsvRows
    # and a sensible HeaderTokenCount is the right one.
    .\tools\Test-CsvFile.ps1 -Path .\data.csv                | Format-List
    .\tools\Test-CsvFile.ps1 -Path .\data.csv -Delimiter ';' | Format-List
    .\tools\Test-CsvFile.ps1 -Path .\data.csv -Delimiter "`t" | Format-List
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string] $Path,
    [char] $Delimiter = ','
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $Path)) {
    throw "File not found: $Path"
}

# ---- BOM ----
# Open shared-read so a CSV currently open in Excel can still be probed.
$bytes = New-Object byte[] 4
$fs = [System.IO.File]::Open($Path, 'Open', 'Read', 'ReadWrite')
try { [void]$fs.Read($bytes, 0, 4) } finally { $fs.Close() }
$hex = ($bytes | ForEach-Object { '{0:X2}' -f $_ }) -join ' '

$encoding = [System.Text.Encoding]::UTF8
$encName  = 'UTF-8 (no BOM)'
if     ($bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) { $encName = 'UTF-8 BOM' }
elseif ($bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE)                        { $encoding = [System.Text.Encoding]::Unicode;        $encName = 'UTF-16 LE' }
elseif ($bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF)                        { $encoding = [System.Text.Encoding]::BigEndianUnicode; $encName = 'UTF-16 BE' }

# ---- File size + full content (decoded with the detected encoding) ----
$size = (Get-Item -LiteralPath $Path).Length
$fs2  = [System.IO.File]::Open($Path, 'Open', 'Read', 'ReadWrite')
$sr   = [System.IO.StreamReader]::new($fs2, $encoding, $true)
try { $allText = $sr.ReadToEnd() } finally { $sr.Close(); $fs2.Close() }

$rawLines = ($allText -split "`r`n|`r|`n" | Where-Object { $_ -ne '' }).Count

# ---- Header tokens (CSV-aware) ----
$headerLine = ($allText -split "`r`n|`r|`n")[0]
$tokens     = New-Object 'System.Collections.Generic.List[string]'
$sb         = New-Object System.Text.StringBuilder
$inQuotes   = $false
for ($i = 0; $i -lt $headerLine.Length; $i++) {
    $c = $headerLine[$i]
    if ($inQuotes) {
        if ($c -eq '"') {
            if ($i + 1 -lt $headerLine.Length -and $headerLine[$i + 1] -eq '"') {
                [void]$sb.Append('"'); $i++
            } else {
                $inQuotes = $false
            }
        } else {
            [void]$sb.Append($c)
        }
    } else {
        if ($c -eq $Delimiter) {
            [void]$tokens.Add($sb.ToString().Trim()); [void]$sb.Clear()
        } elseif ($c -eq '"' -and $sb.Length -eq 0) {
            $inQuotes = $true
        } else {
            [void]$sb.Append($c)
        }
    }
}
[void]$tokens.Add($sb.ToString().Trim())

# ---- What Import-Csv actually does with the file + delimiter ----
$rowCount  = 0
$importErr = $null
try {
    $rowCount = @(Import-Csv -LiteralPath $Path -Delimiter $Delimiter).Count
} catch {
    $importErr = $_.Exception.Message
}

# ---- Report ----
[pscustomobject]@{
    Path             = (Resolve-Path -LiteralPath $Path).Path
    FileSizeBytes    = $size
    FirstFourBytes   = $hex
    DetectedEncoding = $encName
    Delimiter        = if ($Delimiter -eq "`t") { 'TAB' } else { $Delimiter.ToString() }
    RawLineCount     = $rawLines
    HeaderLine       = $headerLine
    HeaderTokenCount = $tokens.Count
    HeaderTokens     = $tokens.ToArray()
    DuplicateHeaders = ($tokens | Group-Object | Where-Object Count -gt 1 |
                        ForEach-Object {
                            if ($_.Name) { "'$($_.Name)' (x$($_.Count))" }
                            else         { "$($_.Count) empty header(s)" }
                        }) -join ', '
    ImportCsvRows    = $rowCount
    ImportCsvError   = $importErr
}
