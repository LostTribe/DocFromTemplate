<#
.SYNOPSIS
    Standalone CSV diagnostic for New-DocFromTemplate.ps1.

.DESCRIPTION
    Reports everything you need to know about why a CSV does or does not
    parse the way New-DocFromTemplate.ps1 expects:

      * File encoding and whether a byte-order mark (BOM) is present.
      * The delimiter the file most likely uses (best guess from the
        header line), alongside the delimiter you asked for.
      * The header list EXACTLY as Import-Csv sees it (same parser the
        main script uses), so you can spot duplicates, stray empty
        columns from trailing delimiters, or a BOM stuck to header #1.
      * Physical line count vs. the number of data rows Import-Csv
        returns - a mismatch usually means a delimiter/encoding problem
        or that the file is empty below the header.

    No external modules. Works on Windows PowerShell 5.1 and PowerShell 7+.
    It never modifies the file - read-only.

.PARAMETER CsvPath
    Path to the .csv file to inspect.

.PARAMETER Delimiter
    Delimiter to test with (default ','). Try ';' or "`t" if the guess
    below disagrees with the comma default.

.EXAMPLE
    .\Diagnose-Csv.ps1 -CsvPath ".\Template_AMS_Data_vars.csv"

.EXAMPLE
    .\Diagnose-Csv.ps1 -CsvPath ".\data.csv" -Delimiter ';'
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string] $CsvPath,
    [ValidateLength(1, 1)] [string] $Delimiter = ','
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $CsvPath)) { throw "CSV file not found: $CsvPath" }
$CsvPath = (Resolve-Path -LiteralPath $CsvPath).Path

# --- Raw bytes + encoding / BOM ------------------------------------------
$bytes = [System.IO.File]::ReadAllBytes($CsvPath)

function Get-CsvFileEncoding {
    param([Parameter(Mandatory)] [byte[]] $Bytes)
    if ($Bytes.Length -ge 3 -and $Bytes[0] -eq 0xEF -and $Bytes[1] -eq 0xBB -and $Bytes[2] -eq 0xBF) {
        return [pscustomobject]@{ Name = 'UTF-8 with BOM'; Encoding = (New-Object System.Text.UTF8Encoding $true); Bom = $true }
    }
    if ($Bytes.Length -ge 2 -and $Bytes[0] -eq 0xFF -and $Bytes[1] -eq 0xFE) {
        return [pscustomobject]@{ Name = 'UTF-16 LE'; Encoding = [System.Text.Encoding]::Unicode; Bom = $true }
    }
    if ($Bytes.Length -ge 2 -and $Bytes[0] -eq 0xFE -and $Bytes[1] -eq 0xFF) {
        return [pscustomobject]@{ Name = 'UTF-16 BE'; Encoding = [System.Text.Encoding]::BigEndianUnicode; Bom = $true }
    }
    return [pscustomobject]@{ Name = 'UTF-8 / ANSI (no BOM)'; Encoding = (New-Object System.Text.UTF8Encoding $false); Bom = $false }
}

$encInfo  = Get-CsvFileEncoding -Bytes $bytes
$bomBytes = if ($encInfo.Bom) {
    $n = if ($encInfo.Name -eq 'UTF-8 with BOM') { 3 } else { 2 }
    ($bytes[0..($n - 1)] | ForEach-Object { '{0:X2}' -f $_ }) -join ' '
} else { '(none)' }

# Decode with the detected encoding, then strip a leading BOM char so it is
# not mistaken for part of the first header.
$text      = $encInfo.Encoding.GetString($bytes).TrimStart([char]0xFEFF)
$lines     = $text -split "`r`n|`n|`r"
# A trailing newline yields one empty element; do not count it as a line.
$lineCount = ($lines | Where-Object { $_.Length -gt 0 }).Count
$firstLine = if ($lines.Count -gt 0) { $lines[0] } else { '' }

# --- Delimiter guess from the header line --------------------------------
$candidates = @{ ',' = 'comma'; ';' = 'semicolon'; "`t" = 'tab'; '|' = 'pipe' }
$counts = foreach ($c in $candidates.Keys) {
    [pscustomobject]@{
        Char  = $c
        Name  = $candidates[$c]
        Count = ($firstLine.ToCharArray() | Where-Object { $_ -eq $c }).Count
    }
}
$guess = $counts | Sort-Object Count -Descending | Select-Object -First 1
$guessText = if ($guess.Count -gt 0) { "$($guess.Name) (x$($guess.Count) in header)" } else { '(no common delimiter found in header)' }
$delimDisplay = if ($Delimiter -eq "`t") { 'tab' } else { "'$Delimiter'" }

# --- Headers as Import-Csv would see them (via the same parser) ----------
$delimCount   = ($firstLine.ToCharArray() | Where-Object { $_ -eq $Delimiter[0] }).Count
$synthHeaders = 0..($delimCount + 1) | ForEach-Object { "__probe$_" }
$parsed       = $firstLine | ConvertFrom-Csv -Delimiter $Delimiter -Header $synthHeaders
$headers      = foreach ($name in $synthHeaders) {
    $v = $parsed.$name
    if ($null -ne $v) { [string]$v }
}
$dupes = $headers | Group-Object | Where-Object { $_.Count -gt 1 }

# --- Does Import-Csv treat the header line as a comment? ------------------
# Import-Csv skips any record whose first field begins with '#'. Rule
# (matches Import-Csv): trim leading whitespace, peel one optional leading
# double-quote, then test for a leading '#'. A '#'-first header is the
# classic cause of an unexplained "no data rows" / 0-row import.
$commentProbe = $firstLine.TrimStart()
if ($commentProbe.StartsWith('"')) { $commentProbe = $commentProbe.Substring(1) }
$headerIsComment = $commentProbe.StartsWith('#')

# --- Row count via Import-Csv (the real thing) ---------------------------
$rowCount   = $null
$importError = $null
try {
    $rowCount = @(Import-Csv -LiteralPath $CsvPath -Delimiter $Delimiter).Count
} catch {
    $importError = $_.Exception.Message
}

# --- Report --------------------------------------------------------------
Write-Host ''
Write-Host '===== CSV diagnostics =====' -ForegroundColor Cyan
Write-Host "File              : $CsvPath"
Write-Host "Size (bytes)      : $($bytes.Length)"
Write-Host "Encoding          : $($encInfo.Name)"
Write-Host "BOM bytes         : $bomBytes"
Write-Host "Physical lines    : $lineCount (non-empty)"
Write-Host "Delimiter (asked) : $delimDisplay"
Write-Host "Delimiter (guess) : $guessText"
if ($guess.Count -gt 0 -and $guess.Char -ne $Delimiter) {
    Write-Host "  -> Guess differs from the delimiter you asked for. Re-run with -Delimiter accordingly." -ForegroundColor Yellow
}

Write-Host ''
Write-Host "Headers ($($headers.Count)) as Import-Csv sees them:" -ForegroundColor Cyan
for ($i = 0; $i -lt $headers.Count; $i++) {
    $h = $headers[$i]
    $shown = if ([string]::IsNullOrEmpty($h)) { '<empty>' } else { $h }
    Write-Host ("  [{0}] {1}" -f ($i + 1), $shown)
}
if ($dupes) {
    Write-Host 'Duplicate / empty headers detected:' -ForegroundColor Yellow
    foreach ($d in $dupes) {
        if ($d.Name) { Write-Host "  '$($d.Name)' appears $($d.Count) times" -ForegroundColor Yellow }
        else         { Write-Host "  $($d.Count) empty header(s) - usually trailing delimiters" -ForegroundColor Yellow }
    }
}
if ($headerIsComment) {
    Write-Host "  -> Header's first column starts with '#'. Import-Csv treats the whole header line as a COMMENT and skips it, so you get garbage columns or 0 rows. Add a 'title' (or any non-'#') first column." -ForegroundColor Yellow
}

Write-Host ''
if ($importError) {
    Write-Host "Import-Csv FAILED : $importError" -ForegroundColor Red
} else {
    Write-Host "Data rows (Import-Csv) : $rowCount" -ForegroundColor Cyan
    if ($rowCount -eq 0) {
        Write-Host '  -> 0 rows. The header parsed but there is no data below it, or every line collapsed into the header row (delimiter/encoding mismatch).' -ForegroundColor Yellow
    }
    $expected = [Math]::Max(0, $lineCount - 1)
    if ($rowCount -ne $expected) {
        Write-Host "  -> Note: $lineCount physical lines imply ~$expected data rows, but Import-Csv returned $rowCount. Check for embedded newlines in quoted fields, or a delimiter mismatch." -ForegroundColor Yellow
    }
}
Write-Host ''
