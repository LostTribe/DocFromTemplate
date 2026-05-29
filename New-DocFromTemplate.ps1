<#
.SYNOPSIS
    Fills a Word .docx template by replacing placeholder strings. Values can
    come from a CSV file, a JSON file, a key=value text file, or a hashtable
    supplied directly. No external module dependencies - everything uses
    built-in PowerShell.

.DESCRIPTION
    Four parameter sets:

    FromCsv (default)
        Reads -CsvPath via Import-Csv. Each column header is a placeholder,
        each row produces one document. Flat single table - no per-group
        nesting. No external module dependency.

    FromJson
        Reads -JsonPath via ConvertFrom-Json. If the file parses to an
        array of objects, each element becomes one document. If it parses
        to a single object, one document is produced (like FromValues).
        Nested objects/arrays in property values are coerced to string by
        PowerShell and likely render as "@{...}" or "System.Object[]".

    FromKeyValue
        Reads -KeyValuePath line by line. Each non-blank, non-comment line
        is split on the first '=' into a key and a value. Blank lines and
        lines starting with ';' are ignored ('#' is reserved for placeholder
        names like #replace1, so it cannot also be the comment marker).
        Produces a single document.

    FromValues
        Substitutes the placeholder/value pairs in -Values into the template
        once and writes a single document. Useful for one-off renders that
        do not warrant a file.

    Naming rules (tabular sources: FromCsv, FromJson-array):
      * If a row has a 'title' column (case-insensitive) its value is the
        output filename.
      * Otherwise a single-row source produces <source>.docx, and a
        multi-row source produces <source>_1.docx, <source>_2.docx, ...
        ('source' is the CSV or JSON file basename.)
      * Duplicates within a run get a numeric suffix to disambiguate.

    Naming rules (single-doc sources: FromValues, FromJson-object,
    FromKeyValue):
      * Use -OutputName if provided, otherwise:
          FromValues   -> <template-base>-filled.docx
          FromJson     -> <json-base>-filled.docx
          FromKeyValue -> <kv-base>-filled.docx

    The template file is opened read-only and is never modified.

    Customising behaviour:
        Three customisation hooks are defined near the top of this script.
        They are no-ops by default:

            Invoke-RowPreProcess     -- mutate row values before replacement
            Get-CustomOutputFileName -- override the output filename rule
            Invoke-DocPostProcess    -- modify the open Word document
                                        after replacements but before save

        In single-doc modes (FromValues, FromJson-object, FromKeyValue),
        Invoke-RowPreProcess and Get-CustomOutputFileName are NOT called.
        Invoke-DocPostProcess still fires; $Row is a PSCustomObject
        synthesised from the values and $SheetName is '(values)', '(json)',
        or '(keyvalue)' respectively.

.PARAMETER CsvPath
    Path to the source .csv file. Required in FromCsv mode.

.PARAMETER JsonPath
    Path to the source .json file. Required in FromJson mode.

.PARAMETER KeyValuePath
    Path to the source key=value text file. Required in FromKeyValue mode.

.PARAMETER Values
    Hashtable of placeholder => value pairs. Required in FromValues mode.

.PARAMETER TemplatePath
    Path to the Word .docx template. Required in every mode.

.PARAMETER OutputDir
    Folder to write filled documents into. Defaults to an 'output' folder
    next to the template. Created if it does not exist.

.PARAMETER OutputName
    FromValues, FromJson, FromKeyValue. Output filename without extension.
    Ignored if the JSON file parses to an array (filenames come from each
    row's 'title' column or the fallback rule instead).

.EXAMPLE
    .\New-DocFromTemplate.ps1 -CsvPath .\rows.csv `
                              -TemplatePath .\template.docx

.EXAMPLE
    .\New-DocFromTemplate.ps1 -JsonPath .\rows.json `
                              -TemplatePath .\template.docx

.EXAMPLE
    .\New-DocFromTemplate.ps1 -KeyValuePath .\data.env `
                              -TemplatePath .\letter.docx `
                              -OutputName 'one-off'

.EXAMPLE
    .\New-DocFromTemplate.ps1 -TemplatePath .\letter.docx `
                              -Values @{ '#replace1' = 'Jane Doe'
                                         '#replace2' = 'Engineer' } `
                              -OutputName 'Jane-welcome'
#>

[CmdletBinding(DefaultParameterSetName = 'FromCsv')]
param(
    [Parameter(Mandatory, ParameterSetName = 'FromCsv')]      [string]    $CsvPath,

    [Parameter(Mandatory, ParameterSetName = 'FromJson')]     [string]    $JsonPath,

    [Parameter(Mandatory, ParameterSetName = 'FromKeyValue')] [string]    $KeyValuePath,

    [Parameter(Mandatory, ParameterSetName = 'FromValues')]   [hashtable] $Values,

    [Parameter(ParameterSetName = 'FromValues')]
    [Parameter(ParameterSetName = 'FromJson')]
    [Parameter(ParameterSetName = 'FromKeyValue')]            [string]    $OutputName,

    [Parameter(Mandatory)]                                    [string]    $TemplatePath,

    [string] $OutputDir
)

$ErrorActionPreference = 'Stop'


# Customisation hooks
#
# These three functions are called at well-defined points during the merge.
# Out of the box they return the input unchanged, so they are safe to leave
# alone. Replace their bodies to inject your own behaviour.

<#
    Invoke-RowPreProcess

    Called once per row, BEFORE any replacement happens. Use it to:
      * Compute derived columns (e.g. add a #today field built from
        the current date).
      * Transform raw cell values (e.g. format a number as currency,
        upper-case a code, look something up from another data source).
      * Drop a row by returning $null -- the main loop will skip any
        row whose pre-processed value is $null.

    Parameters:
      $Row       -- the original PSCustomObject from the source
      $SheetName -- a tag identifying the source (CSV/JSON file basename
                    in tabular mode; '(values)' / '(json)' / '(keyvalue)'
                    in single-doc modes)

    Return:
      A PSCustomObject (the row to use going forward) or $null to skip.
#>
function Invoke-RowPreProcess {
    param(
        [Parameter(Mandatory)] $Row,
        [Parameter(Mandatory)] [string] $SheetName
    )

    # Default behaviour: pass the row through unchanged.
    return $Row

    # Example -- add a #today placeholder built from the current date:
    #   $Row | Add-Member -NotePropertyName '#today' `
    #                     -NotePropertyValue (Get-Date -Format 'yyyy-MM-dd') `
    #                     -Force
    #   return $Row
}


<#
    Get-CustomOutputFileName

    Called once per row to decide the output filename. Return $null to fall
    back to the built-in rule (title column, then <source>.docx /
    <source>_<n>.docx).

    Parameters:
      $Row       -- the (possibly pre-processed) row
      $SheetName -- a tag identifying the source (CSV/JSON file
                    basename in tabular mode)
      $RowNumber -- 1-based index within the source

    Return:
      A filename WITHOUT extension (the script appends .docx) or $null.
#>
function Get-CustomOutputFileName {
    param(
        [Parameter(Mandatory)] $Row,
        [Parameter(Mandatory)] [string] $SheetName,
        [Parameter(Mandatory)] [int]    $RowNumber
    )

    # Default behaviour: defer to the built-in naming rule.
    return $null

    # Example -- name files '<sheet>-<title>':
    #   $titleProp = $Row.PSObject.Properties |
    #                Where-Object { $_.Name -ieq 'title' } |
    #                Select-Object -First 1
    #   if ($titleProp -and $titleProp.Value) {
    #       return "$SheetName-$($titleProp.Value)"
    #   }
    #   return $null
}


<#
    Invoke-DocPostProcess

    Called once per row, AFTER all placeholders have been replaced but
    BEFORE the document is saved. Use it to insert dynamic content that
    cannot be expressed as a simple find & replace, e.g.:
      * Insert an image
      * Adjust styles
      * Update a table of contents
      * Inject the current date into the header

    Parameters:
      $Document  -- the open Word.Document COM object
      $Row       -- the row that drove this document
      $SheetName -- a tag identifying the source (CSV/JSON file
                    basename in tabular mode; '(values)' /
                    '(json)' / '(keyvalue)' in single-doc modes)

    Return:
      Anything; the return value is ignored.
#>
function Invoke-DocPostProcess {
    param(
        [Parameter(Mandatory)] $Document,
        [Parameter(Mandatory)] $Row,
        [Parameter(Mandatory)] [string] $SheetName
    )

    # Default behaviour: do nothing.
    return

    # Example -- update all fields in the document (useful if your
    # template contains TOC, DATE, REF or PAGE fields):
    #   $Document.Fields.Update() | Out-Null
}


# Input validation

if (-not (Test-Path $TemplatePath)) { throw "Template not found: $TemplatePath" }

switch ($PSCmdlet.ParameterSetName) {
    'FromCsv'      { if (-not (Test-Path $CsvPath))      { throw "CSV file not found: $CsvPath" } }
    'FromJson'     { if (-not (Test-Path $JsonPath))     { throw "JSON file not found: $JsonPath" } }
    'FromKeyValue' { if (-not (Test-Path $KeyValuePath)) { throw "Key=value file not found: $KeyValuePath" } }
    'FromValues'   { }
}

# Word COM does not understand relative paths, so resolve everything to
# fully qualified paths up front.
$TemplatePath = (Resolve-Path $TemplatePath).Path
switch ($PSCmdlet.ParameterSetName) {
    'FromCsv'      { $CsvPath      = (Resolve-Path $CsvPath     ).Path }
    'FromJson'     { $JsonPath     = (Resolve-Path $JsonPath    ).Path }
    'FromKeyValue' { $KeyValuePath = (Resolve-Path $KeyValuePath).Path }
}

if (-not $OutputDir) {
    $OutputDir = Join-Path (Split-Path $TemplatePath) 'output'
}
if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir | Out-Null
}
$OutputDir = (Resolve-Path $OutputDir).Path


# Helpers

<#
    Get-SafeFileName

    Strip characters that NTFS will reject in a filename. Returns $null if
    the input is empty or sanitisation removes everything, so the caller
    can fall back to a default naming rule.
#>
function Get-SafeFileName {
    param([string] $Name)

    if ([string]::IsNullOrWhiteSpace($Name)) { return $null }

    $invalid = [IO.Path]::GetInvalidFileNameChars()
    $clean   = -join ($Name.ToCharArray() | Where-Object { $invalid -notcontains $_ })
    $clean   = $clean.Trim()

    if ([string]::IsNullOrWhiteSpace($clean)) { return $null }
    return $clean
}


<#
    Read-KeyValueFile

    Parse a flat key=value text file into a hashtable of placeholder =>
    value pairs. Rules:
      * Blank lines are skipped.
      * Lines whose first non-whitespace character is ';' are comments
        and are skipped. ('#' is reserved for placeholder names like
        #replace1, so it cannot also be the comment marker.)
      * Other lines are split on the FIRST '='. Anything before is the
        key; anything after is the value. Both are trimmed.
      * Lines without an '=' produce a warning and are skipped.
#>
function Read-KeyValueFile {
    param([Parameter(Mandatory)] [string] $Path)

    $result  = @{}
    $lineNum = 0
    foreach ($line in (Get-Content -LiteralPath $Path)) {
        $lineNum++
        $trimmed = $line.Trim()
        if ([string]::IsNullOrEmpty($trimmed)) { continue }
        if ($trimmed.StartsWith(';'))          { continue }

        $eq = $trimmed.IndexOf('=')
        if ($eq -lt 1) {
            Write-Warning "${Path} line ${lineNum}: missing '='; skipping."
            continue
        }
        $key = $trimmed.Substring(0, $eq).Trim()
        $val = $trimmed.Substring($eq + 1).Trim()
        if ($key) {
            $result[$key] = $val
        }
    }
    return $result
}


# Word constants
# Numeric values used by Word.Application's COM API. Naming them here makes
# the call sites readable instead of being littered with magic numbers.

$wdFindStop              = 0
$wdCollapseEnd           = 0
$wdFormatDocumentDefault = 16   # Modern .docx (OOXML)


<#
    Save-FilledDocument

    Opens the template read-only, runs every (placeholder, value) pair
    across all story ranges, calls Invoke-DocPostProcess, then saves to
    $OutPath. Shared across every parameter set (tabular and single-doc) so the
    replacement logic lives in exactly one place.

    Parameters:
      $Word          -- the open Word.Application
      $TemplatePath  -- absolute path to the template
      $Replacements  -- hashtable of placeholder => string
      $OutPath       -- absolute path to write the .docx to
      $Row           -- passed through to Invoke-DocPostProcess
      $SheetName     -- passed through to Invoke-DocPostProcess
                        ('(values)' in FromValues mode)
#>
function Save-FilledDocument {
    param(
        [Parameter(Mandatory)] $Word,
        [Parameter(Mandatory)] [string]    $TemplatePath,
        [Parameter(Mandatory)] [hashtable] $Replacements,
        [Parameter(Mandatory)] [string]    $OutPath,
        $Row,
        [Parameter(Mandatory)] [string]    $SheetName
    )

    # Open the template read-only and run the replacements.
    # Documents.Open(FileName, ConfirmConversions, ReadOnly).
    # ReadOnly = $true guarantees we never modify the source file.
    $doc = $Word.Documents.Open($TemplatePath, $false, $true)
    try {
        # Iterate every story range so headers, footers, footnotes
        # and text boxes are covered as well as the main body.
        foreach ($story in $doc.StoryRanges) {
            $current = $story
            while ($null -ne $current) {

                # Process longer placeholders first so #replace1 doesn't
                # eat the prefix of #replace10/11/... (substring match).
                foreach ($placeholder in ($Replacements.Keys | Sort-Object -Property Length -Descending)) {

                    # Word's Find.Replacement.Text is capped at 255
                    # characters; longer values throw "String parameter
                    # too long". We sidestep that by finding each match
                    # and assigning Range.Text (no length cap).
                    #
                    # Duplicating $current means Find can mutate the
                    # range freely without us losing our outer iterator.
                    $rng  = $current.Duplicate
                    $find = $rng.Find
                    $find.ClearFormatting()

                    $find.Text              = $placeholder
                    $find.Forward           = $true
                    $find.Wrap              = $script:wdFindStop
                    $find.Format            = $false
                    $find.MatchCase         = $true
                    $find.MatchWholeWord    = $false
                    $find.MatchWildcards    = $false
                    $find.MatchSoundsLike   = $false
                    $find.MatchAllWordForms = $false

                    # Find.Execute collapses $rng onto the match. We
                    # assign .Text to replace it, then collapse to the
                    # end so the next Execute starts after the
                    # substitution.
                    while ($find.Execute()) {
                        $rng.Text = $Replacements[$placeholder]
                        $rng.Collapse($script:wdCollapseEnd)
                    }
                }

                $current = $current.NextStoryRange
            }
        }

        # Customisation hook: post-process the document.
        Invoke-DocPostProcess -Document $doc -Row $Row -SheetName $SheetName | Out-Null

        # Save and close. SaveAs2 (rather than SaveAs) avoids the
        # PowerShell 7 [ref]-marshalling quirk that would otherwise
        # throw "Cannot convert ... value of type psobject".
        $doc.SaveAs2($OutPath, $script:wdFormatDocumentDefault)
    }
    finally {
        # Close without saving back to the template path.
        $doc.Close($false)
    }
}


# Source loading
# Each input source is normalised into one of two shapes:
#
#   $dataSheets       -- @( @{ Name='...'; Rows=@(<PSCustomObject>, ...) }, ... )
#                        used by the tabular path (FromCsv,
#                        FromJson when the file is an array).
#
#   $singleValues +   -- hashtable + base filename (no extension) + a tag
#   $singleOutputName    that goes into the result's Sheet column. Used by
#   + $singleSheetTag    the single-doc path (FromValues, FromJson when the
#                        file is a single object, FromKeyValue).
#
# After this block exactly one of those two is populated.

$dataSheets       = $null
$singleValues     = $null
$singleOutputName = $null
$singleSheetTag   = $null

switch ($PSCmdlet.ParameterSetName) {
    'FromCsv' {
        # Pre-validate the header row before handing the file to Import-Csv.
        # Import-Csv refuses duplicate column headers with the cryptic error
        # "The member 'X' is already present", which fires before our own
        # error handling runs. Detect the problem here so the user sees a
        # message that names the duplicate and points at the file.
        #
        # Two common shapes hit the duplicate check:
        #   * A real repeated header ('name,name,#role')
        #   * Trailing commas creating multiple empty-string headers
        #     ('title,#a,#b,,') - the user usually doesn't realise their
        #     CSV ends with stray columns
        $headerLine = Get-Content -LiteralPath $CsvPath -TotalCount 1
        if (-not $headerLine) { throw "CSV is empty: $CsvPath" }
        $headers = $headerLine -split ',' | ForEach-Object { $_.Trim().Trim('"') }
        $dupes = $headers | Group-Object | Where-Object Count -gt 1
        if ($dupes) {
            $detail = foreach ($d in $dupes) {
                if ($d.Name) { "'$($d.Name)' (x$($d.Count))" }
                else         { "$($d.Count) empty header(s) - usually trailing commas" }
            }
            throw "CSV has duplicate column header(s): $($detail -join ', '). Each placeholder name must be unique. File: $CsvPath"
        }

        $dataSheets = @( [pscustomobject]@{
            Name = [IO.Path]::GetFileNameWithoutExtension($CsvPath)
            Rows = @(Import-Csv -LiteralPath $CsvPath)
        } )
    }

    'FromJson' {
        $rawJson = Get-Content -LiteralPath $JsonPath -Raw
        $json    = $rawJson | ConvertFrom-Json
        if ($null -eq $json) { throw "JSON parsed to null: $JsonPath" }

        # ConvertFrom-Json auto-unwraps single-element arrays in Windows
        # PowerShell (and PS 7 without -NoEnumerate), so $json -is [array]
        # is unreliable when there's exactly one row. Inspect the raw text
        # instead: a JSON document whose first non-whitespace character is
        # '[' is an array, regardless of element count.
        $looksLikeArray = $rawJson.TrimStart() -match '^\['

        if ($looksLikeArray) {
            # Array of objects -> tabular, one document per element.
            # @($json) re-wraps the single-element case so the row loop sees
            # an array either way.
            $dataSheets = @( [pscustomobject]@{
                Name = [IO.Path]::GetFileNameWithoutExtension($JsonPath)
                Rows = @($json)
            } )
            if ($OutputName) {
                Write-Warning "-OutputName is ignored when JSON parses to an array (filenames come from each row's 'title' column or the fallback rule)."
            }
        }
        else {
            # Single object -> one document, treated like FromValues.
            $singleValues = @{}
            foreach ($prop in $json.PSObject.Properties) {
                $singleValues[[string]$prop.Name] = $prop.Value
            }
            $singleOutputName = if ($OutputName) { $OutputName } `
                                else { [IO.Path]::GetFileNameWithoutExtension($JsonPath) + '-filled' }
            $singleSheetTag = '(json)'
        }
    }

    'FromKeyValue' {
        $singleValues     = Read-KeyValueFile -Path $KeyValuePath
        $singleOutputName = if ($OutputName) { $OutputName } `
                            else { [IO.Path]::GetFileNameWithoutExtension($KeyValuePath) + '-filled' }
        $singleSheetTag   = '(keyvalue)'
    }

    'FromValues' {
        $singleValues     = $Values
        $singleOutputName = if ($OutputName) { $OutputName } `
                            else { [IO.Path]::GetFileNameWithoutExtension($TemplatePath) + '-filled' }
        $singleSheetTag   = '(values)'
    }
}


# Main merge
# A single Word.Application instance is reused across every row. Spinning
# Word up and down per row is slow and error-prone.

$word      = $null
$results   = @()
$usedNames = @{}

try {
    $word               = New-Object -ComObject Word.Application
    $word.Visible       = $false
    $word.DisplayAlerts = 0

    if ($singleValues) {

        # Single-doc path: values supplied directly (FromValues, FromJson
        # when the file is a single object, FromKeyValue).
        # Row-level hooks are skipped; the user / adapter is in full
        # control. Invoke-DocPostProcess still fires inside
        # Save-FilledDocument.

        $replacements = @{}
        foreach ($key in $singleValues.Keys) {
            if ([string]::IsNullOrEmpty($key)) { continue }
            $name = [string]$key
            if ($name -notmatch '^#') {
                Write-Warning "Placeholder '$name' has no leading '#'; using verbatim."
            }
            $val = if ($null -eq $singleValues[$key]) { '' } else { [string]$singleValues[$key] }
            $replacements[$name] = $val
        }

        $baseName = Get-SafeFileName $singleOutputName
        if (-not $baseName) {
            throw "Output name '$singleOutputName' sanitised to an empty string."
        }
        $outPath = Join-Path $OutputDir "$baseName.docx"

        # Synthesise a $Row for Invoke-DocPostProcess so hook code that
        # reads $Row.SomeProp behaves the same as in tabular modes.
        $synthRow = [pscustomobject]$singleValues

        Save-FilledDocument -Word $word `
                            -TemplatePath $TemplatePath `
                            -Replacements $replacements `
                            -OutPath $outPath `
                            -Row $synthRow `
                            -SheetName $singleSheetTag

        $results += [pscustomobject]@{ Sheet = $singleSheetTag; Row = 1; Output = $outPath }
        Write-Host "Wrote $outPath"
    }
    else {

        # Tabular path: one document per row across every sheet bundle
        # (FromCsv, FromJson when the file is an array).
        foreach ($sheetBundle in $dataSheets) {
            $sheet = $sheetBundle.Name
            $rows  = @($sheetBundle.Rows)

            if (-not $rows) {
                Write-Warning "Sheet '$sheet' has no data rows; skipping."
                continue
            }

            for ($i = 0; $i -lt $rows.Count; $i++) {
                $rowNum = $i + 1

                # Customisation hook: pre-process the row.
                $row = Invoke-RowPreProcess -Row $rows[$i] -SheetName $sheet
                if ($null -eq $row) {
                    Write-Verbose "Sheet '$sheet' row ${rowNum}: skipped by Invoke-RowPreProcess."
                    continue
                }

                # Build the placeholder -> value map.
                # Every column except 'title' becomes a find/replace pair. The
                # column header is the search string verbatim, so a header of
                # '#replace1' looks for the literal '#replace1' in the doc.
                $replacements = @{}
                foreach ($prop in $row.PSObject.Properties) {

                    if ($prop.Name -ieq 'title')             { continue }
                    if ([string]::IsNullOrEmpty($prop.Name)) { continue }

                    if ($prop.Name -notmatch '^#') {
                        Write-Warning "Sheet '$sheet' row ${rowNum}: header '$($prop.Name)' has no leading '#'; using header text as the placeholder verbatim."
                    }

                    $val = if ($null -eq $prop.Value) { '' } else { [string]$prop.Value }
                    $replacements[$prop.Name] = $val
                }


                # Decide the output filename.
                # First chance goes to the customisation hook; if it declines
                # ($null) we fall back to title -> sheet -> sheet_<n>.

                $customName = Get-CustomOutputFileName -Row $row -SheetName $sheet -RowNumber $rowNum
                $customName = Get-SafeFileName ([string]$customName)

                if ($customName) {
                    $baseName = $customName
                }
                else {
                    $titleVal  = $null
                    $titleProp = $row.PSObject.Properties |
                                 Where-Object { $_.Name -ieq 'title' } |
                                 Select-Object -First 1
                    if ($titleProp) {
                        $titleVal = Get-SafeFileName ([string]$titleProp.Value)
                    }

                    if     ($titleVal)        { $baseName = $titleVal }
                    elseif ($rows.Count -eq 1){ $baseName = $sheet }
                    else                      { $baseName = "${sheet}_${rowNum}" }
                }

                # Disambiguate if a previous row produced the same name.
                $candidate = "$baseName.docx"
                $n = 2
                while ($usedNames.ContainsKey($candidate.ToLowerInvariant())) {
                    $candidate = "${baseName}_${n}.docx"
                    $n++
                }
                $usedNames[$candidate.ToLowerInvariant()] = $true
                $outPath = Join-Path $OutputDir $candidate

                Save-FilledDocument -Word $word `
                                    -TemplatePath $TemplatePath `
                                    -Replacements $replacements `
                                    -OutPath $outPath `
                                    -Row $row `
                                    -SheetName $sheet

                $results += [pscustomobject]@{
                    Sheet  = $sheet
                    Row    = $rowNum
                    Output = $outPath
                }
                Write-Host "Wrote $outPath"
            }
        }
    }
}
finally {
    # Always tear Word down -- otherwise an orphan WINWORD.EXE will linger.
    if ($word) {
        $word.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

# Emit one summary object per file produced so callers can pipe results.
$results
