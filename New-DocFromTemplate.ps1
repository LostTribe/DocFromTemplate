<#
.SYNOPSIS
    Fills a Word .docx template by replacing placeholder strings. Values
    can come from rows in an Excel workbook (one document per row) or be
    supplied directly via a hashtable (one document).

.DESCRIPTION
    Two parameter sets:

    FromExcel (default)
        For every worksheet in -ExcelPath the script reads each data row,
        opens the template, then performs a case-sensitive find & replace
        for each column header found in the row. Produces one document
        per row.

    FromValues
        Substitutes the placeholder/value pairs in -Values into the
        template once and writes a single document. Useful for one-off
        renders that do not warrant a workbook.

    Naming rules (FromExcel):
      * If a row has a 'title' column (case-insensitive) its value is the
        output filename.
      * Otherwise a single-row sheet produces <sheet>.docx, and a
        multi-row sheet produces <sheet>_1.docx, <sheet>_2.docx, ...
      * Duplicates within a run get a numeric suffix to disambiguate.

    Naming rules (FromValues):
      * Use -OutputName if provided, otherwise <template-base>-filled.docx.

    The template file is opened read-only and is never modified.

    Customising behaviour:
        Three customisation hooks are defined near the top of this script.
        They are no-ops by default:

            Invoke-RowPreProcess     -- mutate row values before replacement
            Get-CustomOutputFileName -- override the output filename rule
            Invoke-DocPostProcess    -- modify the open Word document
                                        after replacements but before save

        In FromValues mode, Invoke-RowPreProcess and Get-CustomOutputFileName
        are NOT called (you are already supplying values and the name
        directly). Invoke-DocPostProcess still fires; $Row is the values
        hashtable cast to a PSCustomObject and $SheetName is '(values)'.

.PARAMETER ExcelPath
    Path to the source .xlsx workbook. Required in FromExcel mode.

.PARAMETER TemplatePath
    Path to the Word .docx template. Required in both modes.

.PARAMETER OutputDir
    Folder to write filled documents into. Defaults to an 'output' folder
    next to the template. Created if it does not exist.

.PARAMETER Worksheet
    FromExcel only. Optional list of worksheet names. Defaults to every
    sheet in the workbook.

.PARAMETER Values
    FromValues only. Hashtable of placeholder => value pairs. Required
    in FromValues mode.

.PARAMETER OutputName
    FromValues only. Output filename (without extension). Defaults to
    <template-base>-filled.

.EXAMPLE
    .\New-DocFromTemplate.ps1 -ExcelPath .\samples\data.xlsx `
                              -TemplatePath .\samples\template.docx

.EXAMPLE
    .\New-DocFromTemplate.ps1 -ExcelPath .\data.xlsx `
                              -TemplatePath .\letter.docx `
                              -Worksheet 'contoso','tailspin' `
                              -OutputDir 'C:\merged'

.EXAMPLE
    .\New-DocFromTemplate.ps1 -TemplatePath .\letter.docx `
                              -Values @{ '#replace1' = 'Jane Doe'
                                         '#replace2' = 'Engineer' } `
                              -OutputName 'Jane-welcome'
#>

[CmdletBinding(DefaultParameterSetName = 'FromExcel')]
param(
    [Parameter(Mandatory, ParameterSetName = 'FromExcel')] [string]    $ExcelPath,
    [Parameter(Mandatory)]                                 [string]    $TemplatePath,
    [Parameter(ParameterSetName = 'FromExcel')]            [string[]]  $Worksheet,
    [Parameter(Mandatory, ParameterSetName = 'FromValues')][hashtable] $Values,
    [Parameter(ParameterSetName = 'FromValues')]           [string]    $OutputName,
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
      $Row       -- the original PSCustomObject from Import-Excel
      $SheetName -- the worksheet the row came from

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
    back to the built-in rule (title column, then <sheet>.docx /
    <sheet>_<n>.docx).

    Parameters:
      $Row       -- the (possibly pre-processed) row
      $SheetName -- the worksheet the row came from
      $RowNumber -- 1-based index within the sheet

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
      $SheetName -- the worksheet the row came from

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

if ($PSCmdlet.ParameterSetName -eq 'FromExcel') {
    if (-not (Test-Path $ExcelPath)) { throw "Excel file not found: $ExcelPath" }
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        throw "ImportExcel module not installed. Run: Install-Module ImportExcel -Scope CurrentUser"
    }
    Import-Module ImportExcel
}

# Word COM does not understand relative paths, so resolve everything to
# fully qualified paths up front.
$TemplatePath = (Resolve-Path $TemplatePath).Path
if ($PSCmdlet.ParameterSetName -eq 'FromExcel') {
    $ExcelPath = (Resolve-Path $ExcelPath).Path
}

if (-not $OutputDir) {
    $OutputDir = Join-Path (Split-Path $TemplatePath) 'output'
}
if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir | Out-Null
}
$OutputDir = (Resolve-Path $OutputDir).Path


# Worksheet selection (FromExcel only)
# Get-ExcelSheetInfo returns objects describing each sheet; we only need
# the names so they can be filtered against -Worksheet (when supplied).

$sheets = @()
if ($PSCmdlet.ParameterSetName -eq 'FromExcel') {
    $sheets = Get-ExcelSheetInfo -Path $ExcelPath | Select-Object -ExpandProperty Name
    if ($Worksheet) {
        $sheets = $sheets | Where-Object { $Worksheet -contains $_ }
    }
    if (-not $sheets) {
        throw "No matching worksheets found."
    }
}


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
    $OutPath. Shared by both the FromExcel and FromValues paths so the
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

    if ($PSCmdlet.ParameterSetName -eq 'FromValues') {

        # FromValues: single document, values supplied directly.
        # Row-level hooks are skipped; the user is in full control.
        # Invoke-DocPostProcess still fires inside Save-FilledDocument.

        $replacements = @{}
        foreach ($key in $Values.Keys) {
            if ([string]::IsNullOrEmpty($key)) { continue }
            $name = [string]$key
            if ($name -notmatch '^#') {
                Write-Warning "Placeholder '$name' has no leading '#'; using verbatim."
            }
            $val = if ($null -eq $Values[$key]) { '' } else { [string]$Values[$key] }
            $replacements[$name] = $val
        }

        if ($OutputName) {
            $baseName = Get-SafeFileName $OutputName
            if (-not $baseName) {
                throw "OutputName '$OutputName' sanitised to an empty string."
            }
        }
        else {
            $baseName = [IO.Path]::GetFileNameWithoutExtension($TemplatePath) + '-filled'
        }
        $outPath = Join-Path $OutputDir "$baseName.docx"

        # Synthesise a $Row for Invoke-DocPostProcess. Cast to PSCustomObject
        # so hook code that reads $Row.SomeProp behaves the same as in
        # FromExcel mode.
        $synthRow = [pscustomobject]$Values

        Save-FilledDocument -Word $word `
                            -TemplatePath $TemplatePath `
                            -Replacements $replacements `
                            -OutPath $outPath `
                            -Row $synthRow `
                            -SheetName '(values)'

        $results += [pscustomobject]@{ Sheet = '(values)'; Row = 1; Output = $outPath }
        Write-Host "Wrote $outPath"
    }
    else {

        # FromExcel: one document per row across every selected worksheet.
        foreach ($sheet in $sheets) {

            # Force array semantics so a one-row sheet still indexes cleanly
            # via $rows[0] rather than collapsing to a bare object.
            $rows = @(Import-Excel -Path $ExcelPath -WorksheetName $sheet)

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
