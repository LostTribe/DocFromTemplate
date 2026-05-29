# DocFromTemplate

DocFromTemplate is a PowerShell tool that fills a Word `.docx` template by replacing placeholder strings. Values can come from rows in an Excel workbook (one document per row) or be supplied directly via a hashtable (one document — handy for calling from other scripts).

For each row in each worksheet, every placeholder in the template (e.g. `#replace1`, `#replace2`) is replaced with the value from the column whose header matches the placeholder text exactly. One filled `.docx` is produced per row.

Native support for CSV, JSON, txt, key=value, and direct-`-Values` modes — they all can be used with PowerShell.

## Requirements

- Windows with Microsoft Word installed (Word COM automation is used)
- PowerShell 5.1 or 7+

### Optional: Excel mode

Only if you want to use `-ExcelPath`, install the [`ImportExcel`](https://www.powershellgallery.com/packages/ImportExcel) module:

```powershell
Install-Module ImportExcel -Scope CurrentUser
```

The script imports it lazily — it's loaded only when `-ExcelPath` is supplied, so a clone that never touches Excel mode will never see a "module not installed" error.

## Quick start

Two ready-to-run examples live under `ExamplesWithDummyData\`. Each folder
contains a workbook (`data.xlsx`), a Word template, and a `Read me.md`
with full usage notes. Run the new-joiner example (HR welcome letter for
new employees):

```powershell
.\New-DocFromTemplate.ps1 `
    -ExcelPath    ".\ExamplesWithDummyData\New Joiner example\data.xlsx" `
    -TemplatePath ".\ExamplesWithDummyData\New Joiner example\template.docx"
```

Outputs land in `.\ExamplesWithDummyData\New Joiner example\output\`.

For a richer, formatted Change Request template (three sheets of realistic
IT changes):

```powershell
.\New-DocFromTemplate.ps1 `
    -ExcelPath    ".\ExamplesWithDummyData\Change Request example\data.xlsx" `
    -TemplatePath ".\ExamplesWithDummyData\Change Request example\cr-template.docx"
```

See [`ExamplesWithDummyData/New Joiner example/Read me.md`](ExamplesWithDummyData/New%20Joiner%20example/Read%20me.md) and [`ExamplesWithDummyData/Change Request example/Read me.md`](ExamplesWithDummyData/Change%20Request%20example/Read%20me.md) for more options.

## Usage

The script has five parameter sets covering tabular and single-document sources. Use whichever fits your data:

| Source | Parameter set | Shape | One doc or many? |
|---|---|---|---|
| Excel workbook | `FromExcel` (default) | `-ExcelPath` | one per row, all worksheets |
| CSV file | `FromCsv` | `-CsvPath` | one per row, flat file (no sheets) |
| JSON file (array) | `FromJson` | `-JsonPath` | one per array element |
| JSON file (object) | `FromJson` | `-JsonPath` | single document |
| Key=value text file | `FromKeyValue` | `-KeyValuePath` | single document |
| Hashtable in memory | `FromValues` | `-Values` | single document |

`-TemplatePath` and `-OutputDir` are shared across all modes. `-OutputDir` defaults to an `output\` folder beside the template.

All the examples below run against the bundled New Joiner data — copy, paste, and they work. Substitute your own paths once you've seen them produce output.

### Tabular modes (one document per row)

```powershell
# Excel — worksheets become groups; ImportExcel module required
.\New-DocFromTemplate.ps1 `
    -ExcelPath    ".\ExamplesWithDummyData\New Joiner example\data.xlsx" `
    -TemplatePath ".\ExamplesWithDummyData\New Joiner example\template.docx"

# CSV — flat, header row defines placeholders
.\New-DocFromTemplate.ps1 `
    -CsvPath      ".\ExamplesWithDummyData\New Joiner example\data.csv" `
    -TemplatePath ".\ExamplesWithDummyData\New Joiner example\template.docx"

# JSON array — each element is a row
.\New-DocFromTemplate.ps1 `
    -JsonPath     ".\ExamplesWithDummyData\New Joiner example\data.json" `
    -TemplatePath ".\ExamplesWithDummyData\New Joiner example\template.docx"
```

All three produce the same trio of files (`Jane Doe - Welcome.docx`, `John Smith - Welcome.docx`, `Alex Roe - Welcome.docx`) in the example's `output\` folder.

In every tabular mode, a column / property called `title` (case-insensitive) is used as the output filename. Without one, files are named after the source (sheet name / CSV file / JSON file) with a numeric suffix per row.

The Excel data has two worksheets (`engineering`, `sales`); use `-Worksheet 'sales'` to process only one.

### Single-document modes

```powershell
# Key=value text — one file -> one document. Bundled file has Jane Doe (xlsx row 1).
.\New-DocFromTemplate.ps1 `
    -KeyValuePath ".\ExamplesWithDummyData\New Joiner example\data.txt" `
    -TemplatePath ".\ExamplesWithDummyData\New Joiner example\template.docx" `
    -OutputName   'Jane Doe - Welcome (from txt)'

# Hashtable in memory — the original direct mode, no file required
.\New-DocFromTemplate.ps1 `
    -TemplatePath ".\ExamplesWithDummyData\New Joiner example\template.docx" `
    -Values @{
        '#replace1'  = 'Riley Doe'
        '#replace2'  = 'Designer'
        '#replace3'  = '2026-09-01'
        '#replace4'  = 'Design'
        '#replace5'  = 'Pat Lead'
        '#replace6'  = 'pat.lead@acmecorp.com'
        '#replace7'  = 'EMP-2026-9001'
        '#replace8'  = 'London HQ, Floor 2, Desk D-04'
        '#replace9'  = 'Chris People'
        '#replace10' = 'chris.people@acmecorp.com'
    } `
    -OutputName 'Riley Doe - Welcome (from values)'

# JSON object — same shape as a single JSON record on disk
# (No bundled single-object example; the bundled data.json is an array.
#  To try this mode, save one of the array elements as its own .json
#  file and point -JsonPath at it.)
```

`-OutputName` is optional. Defaults: `<template-base>-filled` (FromValues), `<json-base>-filled` (FromJson object), `<kv-base>-filled` (FromKeyValue).

**Gotchas to know:**
- `-Values` keys starting with `#` **must be quoted** (`'#name' = ...`); unquoted `#name` is a PowerShell comment.
- Key=value files use `;` as the comment marker (not `#`, which is reserved for placeholder names). Blank lines and `;`-prefixed lines are ignored. Each non-comment line is split on the **first** `=`.
- Row-level customisation hooks (`Invoke-RowPreProcess`, `Get-CustomOutputFileName`) do not fire in single-document modes. `Invoke-DocPostProcess` still does.

## Calling from other scripts

The direct `-Values` mode is built for scripted callers. Some common shapes:

### 1. Build the values hashtable separately

```powershell
$values = @{
    '#name'  = $personName
    '#role'  = $personRole
    '#start' = (Get-Date).AddDays(30).ToString('yyyy-MM-dd')
}

.\New-DocFromTemplate.ps1 `
    -TemplatePath .\letter.docx `
    -Values       $values `
    -OutputName   $personName
```

### 2. Splat the whole parameter list

`@params` (with `$params` a hashtable) is PowerShell's splatting syntax — keys map to named parameters. Different from `-Values @{...}`, which just passes one hashtable as a parameter value.

```powershell
$params = @{
    TemplatePath = '.\letter.docx'
    Values       = @{
        '#name' = 'Jane Doe'
        '#role' = 'Engineer'
    }
    OutputName   = 'Jane-welcome'
    OutputDir    = '.\out'
}

.\New-DocFromTemplate.ps1 @params
```

### 3. Mix splat + explicit — pin the frame, vary the data

```powershell
$base = @{
    TemplatePath = '.\letter.docx'
    OutputDir    = '.\out'
}

.\New-DocFromTemplate.ps1 @base -Values @{ '#name'='Alice'; '#role'='PM'  } -OutputName 'Alice'
.\New-DocFromTemplate.ps1 @base -Values @{ '#name'='Bob';   '#role'='Eng' } -OutputName 'Bob'
```

### 4. Driven from another data source

```powershell
Get-ADUser -Filter 'Department -eq "Sales"' -Properties EmailAddress, Title |
    ForEach-Object {
        .\New-DocFromTemplate.ps1 `
            -TemplatePath .\welcome.docx `
            -Values @{
                '#name'  = $_.Name
                '#email' = $_.EmailAddress
                '#role'  = $_.Title
            } `
            -OutputName $_.SamAccountName
    }
```

## How it works

- Column headers / object property names / hashtable keys are used as placeholder text **verbatim** — so `#replace1` replaces every occurrence of the literal string `#replace1` in the document. Choose any naming convention you like; the leading `#` is just a convention to make placeholders unlikely to collide with normal prose.
- In tabular modes (Excel, CSV, JSON array), a `title` column (case-insensitive) is treated specially: its value is used as the output filename instead of being a placeholder. Invalid filename characters are stripped.
- Tabular filename fallback (no `title` column): a single-row source produces `<sheet>.docx`; a multi-row source produces `<sheet>_1.docx`, `<sheet>_2.docx`, etc. ("sheet" = worksheet name in Excel, file basename in CSV/JSON.)
- Single-document filename: `-OutputName` if provided, otherwise `<source-base>-filled.docx`.
- The template file is opened read-only and never modified.
- Replacement runs across all story ranges (body, headers, footers, footnotes).
- Replacement is case-sensitive and uses substring matching, with placeholders processed in length-descending order so `#replace1` cannot eat the prefix of `#replace10`.
