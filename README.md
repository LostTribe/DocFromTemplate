# DocFromTemplate

PowerShell tool that fills a Word `.docx` template by replacing placeholder strings. Values can come from rows in an Excel workbook (one document per row) or be supplied directly via a hashtable (one document — handy for calling from other scripts).

For each row in each worksheet, every placeholder in the template (e.g. `#replace1`, `#replace2`) is replaced with the value from the column whose header matches the placeholder text exactly. One filled `.docx` is produced per row.

## Requirements

- Windows with Microsoft Word installed (Word COM automation is used)
- PowerShell 5.1 or 7+
- [`ImportExcel`](https://www.powershellgallery.com/packages/ImportExcel) module: `Install-Module ImportExcel -Scope CurrentUser`
  - Only needed for the Excel mode. The direct `-Values` mode does not require it.

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

The script has two parameter sets. Use whichever fits.

### Excel mode (one document per row)

```powershell
.\New-DocFromTemplate.ps1 `
    -ExcelPath    <path-to.xlsx> `
    -TemplatePath <path-to-template.docx> `
    [-OutputDir   <output-folder>] `
    [-Worksheet   <name1>,<name2>]
```

- `-OutputDir` defaults to an `output\` folder beside the template.
- `-Worksheet` filters which sheets to process; default is all sheets.

### Direct values mode (single document, no workbook)

```powershell
.\New-DocFromTemplate.ps1 `
    -TemplatePath <path-to-template.docx> `
    -Values       <hashtable> `
    [-OutputName  <filename-without-extension>] `
    [-OutputDir   <output-folder>]
```

- `-Values` is a hashtable of `placeholder => value` pairs. **Keys starting with `#` must be quoted** (`'#name' = ...`); unquoted `#name` is a PowerShell comment.
- `-OutputName` defaults to `<template-base>-filled`.
- Row-level customisation hooks (`Invoke-RowPreProcess`, `Get-CustomOutputFileName`) do not fire in this mode. `Invoke-DocPostProcess` still does.

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

- Column headers (or `-Values` keys) are used as placeholder text **verbatim** — so `#replace1` replaces every occurrence of the literal string `#replace1` in the document. Choose any naming convention you like; the leading `#` is just a convention to make placeholders unlikely to collide with normal prose.
- In Excel mode, a `title` column (case-insensitive) is treated specially: its value is used as the output filename instead of being a placeholder. Invalid filename characters are stripped.
- Excel mode filename fallback (no `title` column):
  - A worksheet with one data row produces `<worksheet>.docx`.
  - A worksheet with multiple data rows produces `<worksheet>_1.docx`, `<worksheet>_2.docx`, etc.
- `-Values` mode filename: `-OutputName` if provided, otherwise `<template-base>-filled.docx`.
- The template file is opened read-only and never modified.
- Replacement runs across all story ranges (body, headers, footers, footnotes).
