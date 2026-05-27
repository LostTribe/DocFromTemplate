# DocFromTemplate

PowerShell tool that fills a Word `.docx` template using rows from an Excel `.xlsx` workbook.

For each row in each worksheet, every placeholder in the template (e.g. `#replace1`, `#replace2`) is replaced with the value from the column whose header matches the placeholder text exactly. One filled `.docx` is produced per row.

## Requirements

- Windows with Microsoft Word installed (Word COM automation is used)
- PowerShell 5.1 or 7+
- [`ImportExcel`](https://www.powershellgallery.com/packages/ImportExcel) module: `Install-Module ImportExcel -Scope CurrentUser`

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

```powershell
.\New-DocFromTemplate.ps1 `
    -ExcelPath    <path-to.xlsx> `
    -TemplatePath <path-to-template.docx> `
    [-OutputDir   <output-folder>] `
    [-Worksheet   <name1>,<name2>]
```

- `-OutputDir` defaults to an `output\` folder beside the template.
- `-Worksheet` filters which sheets to process; default is all sheets.

## How it works

- Column headers are used as placeholder text **verbatim** — so a header named `#replace1` replaces every occurrence of the literal string `#replace1` in the document. Choose any naming convention you like; the leading `#` is just a convention to make placeholders unlikely to collide with normal prose.
- A `title` column (case-insensitive) is treated specially: its value is used as the output filename instead of being a placeholder. Invalid filename characters are stripped.
- If no `title` column is present:
  - A worksheet with one data row produces `<worksheet>.docx`.
  - A worksheet with multiple data rows produces `<worksheet>_1.docx`, `<worksheet>_2.docx`, etc.
- The template file is opened read-only and never modified.
- Replacement runs across all story ranges (body, headers, footers, footnotes).
