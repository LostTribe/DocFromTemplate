New Joiner example - dummy data for New-DocFromTemplate.ps1
============================================================

What's in this folder
---------------------
  data.xlsx       Two worksheets, three new joiners in total:
                    engineering : Jane Doe, John Smith
                    sales       : Alex Roe

                  Scenario: HR / People team welcoming new employees
                  to "Acme Corp" on their first day.

                  Names are stock placeholders (Doe / Smith / Roe etc.)
                  so the data is obviously fake.

                  Columns (legend for the numbered placeholders):
                    title       -> used as the output filename
                    #replace1   -> New employee's full name
                    #replace2   -> Job title / role
                    #replace3   -> Start date (first day)
                    #replace4   -> Team / department
                    #replace5   -> Line manager name
                    #replace6   -> Line manager email
                    #replace7   -> Employee ID
                    #replace8   -> Office & desk location
                    #replace9   -> HR contact name
                    #replace10  -> HR contact email

  data.csv        Same three joiners as data.xlsx, in flat CSV form. No
                  worksheets concept (CSV doesn't have them). Use this
                  if you don't want the ImportExcel module dependency.

  data.json       Same three joiners as data.xlsx, as a JSON array of
                  objects. Each object becomes one filled document.
                  Property names match the column headers above.

  data.txt        Key=value format ('.env'-style), one record only -
                  Jane Doe, the first row from the xlsx. Demonstrates
                  the single-document input mode. Comments use ';'
                  because '#' is reserved for placeholder names.

  template.docx   A formatted welcome letter that contains the literal
                  strings #replace1 ... #replace10 wherever the matching
                  values should be inserted. Replacement runs across all
                  story ranges (body, header, footer), so any placeholder
                  reused in the page footer is filled too.

                  The navy bar at the top of the letter says "Acme Corp"
                  as static text. To make this your own template, open
                  template.docx once and swap "Acme Corp" (in the title
                  bar and the page footer) for your organisation's name.

  output\         Default folder the merge script writes filled .docx
                  files into (created on first run if missing).


How to run
----------
Open PowerShell at the repository root (..\..). All four data files
produce the same three Welcome.docx files in this folder's "output"
subdirectory; pick whichever input matches where your data lives.


From the Excel workbook (requires the ImportExcel module):

    .\New-DocFromTemplate.ps1 `
        -ExcelPath    ".\ExamplesWithDummyData\New Joiner example\data.xlsx" `
        -TemplatePath ".\ExamplesWithDummyData\New Joiner example\template.docx"


From the CSV (no module required, built-in PowerShell):

    .\New-DocFromTemplate.ps1 `
        -CsvPath      ".\ExamplesWithDummyData\New Joiner example\data.csv" `
        -TemplatePath ".\ExamplesWithDummyData\New Joiner example\template.docx"


From the JSON array (no module required):

    .\New-DocFromTemplate.ps1 `
        -JsonPath     ".\ExamplesWithDummyData\New Joiner example\data.json" `
        -TemplatePath ".\ExamplesWithDummyData\New Joiner example\template.docx"


From the key=value text file (single document, no module required):

    .\New-DocFromTemplate.ps1 `
        -KeyValuePath ".\ExamplesWithDummyData\New Joiner example\data.txt" `
        -TemplatePath ".\ExamplesWithDummyData\New Joiner example\template.docx" `
        -OutputName   'Jane Doe - Welcome'


From all three rows above produce, in the output folder:

    Jane Doe - Welcome.docx
    John Smith - Welcome.docx
    Alex Roe - Welcome.docx

(The key=value run produces only Jane Doe's letter, since that mode
takes a single record per file.)


Only process one team (Excel only - CSV and JSON are flat):

    .\New-DocFromTemplate.ps1 `
        -ExcelPath    ".\ExamplesWithDummyData\New Joiner example\data.xlsx" `
        -TemplatePath ".\ExamplesWithDummyData\New Joiner example\template.docx" `
        -Worksheet    'engineering'


Send output somewhere else (omit -OutputDir to use the default
"output\" folder beside the template):

    .\New-DocFromTemplate.ps1 `
        -CsvPath      ".\ExamplesWithDummyData\New Joiner example\data.csv" `
        -TemplatePath ".\ExamplesWithDummyData\New Joiner example\template.docx" `
        -OutputDir    .\out


Requirements
------------
  * Windows with Microsoft Word installed (Word COM automation is used)
  * PowerShell 5.1 or 7+
  * The ImportExcel module - ONLY for the Excel mode (-ExcelPath).
    CSV, JSON, and key=value modes use built-in PowerShell:
        Install-Module ImportExcel -Scope CurrentUser


Naming rules (reminder)
-----------------------
  * If a row has a "title" column (case-insensitive), its value becomes
    the output filename.
  * Otherwise a single-row sheet produces  <sheet>.docx,
            and a multi-row sheet produces  <sheet>_1.docx, <sheet>_2.docx, ...
  * Duplicates within a single run get a numeric suffix to disambiguate.

The template is opened read-only and is never modified.
