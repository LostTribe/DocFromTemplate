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
Open PowerShell at the repository root (..\..) and run:

    .\New-DocFromTemplate.ps1 `
        -ExcelPath    ".\ExamplesWithDummyData\New Joiner example\data.xlsx" `
        -TemplatePath ".\ExamplesWithDummyData\New Joiner example\template.docx"

This produces three files in this folder's "output" subdirectory:

    Jane Doe - Welcome.docx
    John Smith - Welcome.docx
    Alex Roe - Welcome.docx


Only process one team:

    .\New-DocFromTemplate.ps1 `
        -ExcelPath    ".\ExamplesWithDummyData\New Joiner example\data.xlsx" `
        -TemplatePath ".\ExamplesWithDummyData\New Joiner example\template.docx" `
        -Worksheet    'engineering'


Send output somewhere else:

    .\New-DocFromTemplate.ps1 `
        -ExcelPath    ".\ExamplesWithDummyData\New Joiner example\data.xlsx" `
        -TemplatePath ".\ExamplesWithDummyData\New Joiner example\template.docx" `
        -OutputDir    'C:\Temp\welcomes'


Requirements
------------
  * Windows with Microsoft Word installed (Word COM automation is used)
  * PowerShell 5.1 or 7+
  * The ImportExcel module:
        Install-Module ImportExcel -Scope CurrentUser


Naming rules (reminder)
-----------------------
  * If a row has a "title" column (case-insensitive), its value becomes
    the output filename.
  * Otherwise a single-row sheet produces  <sheet>.docx,
            and a multi-row sheet produces  <sheet>_1.docx, <sheet>_2.docx, ...
  * Duplicates within a single run get a numeric suffix to disambiguate.

The template is opened read-only and is never modified.
