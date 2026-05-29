New Joiner example - dummy data for New-DocFromTemplate.ps1
============================================================

What's in this folder
---------------------
  data.csv        Three new joiners in flat CSV form. Header row
                  defines the placeholders; each data row produces
                  one filled .docx.

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

  data.json       Same three joiners as a JSON array. Property
                  names match the column headers above. Each
                  object becomes one filled .docx.

  data.txt        Key=value format ('.env'-style), ONE record only -
                  Jane Doe, the first record from data.csv.
                  Demonstrates the single-document input mode.
                  Comments use ';' because '#' is reserved for
                  placeholder names.

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
Open PowerShell at the repository root (..\..). All three data files
produce the same Welcome.docx files in this folder's "output"
subdirectory; pick whichever input matches where your data lives.


From the CSV:

    .\New-DocFromTemplate.ps1 `
        -CsvPath      ".\ExamplesWithDummyData\New Joiner example\data.csv" `
        -TemplatePath ".\ExamplesWithDummyData\New Joiner example\template.docx"


From the JSON array:

    .\New-DocFromTemplate.ps1 `
        -JsonPath     ".\ExamplesWithDummyData\New Joiner example\data.json" `
        -TemplatePath ".\ExamplesWithDummyData\New Joiner example\template.docx"


From the key=value text file (single document):

    .\New-DocFromTemplate.ps1 `
        -KeyValuePath ".\ExamplesWithDummyData\New Joiner example\data.txt" `
        -TemplatePath ".\ExamplesWithDummyData\New Joiner example\template.docx" `
        -OutputName   'Jane Doe - Welcome'


CSV and JSON both produce, in the output folder:

    Jane Doe - Welcome.docx
    John Smith - Welcome.docx
    Alex Roe - Welcome.docx

(The key=value run produces only Jane Doe's letter, since that mode
takes a single record per file.)


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

  No external modules. Every mode uses built-in PowerShell.


Naming rules (reminder)
-----------------------
  * If a row has a "title" column (case-insensitive), its value becomes
    the output filename.
  * Otherwise a single-row source produces  <source>.docx,
            and a multi-row source produces  <source>_1.docx,
            <source>_2.docx, ... (where <source> is the CSV or JSON
            file basename)
  * Duplicates within a single run get a numeric suffix to disambiguate.

The template is opened read-only and is never modified.
