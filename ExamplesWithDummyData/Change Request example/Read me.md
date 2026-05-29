Change Request example - dummy data for New-DocFromTemplate.ps1
===============================================================

What's in this folder
---------------------
  data.csv           Twelve realistic IT Change Requests in flat CSV
                     form. Header row defines the placeholders; each
                     data row produces one filled .docx.

                     Names are stock placeholders (A. Doe / B. Smith
                     etc.) and vendor / product references are
                     genericised ("datacentre core switch", "CRM
                     platform major release") so the data is
                     obviously fake.

                     Shared headers (legend for the numbered
                     placeholders):
                       title       -> output filename (NOT a placeholder)
                       #Replace1   -> CR Reference
                       #Replace2   -> Date submitted
                       #Replace3   -> Requester
                       #Replace4   -> Unit / department
                       #Replace5   -> Priority
                       #Replace6   -> Change type
                       #Replace7   -> Summary of requested change
                       #Replace8   -> Business justification
                       #Replace9   -> Proposed change / implementation plan
                       #Replace10  -> Risk & impact assessment
                       #Replace11  -> Approver / current decision
                       #Replace12  -> Implementation window
                       #Replace13  -> Status
                       #Replace14  -> Backout / rollback plan
                       #Replace15  -> Stakeholders / affected parties

  data.json          Same twelve CRs as a JSON array of objects.
                     Property names match the column headers above.
                     Each object becomes one filled .docx.

  data.txt           Key=value format, ONE record only - CR-2026-0142
                     (the Core Switch Firmware Upgrade, first record
                     of data.csv). Demonstrates the single-document
                     input mode. Comments use ';' because '#' is
                     reserved for placeholder names.

  cr-template.docx   A formatted Change Request form: page header with CR
                     reference, navy title bar, 3-column metadata grid,
                     requester table, accent-blue section headings,
                     red/amber/green callouts for risk, backout and
                     approval, signature block, and a footer rule.

  output\            Default folder the merge script writes filled .docx
                     files into (created on first run if missing).


How to run
----------
Open PowerShell at the repository root (..\..). Pick the input format
that matches where your data lives - all three produce filled .docx
files in this folder's "output" subdirectory.


From the CSV:

    .\New-DocFromTemplate.ps1 `
        -CsvPath      ".\ExamplesWithDummyData\Change Request example\data.csv" `
        -TemplatePath ".\ExamplesWithDummyData\Change Request example\cr-template.docx"


From the JSON array:

    .\New-DocFromTemplate.ps1 `
        -JsonPath     ".\ExamplesWithDummyData\Change Request example\data.json" `
        -TemplatePath ".\ExamplesWithDummyData\Change Request example\cr-template.docx"


From the key=value text file (single document):

    .\New-DocFromTemplate.ps1 `
        -KeyValuePath ".\ExamplesWithDummyData\Change Request example\data.txt" `
        -TemplatePath ".\ExamplesWithDummyData\Change Request example\cr-template.docx" `
        -OutputName   'CR-2026-0142-Core-Switch-Firmware-Upgrade'


CSV and JSON each produce twelve files - one CR per row - named after
each row's "title" column, e.g.:

    CR-2026-0142-Core-Switch-Firmware-Upgrade.docx
    CR-2026-0228-CRM-Major-Release.docx
    CR-2026-0314-Emergency-Critical-Vuln-Patch.docx
    ...

The key=value run produces only the one CR (data.txt holds CR-2026-0142
only).


Send output somewhere else (omit -OutputDir to use the default
"output\" folder beside the template):

    .\New-DocFromTemplate.ps1 `
        -CsvPath      ".\ExamplesWithDummyData\Change Request example\data.csv" `
        -TemplatePath ".\ExamplesWithDummyData\Change Request example\cr-template.docx" `
        -OutputDir    .\out


Requirements
------------
  * Windows with Microsoft Word installed (Word COM automation is used)
  * PowerShell 5.1 or 7+

  No external modules. Every mode uses built-in PowerShell.


Notes
-----
  * The CR template uses placeholders in the page header too (#Replace1).
    New-DocFromTemplate.ps1 replaces across all story ranges (body,
    headers, footers, footnotes), so the CR reference appears on every
    page of the rendered document automatically.
  * The template is opened read-only and is never modified.
  * Replacement is case-sensitive: the data headers use #Replace1
    (capital R) on purpose - the template uses the same casing.
