This Excel VBA project automates the process of importing HTML table data into Excel, cleaning it, and merging it into a master workbook.

## Features

- Browse HTML files from a specific folder (`html_files/`) and list them in a log sheet.
- Validate if HTML files exist and are not blank.
- Import HTML tables into a `RawDump` sheet with headers.
- Copy cleaned data to a `MasterData` sheet, avoiding repeated headers.
- Auto-format columns and left-align text for readability.
- Merge `MasterData` into a main workbook (`Main Workbook.xlsx`) with automatic row appending.
- Move processed HTML files to `Processed_Files/` with renamed filenames including creation date.

