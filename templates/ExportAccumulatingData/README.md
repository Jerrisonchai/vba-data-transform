# Data Conversion – Export Utilities

## Overview
This module provides several export utilities to convert Excel data into **text files**, **CSV files**, or **new workbooks**.  
It is designed for daily report generation and archiving, with automated file naming and optional formatting.

---

## Features
- **Export Excel → Text (.txt)**  
  Converts data from the `Summary 2023` sheet into a tab-delimited `.txt` file.
- **Export Excel → CSV (.csv)**  
  Similar to the TXT export but saves in `.csv` format for downstream systems.
- **Copy Sheets to New Workbook**  
  Copies `Monthly` and `Summary 2023` sheets into a new workbook, retaining formats.
- **Save As with Custom Path**  
  Allows the user to choose a folder and automatically:
  - Saves the file with timestamped name
  - Updates linked formulas
  - Deletes unused sheets (`Dashboard`, `LOG`) for a clean output

---

## Workflow

### 1. Excel to Text
- Reads `Summary 2023` data
- Builds a tab-delimited `.txt` file
- Saves to folder specified in `Dashboard!C20`

### 2. Excel to CSV
- Identical to TXT export but with `.csv` extension
- Data is tab-delimited (public users may change delimiter)

### 3. Copy Sheets to New Workbook
- Creates a new workbook
- Copies `Monthly` and `Summary 2023` sheets
- Applies formatting from original sheets
- Saves to `Dashboard!C20`

### 4. Save As (with Folder Picker)
- Lets user choose save folder
- Saves with timestamped name
- Updates all Excel links
- Deletes extra sheets (`Dashboard`, `LOG`) before final save

---

## File Naming Convention
All exports follow the pattern:
```
Daily Sample List YYYY-MM-DD hh mm AMPM.xlsx
DailyData YYYY-MM-DD hh mm AMPM.txt
DailyData YYYY-MM-DD hh mm AMPM.csv
```


---

## Customization Guidelines

Public users can **adapt this tool** by editing:
- **Source Sheets**  
  Change `"Summary 2023"` or `"Monthly"` to your own sheet names.
- **Save Path**  
  Controlled via `Dashboard!C20`. Replace this with a fixed path if you prefer hardcoding:
  ```vba
  FileName = "C:\Exports\DailyData " & Format(Now(), "YYYY-MM-DD hh mm AMPM") & ".txt"
  ```
- Delimiter
  - Change:
```
Deliminator = vbTab
```
  - to "," (comma) or ";" (semicolon) for CSV variants.

- Sheet Cleanup (Save As routine)
  - Modify or remove:
```vba
Sheets("Dashboard").Delete
Sheets("LOG").Delete
```
  - if your workbook uses different admin sheets.

## Example Use Cases
- Daily operational exports for ERP imports (CSV/TXT).
- Automated generation of client-ready reports with formatting.
- Archiving monthly and yearly data into standalone files.

## Key Procedures
| Procedure                 | Purpose                                  |
| ------------------------- | ---------------------------------------- |
| `ExceltoText`             | Export Summary sheet → `.txt`            |
| `ExceltoCsv`              | Export Summary sheet → `.csv`            |
| `CopySheetsToNewWorkbook` | Copy Monthly & Summary → new workbook    |
| `SaveAs`                  | Save cleaned workbook with updated links |

## Notes
- Relies on helper macros: capturetime, captureendtime, MyShape_Click, and OptimizedMode.
- Ensure these are present in your project or remove if not needed.
- Works best when Dashboard!C20 points to a valid folder path.
- File operations overwrite existing files with the same name.

## Status
- ✅ Tested for TXT/CSV export up to 50k rows
- ⚠️ Very large datasets (>200k rows) may take noticeable time
- ⚠️ If Dashboard!C20 is blank, export will fail

