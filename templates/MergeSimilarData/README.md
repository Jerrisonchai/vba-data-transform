# Excel VBA – File & Folder Automation

## Overview
This module provides a set of macros for handling **Excel files stored in your PC folders**.  
You can use it to:
- Select a folder and extract data from multiple `.xlsx` files.
- Perform row counts and validations.
- Export specific sheets into a new workbook.
- Save reports automatically into a chosen folder.

These tools are designed for **Windows Excel Desktop only** (not Excel Online / Mac).

---

## Setup Instructions

### 1. Enable Macros
Ensure macros are enabled in your Excel settings:
- Go to `File > Options > Trust Center > Trust Center Settings > Macro Settings`.
- Select **Enable all macros** (or set a trusted folder).

### 2. Dashboard Setup
The module uses the **Dashboard sheet** to capture folder paths:
- `C18` → Input/output folder for data extraction.
- `C20` → Export folder for saving new reports.

⚠️ If these cells are empty, the folder picker dialog will appear.

### 3. File Format
Macros only work with **Excel `.xlsx` or `.xlsm` files**.  
Other file types will be ignored.

---

## Main Procedures

### 1. `CopyData2`
- Lets you choose a folder containing Excel files.
- Extracts data from the first sheet of each file.
- Appends all rows into the **Data sheet**.
- Adds a `FileName` column for traceability.
- Cleans blanks and refreshes dashboard status.

👉 Customization:
- Change which sheet is used:
  ```vba
  Set wshS = wbkS.Worksheets(1)   ' Use first sheet
  ```
- Replace 1 with a sheet index or name, e.g.:
```vba
  Set wshS = wbkS.Sheets("Report")
```
- Restrict files to a certain naming convention:
```vba
sFile = Dir(sFolder & "Sales_*.xlsx")
```
### 2. CopyData3
- Similar to CopyData2, but results go into the Checking sheet.
- Records:
  - FileName
  - Data Count (number of rows in each file)
- Auto-generates formulas for:
  - Countif → number of records per file (cross-check with Data sheet).
  - Sumif → total amounts per file.

👉 Customization:
- Adjust minimum rows to include:
```vba
If m < 11 Then
    ' Skip file
End If
```
Change 11 to a higher/lower threshold.

### 3. CopySheetsToNewWorkbook
- Copies two sheets (Monthly and Summary 2023) from your source workbook.
- Creates a new workbook and pastes them in.
- Preserves formatting from the originals.
- Saves into the folder specified in Dashboard!C20.
👉 Customization:
- Add more sheets:
```vba
Set sourceSheet3 = sourceWorkbook.Sheets("ExtraSheet")
sourceSheet3.Copy Before:=newWorkbook.Sheets(1)
```
- Change default sheet names by editing:
```vba
Set sourceSheet1 = sourceWorkbook.Sheets("Monthly")
Set sourceSheet2 = sourceWorkbook.Sheets("Summary 2023")
```
### 4. SaveAs
- Saves the current workbook into the folder from Dashboard!C20.
- Renames the file automatically:
```
Daily Sample List YYYY-MM-DD hh mm AMPM.xlsx
```
- Updates external links to the new file.
- Deletes unused sheets (Dashboard and LOG) in the saved version.
👉 Customization:
- Change the file naming pattern:
```vba
"Daily Report " & Format(Now(), "YYYYMMDD_hhmm") & ".xlsx"
```
- Prevent sheet deletion by commenting out:
```vba
'Sheets("Dashboard").Delete
'Sheets("LOG").Delete
```
## Customization Tips
### 1. Default Folder Path
- Pre-fill Dashboard!C18 or Dashboard!C20 with your default folder path.
- The picker will still let you override it if needed.
### 2. File Filters
- Modify:
```vba
sFile = Dir(sFolder & "*.xlsx*")
```
to target other file types:
- .xlsm → *.xlsm
- .csv → *.csv
### 3. Performance Tuning
- Wrap loops with OptimizedMode True/False (already included) to speed up large batch processing.
- Avoid opening very large workbooks unnecessarily.

## Example Use Cases
- Consolidating weekly sales reports from multiple files into one dataset.
- Checking row counts across branch reports for auditing.
- Exporting only summary tabs to share with stakeholders.
- Automating daily save/export workflows.

## Error Handling
- If no folder is selected → process ends.
- If a file has fewer than the required rows (threshold), it is skipped.
- Blanks in consolidated data are cleaned automatically.
- Success status, start time, and username are logged.

## Status
- ✅ Tested with Windows Excel 2016, 2019, 2021
- ⚠️ May break if files contain protected sheets or unsupported formats
- ⚠️ Avoid running with >1000 files at once (Excel memory limits)
