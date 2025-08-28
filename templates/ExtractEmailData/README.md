# Outlook Email Integration (Windows Only)

## Overview
This module connects Excel with **Outlook Classic (Windows Desktop app only)** to pull emails into a worksheet for analysis.  
It allows you to:
1. Import email headers (Sender, Subject, Received Date).
2. Drill down into the **email body** of selected messages.

⚠️ This tool is **not compatible with Outlook on the web (Outlook 365 online)**.  
It requires the classic desktop version of Outlook installed on Windows.

---

## Setup Instructions

### 1. Enable Outlook Reference in VBA
Before running the code, you must add a reference to the **Microsoft Outlook Object Library**:

1. Open Excel → `Alt + F11` (VBA Editor).
2. Go to `Tools > References`.
3. Check ✅ **Microsoft Outlook XX.X Object Library**  
   (Version number depends on your Office installation).

### 2. Configure Target Folder
- This script reads emails from:
  - The **Inbox** by default, then
  - A **custom subfolder**, defined in `Dashboard!C16`.
- Example: If your Outlook Inbox contains a subfolder named **"Reports"**, type:
```Reports``` in cell `Dashboard!C16`.

### 3. Prepare Worksheets
- **Dashboard sheet**:  
- Cell `C16` → subfolder name (e.g., `"Reports"`).
- **Data sheet**:  
- Emails will be listed starting from cell `A2`.

---

## Workflow

### Step 1 – Get Inbox Items
Runs through the target folder and extracts:

| Column | Data Extracted |
|--------|----------------|
| A      | Sender Name    |
| B      | Subject        |
| C      | Received Time  |

Clears any old data in the `Data` sheet before writing.

```vba
Call GetInboxItems
```

## Customization
Public users may modify:

### 1. Change Source Folder
- Replace:
```
Set fol = fol.Folders(Sheets("Dashboard").Range("C16").Value)
```
with a hard-coded folder name if desired:
```
Set fol = ns.GetDefaultFolder(olFolderInbox).Folders("Reports")
```
### 2. Extract More Email Fields
- Add more columns, e.g.:
```
Sheets("Data").Cells(n, 4).Value = mi.To
Sheets("Data").Cells(n, 5).Value = mi.CC
Sheets("Data").Cells(n, 6).Value = mi.Body
```
### 3. Filter by Date Range
- Limit retrieval by checking:
```
If mi.ReceivedTime >= Date - 7 Then
    ' Only last 7 days
End If
```
### 4. Change Email Body Handling
- Currently the body is copied as plain text.
- To paste as HTML or RTF, you may adapt the Word editor logic.

---

## Example Use Cases
- Pulling all client report emails from a subfolder into Excel.
- Extracting only specific subjects for reconciliation.
- Creating a quick audit trail of received communications.

## Error Handling
- If Outlook is not open, user is prompted: "Please open Outlook apps first!"
- If no sender is selected, process is aborted with a message.
- If no matching item is found, message: "Nothing was found".

## Key Procedures
| Procedure         | Purpose                                    |
| ----------------- | ------------------------------------------ |
| `GetInboxItems`   | Pulls email headers from subfolder → Excel |
| `GetEmailDetails` | Fetches full email body of selected row    |

## Status
- ✅ Tested with Outlook Classic (Office 2016 & 2019, Windows 10/11)
- ⚠️ Does not work with Outlook on the web or Mac
- ⚠️ Large folders (>10k emails) may load slowly

