# DataSegregator (UserForm Tool)

## Overview
The **DataSegregator** utility provides a flexible UserForm-driven interface to split and export data from the `Data` sheet into new workbooks or worksheets.  
It enables business users to:
- Select fields for segregation
- Choose export formats (one file, multiple files, or multiple sheets)
- Save outputs into organized folders with optional subfolder creation

---

## Features
- **Interactive UserForm**:
  - ListBox1 / ListBox2: select fields to segregate on
  - ComboBox1: set primary field (for multi-level splits)
  - OptionButtons: choose export strategy
  - CheckBox1: enable/disable file saving
  - CheckBox2: optional subfolder creation
  - TextBox2: pick save folder
- **3 Export Modes**:
  1. **Multiple Sheets in One File**  
     Splits data into multiple worksheets within a single Excel file.
  2. **Multiple Files with One Sheet Each**  
     Creates individual Excel files, each containing one filtered dataset.
  3. **Multiple Files with Multiple Sheets**  
     Uses a primary field (selected via ComboBox) to generate one workbook per unique primary value, each containing sub-sheets for other splits.
- **Dynamic Field Selection**: Automatically pulls field names from the first row of the `Data` sheet.
- **Validation & Error Handling**: Prevents invalid folder paths or missing selections.
- **Optional Saving**: Run segregation without saving, or enable saving to a chosen folder.

---

## Workflow
1. Load the **UserForm**.
2. The field list (row 1 from `Data` sheet) populates into **ListBox1**.
3. Move fields between `ListBox1` and `ListBox2` using buttons:
   - ➡️ Move selected field to `ListBox2` (segregation fields)
   - ⬅️ Move selected field back to `ListBox1`
4. Choose an export option:
   - **OptionButton1** → Multiple sheets in one file
   - **OptionButton2** → Multiple files with one sheet
   - **OptionButton3** → Multiple files with multiple sheets (requires primary field)
5. (Optional) Enable `CheckBox1` to save results, and use `CommandButton5` to pick a folder.
   - If **Option 2** is selected, you can also enable `CheckBox2` to create subfolders.
6. Click **Run** (`CommandButton6`) to execute.
7. Data is filtered and exported as per your selections.

---

## Example Use Cases
- Splitting sales data by **Region** and saving each into separate files.  
- Creating a single workbook with multiple sheets, one per **Product Category**.  
- Generating one workbook per **Customer**, with sheets split by **Order Status**.  

---

## Key Procedures

- **Left_List_Box / Right_List_Box**  
  Populate ListBoxes from `Data!Row1`.

- **Primay_Field_List**  
  Enables/disables ComboBox and OptionButton3 depending on selected fields.

- **Multiple_Sheet_in_one_file**  
  Splits data by selected fields → one Excel file, multiple sheets.

- **Multiple_file_with_one_sheet**  
  Creates multiple Excel files, each containing a single filtered dataset.

- **Multiple_file_with_multiple_sheet**  
  Uses a primary field and generates one file per primary value, with multiple sheets inside.

---

## UI Elements Mapping
| Control        | Purpose |
|----------------|----------|
| **CheckBox1**  | Enable file saving |
| **TextBox2**   | Save folder path |
| **CommandButton5** | Pick folder (via FileDialog) |
| **ListBox1**   | Available fields |
| **ListBox2**   | Selected fields for segregation |
| **ComboBox1**  | Choose primary field |
| **OptionButton1** | Export: multiple sheets in one file |
| **OptionButton2** | Export: multiple files (one sheet each) |
| **OptionButton3** | Export: multiple files (multiple sheets) |
| **CheckBox2**  | Enable subfolders (only for Option 2) |
| **CommandButton6** | Run segregation |
| **CommandButton7** | Reset form |

---

## Notes
- Requires a `Data` sheet (with headers in row 1).
- Uses a `Support` sheet as a temporary helper for unique values.
- Large datasets may take time due to repeated filtering/copying.
- Saved files are in `.xlsx` format.

---

## Status
✅ Tested with multi-column segregation scenarios.  
⚠️ Ensure `Data` and `Support` sheets exist before running.
