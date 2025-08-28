# DataCount

## Overview
The **DataCount** utility automates the process of:
- Scanning multiple Excel files in a folder
- Counting rows of data
- Categorizing and tallying values based on specified columns
- Storing results in a summary sheet (`Checking`)
- Optionally copying headers into a central `Data` sheet

This is especially useful when consolidating KPIs or preparing reports across multiple files.

---

## Features
- Batch processing of all `.xlsx` files in a target folder.
- User-specified **category column** and **count column** for flexible analysis.
- Dynamic extraction of categories from the `Checking` sheet (row 1, starting from column C).
- Row counts and category tallies automatically summarized in the results area.
- Smart header copy: only done once into the `Data` sheet to standardize further transformations.
- Built-in validation (e.g., folder path, required columns).
- Optimized for performance using Excel’s `Application` settings.

---

## Workflow
1. The folder path is taken from `Dashboard!C20`.
2. The category column and count column are read from `Dashboard!C21` and `C22`.
3. The script loops through all `.xlsx` files in the folder.
4. For each file:
   - Count the total rows (excluding header).
   - Count categories defined in the `Checking` sheet.
   - Write results into `Checking` starting from row 2.
   - Optionally copy headers to the `Data` sheet (first time only).
5. Results are displayed, and a success message is shown.

---

## Example Output (Checking Sheet)

| Filename     | Total Rows | CatA | CatB | CatC | ... |
|--------------|-----------:|------|------|------|-----|
| File1.xlsx   |       120  |  45  |  30  |  45  |     |
| File2.xlsx   |       95   |  25  |  40  |  30  |     |
| ...          |      ...   | ...  | ...  | ...  |     |

---

## Key Functions

- **`GetCategories(wshT)`**  
  Reads category values from `Checking!Row1` (starting column C).

- **`CountCategories(wshS, wshT, catCol, categories, t)`**  
  Uses `COUNTIFS` to calculate frequency of each category.

- **`CopyHeaders(wshS, wshD)`**  
  Copies headers once into `Data` sheet if not al
