
---

## 📄 **USER_GUIDE.md** (End User Guide)

# User Guide – vba-data-transform

This guide explains how to use the VBA utilities for transforming Excel data.

---

## 🖥️ Getting Started

1. Download the latest release from GitHub.  
2. Open `Excel` and press `Alt + F11` to access the VBA editor.  
3. Import the module files (`.bas` or `.cls`) from the feature folder.  
4. Save as a **Macro-Enabled Workbook** (`.xlsm`).  

---

## 📂 Features

### 1. Data Conversion (Read & Update)
- Reads data from sheets and applies structured updates.  
- Use when bulk-cleaning or modifying rows/columns.  

### 2. Data Count
- Counts records, rows, or specific column occurrences.  
- Useful for auditing dataset sizes.  

### 3. Data Segregation
- Splits data into multiple sheets/files based on conditions.  
- Example: split sales by region or category.  

### 4. Export Accumulating Data
- Collects and exports new data over time into a cumulative sheet/file.  
- Useful for monthly/weekly data pipelines.  

### 5. Extract Email Data
- Extracts emails from text/cells into a clean list.  
- Automatically removes duplicates.  

### 6. Merge Similar Data
- Identifies and merges similar rows across multiple sheets/files.  
- Reduces duplication in consolidated datasets.  

---

## 🚦 Running the Macros

1. Navigate to the **Dashboard sheet**.  
2. Select input/output folders (where applicable).  
3. Click the corresponding macro button or run from VBA Editor.  
4. Check the `Status`, `Start_Time`, `Time_Taken`, and `UserName` cells for logs.  

---

## 📑 Example

- Place your raw data in `/input/`.  
- Run **Data Segregation** → output files will be created in `/output/split/`.  
- Run **Merge Similar Data** → consolidated file will be available in `/output/merged/`.  

---

## ❓ FAQ

**Q: Do I need extra libraries?**  
A: No. All modules use native VBA + FileSystemObject (late binding).  

**Q: Can I undo transformations?**  
A: Always keep a backup. Modules overwrite target files.  

**Q: Where are results stored?**  
A: In the folder selected on the Dashboard or subfolder `/output/`.  

---
