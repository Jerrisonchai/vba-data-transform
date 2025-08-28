# VBA Data Transform

This repository contains a collection of **VBA macros** designed to simplify and automate data transformation workflows in Excel.  
It focuses on common but time-consuming tasks such as converting, merging, segregating, and exporting data — enabling faster analysis and cleaner reporting.  

---

## 📂 Repository Structure

The repo is organized into **6 functional modules** (folders), each handling a specific data transformation use case:

1. **Data conversion - Read and update**  
   - Automates data conversion from one format/structure to another.  
   - Reads existing data, applies transformations, and updates records.  

2. **Data count**  
   - Counts rows, values, or categories within datasets.  
   - Useful for generating quick summaries and validation reports.  

3. **Data Segregation**  
   - Splits or separates datasets into meaningful groups.  
   - Enables category-wise analysis or multi-sheet/multi-file segregation.  

4. **Export Accumulating Data**  
   - Extracts growing/rolling datasets and exports them to files.  
   - Designed for cases where historical accumulation needs to be maintained.  

5. **Extract Email Data**  
   - Pulls data from Outlook emails or email-like text structures.  
   - Converts raw email content into structured Excel tables.  

6. **Merge Similar Data**  
   - Consolidates datasets with similar structure into one unified view.  
   - Ideal for handling multiple sources, files, or sheets with identical schema.  

---

## ⚡ Key Features

- End-to-end **Excel VBA automation** for data transformation.  
- Consistent **logging and status capture** for transparency.  
- Focused on **business analyst workflows** — quick insights, faster prep.  
- Modular codebase: each folder can run independently.  

---

## 🛠️ Requirements

- Microsoft Excel (2016 or later recommended).  
- VBA enabled (ensure macros are allowed).  
- Basic knowledge of Excel folder paths and sheet references.  

---

## 🚀 Usage

1. Clone/download this repository.  
2. Open the relevant `.xlsm` file or import the VBA module into your workbook.  
3. Ensure any required paths (folders, email sources, export locations) are set in the **Dashboard** or relevant configuration sheet.  
4. Run the macro from the VBA editor (`Alt + F11`) or via assigned buttons.  

---

## 📊 Example Output Table

| Module                  | Input                 | Output                        |
| ----------------------- | --------------------- | ----------------------------- |
| Data conversion         | Raw Excel/CSV         | Updated, formatted Excel file |
| Data count              | Dataset               | Summary counts (table)        |
| Data Segregation        | Single dataset        | Multiple sheets/files         |
| Export Accumulating Data| Periodic datasets     | Cumulative export file        |
| Extract Email Data      | Outlook mailbox/text  | Structured Excel table        |
| Merge Similar Data      | Multiple datasets     | Unified master dataset        |

---

## 📈 Roadmap

- Add **performance benchmarking** for each module.  
- Extend support for **dynamic folder paths** via user prompt.  
- Expand **error handling** and detailed logs.  

---

## 🧪 Testing

Each folder will include:  
- **Test cases**: Step-by-step instructions for validating macro correctness.  
- **Benchmark cases**: Sample data to test performance under load.  

---

## 🤝 Contribution

Contributions are welcome!  
- Fork this repo.  
- Submit PRs with improvements, new data transformations, or bug fixes.  
- Open issues for feedback or requests.  

---

## 👤 Author

Developed and maintained by **Jerrison**  
- Business Analyst | Digital Strategist  
- Specialized in automation, analytics, and VBA system design.  
