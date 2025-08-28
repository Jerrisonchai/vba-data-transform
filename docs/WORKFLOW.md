# Workflow – vba-data-transform

This document outlines how different utilities interact and how data flows through the repo.

---

## 🔄 Data Flow Overview

```
Raw Data (Excel / CSV)
│
▼
Data Conversion (cleanup, updates)
│
▼
Data Count (audit, verify sizes)
│
▼
Data Segregation (split by rules)
│
├──> Export Accumulating Data (log for history)
│
▼
Extract Email Data (if applicable)
│
▼
Merge Similar Data (final consolidation)
│
▼
Output Dataset (ready for analysis/reporting)
```

---

## 🛠 Typical Workflow

1. **Start with Raw Data**  
   Place Excel/CSV files in your designated `/input/` folder.  

2. **Transform & Clean**  
   Use `Data Conversion` module to apply updates/standard formatting.  

3. **Audit**  
   Run `Data Count` to verify record totals.  

4. **Split or Segregate**  
   Use `Data Segregation` to create separate files per category.  

5. **Accumulate Over Time**  
   Export recurring new data into a cumulative log.  

6. **Extract & Merge**  
   Extract emails (if needed), then merge similar datasets.  

7. **Final Output**  
   Save results to `/output/` for use in dashboards, BI tools, or emailing.  

---

## 📊 Logging

Every macro logs:  
- `Status` (Success/Failure)  
- `Start_Time` and `Time_Taken`  
- `UserName` (system user running the macro)  

This ensures traceability for each run.  

---

## 🔗 Integration

- Compatible with **Excel Dashboards** (cell references for paths).  
- Can be combined with **Power Query** or **Python post-processing**.  
- Recommended for e-commerce, reporting pipelines, and recurring data prep tasks.  

---

