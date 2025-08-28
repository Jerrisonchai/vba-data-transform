# Performance Benchmark – vba-data-transform

This document tracks performance of the VBA modules on different dataset sizes.  
Benchmarks are recorded to monitor scalability and optimize execution.

---

## ⚙️ Test Environment

- Microsoft Excel 365 (64-bit)  
- Intel i7 (2.8GHz), 16GB RAM  
- Windows 11 Pro  
- FileSystemObject (late binding) for file operations  

---

## 📊 Benchmark Summary

| Module                  | Dataset Size  | Avg. Time Taken | Memory Impact | Notes |
|--------------------------|---------------|-----------------|---------------|-------|
| Data Conversion          | 10k rows      | 0.8s            | Low           | Scales linearly |
|                          | 100k rows     | 4.2s            | Medium        | |
| Data Count               | 10k rows      | 0.5s            | Negligible    | |
|                          | 100k rows     | 2.9s            | Low           | |
| Data Segregation         | 10k rows (5 splits)  | 1.6s | Medium | Performance depends on # of splits |
|                          | 100k rows (10 splits)| 8.7s | Medium | |
| Export Accumulating Data | 12 monthly files (2k rows each) | 3.2s | Low | Appends fast |
| Extract Email Data       | 10k cells     | 1.3s            | Low           | Regex operations efficient |
|                          | 100k cells    | 6.9s            | Medium        | |
| Merge Similar Data       | 10k rows (2 files)  | 2.4s | Medium | Sorting-intensive |
|                          | 100k rows (5 files) | 15.8s | High | Optimizations needed |

---

## 🚀 Optimization Notes

- **Data Conversion** → Use `Application.ScreenUpdating = False` to reduce runtime.  
- **Data Segregation** → Chunked writing reduces memory spikes.  
- **Merge Similar Data** → Performance bottleneck when handling >100k rows across multiple files. Recommend pre-sorting before merge.  
- **Extract Email Data** → Scales well but regex-heavy operations can slow down at >250k cells.  

---

## 📌 Key Takeaways

- All modules are stable up to **100k rows** in <20s runtime.  
- For >1M rows, recommend chunking with Power Query or Python post-processing.  
- Log files confirm that **Data Segregation** and **Merge Similar Data** are the most resource-intensive modules.  

---
