# Test Cases – vba-data-transform

This document defines test scenarios to validate the functionality of each module.

---

## ✅ Data Conversion (Read & Update)

| Test ID | Scenario | Input | Expected Output | Result |
|---------|----------|-------|-----------------|--------|
| DC-01 | Update column values | 10 rows, column B with blanks | Column B filled with defaults | Pass/Fail |
| DC-02 | Normalize text | Mixed case data | All converted to UPPERCASE | Pass/Fail |
| DC-03 | Handle empty sheet | No rows | Macro ends gracefully | Pass/Fail |

---

## ✅ Data Count

| Test ID | Scenario | Input | Expected Output | Result |
|---------|----------|-------|-----------------|--------|
| CNT-01 | Count rows | 100 rows | Returns 100 | Pass/Fail |
| CNT-02 | Count unique values | Column with 10 unique IDs | Returns 10 | Pass/Fail |
| CNT-03 | Handle empty column | Blank column | Returns 0 | Pass/Fail |

---

## ✅ Data Segregation

| Test ID | Scenario | Input | Expected Output | Result |
|---------|----------|-------|-----------------|--------|
| SEG-01 | Split by region | 1 sheet, 100 rows, 2 regions | 2 new sheets with 50 rows each | Pass/Fail |
| SEG-02 | Multi-condition split | 3 categories | 3 new files created | Pass/Fail |
| SEG-03 | No matching condition | Data not matching rule | Output file empty but created | Pass/Fail |

---

## ✅ Export Accumulating Data

| Test ID | Scenario | Input | Expected Output | Result |
|---------|----------|-------|-----------------|--------|
| EXP-01 | Append monthly file | 2k rows (Jan), then 2k rows (Feb) | Final log = 4k rows | Pass/Fail |
| EXP-02 | Handle duplicates | Same row added twice | No duplicates in final log | Pass/Fail |
| EXP-03 | Empty new data | Empty Feb file | Log unchanged | Pass/Fail |

---

## ✅ Extract Email Data

| Test ID | Scenario | Input | Expected Output | Result |
|---------|----------|-------|-----------------|--------|
| EM-01 | Extract simple emails | 10 emails in text | Output = 10 clean emails | Pass/Fail |
| EM-02 | Remove duplicates | 20 rows, 5 duplicates | Output = 15 unique emails | Pass/Fail |
| EM-03 | Handle invalid emails | "abc@, test@" | Invalid entries skipped | Pass/Fail |

---

## ✅ Merge Similar Data

| Test ID | Scenario | Input | Expected Output | Result |
|---------|----------|-------|-----------------|--------|
| MRG-01 | Merge identical rows | 2 sheets with same 100 rows | Final sheet = 100 rows | Pass/Fail |
| MRG-02 | Merge partial overlap | Sheet A = 100 rows, Sheet B = 120 rows (20 overlap) | Final = 200 rows | Pass/Fail |
| MRG-03 | Handle large dataset | 5 sheets, 100k rows | Final merged sheet created successfully | Pass/Fail |

---

## 📌 Test Execution

1. Store test input files under `/tests/data/`.  
2. Run macros manually via Dashboard or VBA Editor.  
3. Record outcome in the **Result** column.  

---
