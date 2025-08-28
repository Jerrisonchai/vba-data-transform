# Developer Guide – vba-data-transform

This guide is for developers contributing to or maintaining the VBA Data Transform utilities.

---

## 🔧 Project Structure

```
vba-data-transform/
│
├── data_conversion_read_update/
├── data_count/
├── data_segregation/
├── export_accumulating_data/
├── extract_email_data/
├── merge_similar_data/
│
├── README.md
├── DEV_GUIDE.md
├── USER_GUIDE.md
├── WORKFLOW.md
├── performance_benchmark.md
└── test_cases.md
```

---

## ⚙️ Development Setup

1. **Environment**  
   - Microsoft Excel 2016 or later (recommended: Microsoft 365).  
   - VBA enabled in the Excel Trust Center.  
   - FileSystemObject library available (late binding used, no external references required).  

2. **Code Location**  
   - Each folder contains `.bas` or `.cls` VBA modules.  
   - Import them into an Excel VBA project using **File > Import File**.  

3. **Naming Conventions**  
   - Functions and Subs use `PascalCase` (e.g., `TransformData`).  
   - Variables follow `camelCase` (e.g., `rowCount`, `sourcePath`).  
   - Modules are prefixed with the feature name (e.g., `modDataCount`).  

4. **Error Handling**  
   - Use `On Error Resume Next` only for recoverable errors.  
   - Implement `On Error GoTo ErrorHandler` for critical sections.  

---

## 🧪 Testing

- Unit tests are documented in `test_cases.md`.  
- Manual tests should be run by preparing sample Excel/CSV files in `/tests/data/`.  
- Ensure consistent results before committing.  

---

## 🚀 Contribution Workflow

1. Fork the repo and create a feature branch:  
   ```bash
   git checkout -b feature/new-transform
   ```
2. Add or update VBA modules.
3. Update corresponding docs (USER_GUIDE.md, test_cases.md).
4. Submit a Pull Request with clear description of changes.

## 🚀 Best Practices

- Keep transformations modular (one function = one purpose).
- Always log runtime, start/end time, and user (using Environ("UserName")).
- Avoid hardcoding file paths; use Dashboard sheet cells as config.
