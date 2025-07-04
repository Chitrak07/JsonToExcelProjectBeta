# JsonToExcelProjectBeta

## 💡 Code Explanation

This project reads a `.txt` file containing JSON-like data and writes selected fields (like `state` and `employees`) into an Excel file using Apache POI. The Excel file is named based on the `employeeId`.

---

### 🔍 Main Features:
- Reads structured data from `input.txt`
- Extracts values like `state`, `employees`, and `employeeId` (all values are in quotes)
- Writes the extracted data into an Excel file
- Output file is named as: `<employeeId>.xlsx` (e.g., `EMP1024.xlsx`)

---

### 📁 input.txt format (example):
```json
{
  "employeeId": "EMP1024",
  "state": "California",
  "employees": "150000"
}
