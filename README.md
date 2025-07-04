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
```

## 📘 Code Explanation (Line-by-Line)

This section explains how the `Main.java` code works.

---

```java
import java.io.*;
import java.util.regex.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

These import all the required libraries:

java.io.* for reading/writing files

java.util.regex.* for extracting values using regular expressions

Apache POI for creating Excel files



