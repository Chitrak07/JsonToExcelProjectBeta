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

```

These import all the required libraries:

java.io.* for reading/writing files

java.util.regex.* for extracting values using regular expressions

Apache POI for creating Excel files


```java
        String inputFile = "input.txt";
        String state = "";
        String employees = "";
        String employeeId = "";
        // Opens and reads the file using a BufferedReader and reads all lines and combines them into one string.
        try (BufferedReader reader = new BufferedReader(new FileReader(inputFile))) {
            StringBuilder sb = new StringBuilder();
            String line;
            while ((line = reader.readLine()) != null) {
                sb.append(line);
            }
```




```java
            //Converts the built string to a single variable for processing.
            String content = sb.toString();

            //Extracts the values of state, employees, and employeeId using a helper method.
            state = extractQuotedValue(content, "state");
            employees = extractQuotedValue(content, "employees");
            employeeId = extractQuotedValue(content, "employeeId");


```



