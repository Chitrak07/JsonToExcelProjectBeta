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
       
try (BufferedReader reader = new BufferedReader(new FileReader(inputFile))) {
            StringBuilder sb = new StringBuilder();
            String line;
            while ((line = reader.readLine()) != null) {
                sb.append(line);
}
```
Opens and reads the file using a BufferedReader and reads all lines and combines them into one string.



```java

String content = sb.toString();
        
state = extractQuotedValue(content, "state");
employees = extractQuotedValue(content, "employees");
employeeId = extractQuotedValue(content, "employeeId");

String outputFile = employeeId + ".xlsx";

Prepares the output file name using the employee ID.

```
Converts the built string to a single variable for processing.
Extracts the values of state, employees, and employeeId using a helper method.
Prepares the output file name using the employee ID.
Creates a new Excel workbook and a sheet named "Details".

```java

Row row1 = sheet.createRow(0);
row1.createCell(0).setCellValue("State");
row1.createCell(1).setCellValue(state);

Row row2 = sheet.createRow(1);
row2.createCell(0).setCellValue("Employees");
row2.createCell(1).setCellValue(employees);

```
Writes the first row with State and its value.
Writes the second row with Employees and its value.

```java

 try (FileOutputStream fileOut = new FileOutputStream(outputFile)) {
                    workbook.write(fileOut);
 }

System.out.println("Excel created: " + outputFile);

```
Saves the Excel workbook to the file named after the employee ID.
Prints a confirmation message in the console.

```java
    private static String extractQuotedValue(String text, String key) {
        String pattern = "\"" + key + "\"\\s*:\\s*\"([^\"]*)\"";
        Matcher matcher = Pattern.compile(pattern).matcher(text);
        return matcher.find() ? matcher.group(1) : "";
    }
```
A helper method that uses regex to extract quoted values for a given key like "state": "California".



