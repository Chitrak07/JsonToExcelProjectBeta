import java.io.*;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.*;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {
    public static void main(String[] args) {
        String inputFile = "test1.txt";

//        String state = "";
//        String employees = "";
        String templateFile = "template.xlsx";
        String employeeId = "";
        try {
            StringBuilder sb = new StringBuilder();

            try (BufferedReader reader = new BufferedReader(new FileReader(inputFile))) {
                //StringBuilder sb = new StringBuilder();
                String line;
                while ((line = reader.readLine()) != null) {
                    sb.append(line);
                }
            }

            String content = sb.toString();

            // Read fields to extract from template excel
            List<String> fieldsToExtract = new ArrayList<>();
            try (Workbook workbook = new XSSFWorkbook(new FileInputStream(templateFile))) {
                Sheet sheet = workbook.getSheetAt(0);
                for (Row row : sheet) {
                    Cell cell = row.getCell(0); // assume field names in column A
                    if (cell != null && cell.getCellType() == CellType.STRING) {
                        fieldsToExtract.add(cell.getStringCellValue());
                    }
                }
            }


            //extract values

            Map<String, String> extractedValues = new LinkedHashMap<>();
            for (String key : fieldsToExtract) {
                String value = extractQuotedValue(content, key);
                if (key.equalsIgnoreCase("employeeId")) {
                    employeeId = value;
                }
                extractedValues.put(key, value);
            }


//            // All values are quoted — treat everything as string
//            state = extractQuotedValue(content, "state");
//            employees = extractQuotedValue(content, "employees");
//            employeeId = extractQuotedValue(content, "employeeId");

            //  Prepare output file name
            if (employeeId.isEmpty()) {
                System.out.println("employeeId not found. Using default name: output.xlsx");
                employeeId = "output";
            }

            String outputFile = employeeId + ".xlsx";

            // Write extracted data to new Excel


//                Row row2 = sheet.createRow(1);
            try (Workbook outWorkbook = new XSSFWorkbook()) {
                Sheet outSheet = outWorkbook.createSheet("Extracted Data");

                int rowNum = 0;
                for (Map.Entry<String, String> entry : extractedValues.entrySet()) {
                    Row row = outSheet.createRow(rowNum++);

//                Row row1 = sheet.createRow(0);
//                row1.createCell(0).setCellValue("State");
//                row1.createCell(1).setCellValue(state);

                    row.createCell(0).setCellValue(entry.getKey());
                    row.createCell(1).setCellValue(entry.getValue());
                }

                try (FileOutputStream fileOut = new FileOutputStream(outputFile)) {
                    outWorkbook.write(fileOut);
                }

                System.out.println("Excel created: " + outputFile);
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    //  // Helper to extract value from "key": "value"
    private static String extractQuotedValue(String text, String key) {
        String pattern = "\"" + key + "\"\\s*:\\s*\"([^\"]*)\"";
        Matcher matcher = Pattern.compile(pattern).matcher(text);
        return matcher.find() ? matcher.group(1) : "";
    }
}
