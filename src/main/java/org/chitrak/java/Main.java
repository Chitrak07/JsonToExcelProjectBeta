import java.io.*;
import java.util.regex.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {
    public static void main(String[] args) {
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

            String content = sb.toString();

            // All values are quoted — treat everything as string
            state = extractQuotedValue(content, "state");
            employees = extractQuotedValue(content, "employees");
            employeeId = extractQuotedValue(content, "employeeId");

            if (employeeId.isEmpty()) {
                System.out.println("Error: employeeId not found in input file.");
                return;
            }

            String outputFile = employeeId + ".xlsx";

            // Write to Excel
            try (Workbook workbook = new XSSFWorkbook()) {
                Sheet sheet = workbook.createSheet("Details");

                Row row1 = sheet.createRow(0);
                row1.createCell(0).setCellValue("State");
                row1.createCell(1).setCellValue(state);

                Row row2 = sheet.createRow(1);
                row2.createCell(0).setCellValue("Employees");
                row2.createCell(1).setCellValue(employees);

                try (FileOutputStream fileOut = new FileOutputStream(outputFile)) {
                    workbook.write(fileOut);
                }

                System.out.println("Excel created: " + outputFile);
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    //  Extract any quoted value, regardless of type
    private static String extractQuotedValue(String text, String key) {
        String pattern = "\"" + key + "\"\\s*:\\s*\"([^\"]*)\"";
        Matcher matcher = Pattern.compile(pattern).matcher(text);
        return matcher.find() ? matcher.group(1) : "";
    }
}
