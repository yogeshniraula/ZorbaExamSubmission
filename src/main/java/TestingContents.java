import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

public class TestingContents {

        public static void main(String[] args) {
            String filePath = "C:\\Users\\yoges\\IdeaProjects\\ZorbaExamSubmission\\src\\main\\resources\\employee.xlsx"; // Path to your XLSX file
            int rowIndex = 0; // Index of the row whose length you want to find

            try (FileInputStream fileInputStream = new FileInputStream(filePath);
                 Workbook workbook = new XSSFWorkbook(fileInputStream)) {

                Sheet sheet = workbook.getSheetAt(0); // Get the first sheet

                // Get the row at the specified index
                Row row = sheet.getRow(0);

                if (row != null) {
                    // Get the number of columns in the row
                    int numberOfColumns = row.getLastCellNum(); // Returns the number of cells in the row

                    // Display the length of the row
                    System.out.println("Number of columns in row " + rowIndex + ": " + numberOfColumns);
                } else {
                    System.out.println("Row " + rowIndex + " does not exist.");
                }

            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

