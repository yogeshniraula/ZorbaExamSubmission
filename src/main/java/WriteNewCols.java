import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

public class WriteNewCols {

    public static void main(String[] args) throws Exception {

        File file = new File("C:\\Users\\yoges\\IdeaProjects\\ZorbaExamSubmission\\src\\main\\resources\\employee.xlsx");
        FileInputStream fileInputStream = new FileInputStream(file);

        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(0);

        String[] newHeaders = {"Manager_id  ", "emp_dept", "emp_share (%)"};

        String[][] newData = {
                {"Null", "Finance", "60"},
                {"1001", "Finance", "20"},
                {"1004", "R&D", "30"},
                {"1004", "R&D", " 40"},
                {"1001", "Finance", " 20"},
                {"1005", "Finance", "15"},
                {"1001", "Finance", "25"}
        };

        Row headerRow = xssfSheet.getRow(0);
         int nCols =  headerRow.getLastCellNum();



        for (int i = 0; i < newHeaders.length; i++) {
            Cell cell = headerRow.createCell(nCols+i);
            cell.setCellValue(newHeaders[i]);
        }

        for (int i = 0; i < newData.length; i++) {
            Row row = xssfSheet.getRow(i + 1);
            for (int j = 0; j < newData[i].length; j++) {
                Cell cell1 = row.createCell(nCols+j);
                cell1.setCellValue(newData[i][j]);
            }

        // Write changes to the file
        try (FileOutputStream fileOutputStream = new FileOutputStream(file)) {
            xssfWorkbook.write(fileOutputStream);
        }

    }
        FileOutputStream fileOutputStream = new FileOutputStream(file);
        xssfWorkbook.write(fileOutputStream);
        fileOutputStream.close();
        System.out.println("Successfully Write back to excel file.." + headerRow.getLastCellNum());
    }
}



