import java.io.File;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileOutputStream;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

public class WriteToExcel {
    public static void main(String[] args) throws Exception {
        //Read the existing excel file
        File file = new File("C:\\Users\\yoges\\IdeaProjects\\ZorbaExamSubmission\\src\\main\\resources\\employee.xlsx");
        FileInputStream fileInputStream = new FileInputStream(file);

        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(0);

        XSSFRow row;


        Map <String, Object[]> EmployeeData = new TreeMap<String, Object[]>();
        EmployeeData.put( "2",
               new Object[]{ "1001", "Jack", "1482.45", "0809808008", "NYC"}
        );
        EmployeeData.put( "3",
                new Object[]{ "1002", "Joy", "5282.12", "9809808008", "SD"}
        );
        EmployeeData.put( "4",
                new Object[]{ "1003", "Nick", "3454.11", "8976876786", "Dayton"}
        );

        EmployeeData.put( "5",
                new Object[]{ "1004", "Joe", "6482.45", "8809808008", "NYC"}
        );
        EmployeeData.put( "6",
                new Object[]{ "1005", "Nick", "5482.45", "5809808008", "CA"}
        );
        EmployeeData.put( "7",
                new Object[]{ "1006", "Hyder", "9482.45", "2809808008", "LA"}
        );
        EmployeeData.put( "8",
                new Object[]{ "1007", "Harry", "1182.45", "4809808008", "Ohio"}
        );
        Set<String> keyid = EmployeeData.keySet();
            int rowid = 1;

         for (String key : keyid) {

row = xssfSheet.createRow(rowid++);
Object[] objectArr = EmployeeData.get(key);
int cellid = 0;

            for (Object obj : objectArr) {
Cell cell = row.createCell(cellid++);
                cell.setCellValue((String)obj);
        }
        }
        //Write back to Excel file
        FileOutputStream fileOutputStream = new FileOutputStream(file);
        xssfWorkbook.write(fileOutputStream);
        fileOutputStream.close();
        System.out.println("Successfully Write back to excel file..");
    }
}
