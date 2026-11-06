import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileInputStream;
import java.io.IOException;

        public class ExcelReader {
            public static void main(String[] args) throws IOException {

                // file path
                String filePath = "StudentsData.xlsx";

                // Open the Excel files

                FileInputStream fs = new FileInputStream(filePath);

                // read data from the Excel
                Workbook workbook = new XSSFWorkbook(fs);

                // access the particular sheet by using index position
                Sheet sheet = workbook.getSheetAt(0);

                // converting sheet into rows
                for (Row row : sheet) {

                    // rows converted into cell(to extract particular data)
                    for (Cell cell : row) {

                        // print the cell values
                        System.out.print(cell.toString()+ " ");
                    }

                    //to start with next line, after the row ends.
                    System.out.println();
                }
            }
        }

