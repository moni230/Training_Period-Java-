import org.apache.poi.ss.formula.functions.Column;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelWriter {
    public static void main(String[] args) throws IOException {

        XSSFWorkbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Students");

        Row row1 = sheet.createRow(0);
        Cell cell1 = row1.createCell(0);
        Cell cell2 = row1.createCell(1);
        Cell cell3 = row1.createCell(2);

        cell1.setCellValue("ID");
        cell2.setCellValue("NAME");
        cell3.setCellValue("SALARY");

        Row row2 = sheet.createRow(1);
        Cell id1 = row2.createCell(0);
        Cell name1 = row2.createCell(1);
        Cell salary1 = row2.createCell(2);

        id1.setCellValue("101");
        name1.setCellValue("Arun");
        salary1.setCellValue("50000");

        Row row3 = sheet.createRow(2);
        Cell id2 = row3.createCell(0);
        Cell name2 = row3.createCell(1);
        Cell salary2 = row3.createCell(2);

        id2.setCellValue("102");
        name2.setCellValue("Ajay");
        salary2.setCellValue("40000");

        Row row4 = sheet.createRow(3);
        Cell id3 = row4.createCell(0);
        Cell name3 = row4.createCell(1);
        Cell salary3 = row4.createCell(2);

        id3.setCellValue("103");
        name3.setCellValue("Ajith");
        salary3.setCellValue("70000");

        FileOutputStream fs = new FileOutputStream("StudentsDataInfo.xlsx");
        workbook.write(fs);
        System.out.println("Contents are inserted into excel sheet");

//      System.out.println(new File ("StudentsDataSet.xlsx").getAbsolutePath());


     /*
      Create a FileOutputStream object.
   Specify the filename as employees.xlsx.
   Pass that output stream to workbook.write()
      */
    }

}



