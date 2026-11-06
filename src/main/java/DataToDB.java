import java.io.FileInputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataToDB {
    public static void main(String[] args) throws IOException, SQLException {

        //open the excel file

        FileInputStream fs = new FileInputStream("Employees_Info_data.xlsx");

        //read data from excel
        XSSFWorkbook workbook = new XSSFWorkbook(fs);

        //open the first sheet
        XSSFSheet sheet = workbook.getSheetAt(0);

        //establish a JDBC Connection
        Connection con = DriverManager.getConnection("jdbc:mysql://localhost:3306/employee_db", "root", "Hope@2026");

        // store the SQL query
        String query = "INSERT INTO employee(id, name, salary) VALUES(?,?,?)";

        //pass the query
        PreparedStatement ps =   con.prepareStatement(query);

        // sheet to be red until the last row.
        for(int i = 1; i<= sheet.getLastRowNum(); i++) {

            // store the current row
            Row row = sheet.getRow(i);


            // checks the row, if null(no-value), skips & proceed with next iteration.
            if(row == null) {
                continue;
            }

            // used to extract the excel values into readable string.
            DataFormatter data = new DataFormatter();

            // stores the data values based on their index position.
            String id = data.formatCellValue(row.getCell(0));
            String name = data.formatCellValue(row.getCell(1));
            String salary =data.formatCellValue(row.getCell(2));


            // if id, name & salary any of the field is blank, skips current iteration and proceed with next.
            if(id.isBlank() || name.isBlank() || salary.isBlank()) {
                continue;
            }

            // set the values to the placeholder based on respective db column
            ps.setInt(1, Integer.parseInt(id));
            ps.setString(2, name);
            ps.setInt(3,Integer.parseInt(salary));

            // updates the values in db.
            ps.executeUpdate();

        }

        // print this statement to confirm the values are inserted.
        System.out.println("records inserted successfully!");

    }

}

