import java.io.FileInputStream;
import java.io.IOException;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.pdfbox.pdmodel.font.Standard14Fonts.FontName;
import org.apache.pdfbox.util.Matrix;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelToPdf {
    public static void main(String[] args) throws IOException {


        // create the pdf document.
        PDDocument document = new PDDocument();

        // create the page
        PDPage page = new PDPage();

        // add page into the document
        document.addPage(page);

        // to insert/write the content on the page.
        PDPageContentStream content = new PDPageContentStream(document, page);

        // set the font.
        PDType1Font font = new PDType1Font(FontName.TIMES_ROMAN);

        // start to write the text.
        content.beginText();

        // set the font size.
        content.setFont(font, 12);



        // set the position of the values.
        content.newLineAtOffset(50, 700);

//		content.showText("Hello World");

        // open the excel file in the current working directory.
        FileInputStream fs = new FileInputStream("Employees.xlsx");


        // read data from the excel
        XSSFWorkbook workbook = new XSSFWorkbook(fs);

        // access the specified sheet
        XSSFSheet sheet = workbook.getSheetAt(0);

        // converting the sheets into row
        for(Row row : sheet) {
            // converting rows into cell, to access data.
            for (Cell cell : row) {

                // store the row & column index position.
                int rowIndex = row.getRowNum();
                int columnIndex = cell.getColumnIndex();

                // calculate the horizontal & vertical position to place values in pdf.
                float  x = 50 + (columnIndex * 120);
                float  y = 750 - (rowIndex *30);

                // tells the PDFBox where to position the text.
                content.setTextMatrix(Matrix.getTranslateInstance(x, y));

                // used to convert the cell data into readable string.
                DataFormatter data = new DataFormatter();

                // used to add the text into the pdf.
                content.showText(data.formatCellValue(cell)+ " ");

            }
            //to start with next line, after the row ends.
            System.out.println();

        }

        // end the text.
        content.endText();
        //close the content.
        content.close();
        // save the document.
        document.save("EmployeesDataInfo.pdf");
        // close the document.
        document.close();

        // to print this statement to verify the file is converted or not!
        System.out.println("excel file converted into PDF");

    }

}

