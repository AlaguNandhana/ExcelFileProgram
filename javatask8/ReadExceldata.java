package javatask8;

import java.io.FileInputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExceldata {

	public static void main(String[] args) {
		try {
            FileInputStream file = new FileInputStream("userdata.xlsx");
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            XSSFSheet sheet = workbook.getSheetAt(0); 

            int rowCount = sheet.getLastRowNum();

            //read row
            for (int i = 0; i <= rowCount; i++) {
                XSSFRow row = sheet.getRow(i);
                int colCount = row.getLastCellNum();
                //read cell
                for (int j = 0; j < colCount; j++) {
                    XSSFCell cell = row.getCell(j);
                    System.out.print(cell.toString() + "\t");
                }
                System.out.println(); // Newline after each row
            }

            workbook.close();
            file.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
