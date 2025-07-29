package javatask8;

import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.*;

public class WriteExceldata {

	public static void main(String[] args) throws Exception {
		
		try {
			XSSFWorkbook workbook = new XSSFWorkbook(); //create a workbook
			XSSFSheet sheet = workbook.createSheet("Emp Info");
			
			//write data
			String[][] empdata = {
					{"Name" , "Age" , "Email"},
					{"John Doe","30","john@test.com"},
					{"Jane Doe","28","jane@test.com"},
					{"Bob Smith","35","jacky@example.com"},
					{"Swapnil" , "37" ,"swapnil@example.com"}
			};

			int rows= empdata.length;
			int cols= empdata[0].length;
			
			System.out.println(rows);//5
			System.out.println(cols);//3
			
			for(int r=0; r<rows;r++) {
				XSSFRow row = sheet.createRow(r);
				
				for (int c=0; c<cols;c++) {
					XSSFCell cell = row.createCell(c);
					
					Object value = empdata[r][c];
					
					if(value instanceof String)
						cell.setCellValue((String)value);
					if(value instanceof Integer)
						cell.setCellValue((Integer)value);
					if(value instanceof Boolean)
						cell.setCellValue((Boolean)value);
				}
			}
			String filePath= "C:\\Users\\Welome-Pc\\eclipse-workspace\\ExcelFileProgram\\userdata\\userdata.xlsx";
			FileOutputStream outstream = new FileOutputStream(filePath);
			workbook.write(outstream);
			outstream.close();
		} catch (Exception e) {
			e.printStackTrace();
			
		}
			
			System.out.println("userdata.xlsx file written successfully..!!");
			
		
	}	
}		
			
			
			
			
			         


