package excelfile;

import java.io.FileOutputStream;
import java.io.FileNotFoundException;
import java.io.FileInputStream;
import java.io.IOException;


import org.apache.poi.xssf.usermodel.*;


public class WritingExcelProgram {

//workbook-->sheet-->Rows-->Cells
	
	public static void main(String[] args) throws Exception {
		
		
		@SuppressWarnings("resource")
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Emp Info");
		
		
		
		Object empdata[][] = { {"EmpID", "Name","Designation"},
				{101,"Simon", "Developer" },
				{102, "Nihal", "Test Engineer" },
				{103, "Sumint", "Tech Lead"}
				};
		
		
		int rows= empdata.length;
		int cols= empdata[0].length;
		
		System.out.println(rows);//4
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
		String filePath= "C:\\Users\\Welome-Pc\\eclipse-workspace\\ExcelFileProgram\\datafiles\\employees.xlsx";
		FileOutputStream outstream = new FileOutputStream(filePath);
		workbook.write(outstream);
		outstream.close();
		
		System.out.println("employees.xlsx file written successfully..!!");
	}

}
