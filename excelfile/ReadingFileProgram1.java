package excelfile;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.*;


public class ReadingFileProgram1 {

	public static void main(String[] args) throws IOException {
		String excelFilePath = ".\\datafiles\\test.xlsx";
		FileInputStream inputstream = new FileInputStream(excelFilePath);
		
		
		@SuppressWarnings("resource")
		XSSFWorkbook workbook = new XSSFWorkbook(inputstream);
		
	//	XSSFSheet sheet = workbook.getSheet("Sheet 1");
		XSSFSheet sheet = workbook.getSheetAt(0);
		
		int rows = sheet.getLastRowNum();
		int cols = sheet.getRow(1).getLastCellNum();
		 System.out.println(rows);
		 System.out.println(cols);
		
		 
	/*	 for (int r= 0; r<=rows;r++) {
			 
			 XSSFRow row = sheet.getRow(r);
			 
			 for (int c= 0; c<cols;c++) {
				 
				 XSSFCell cell = row.getCell(c);
				 
				 switch (cell.getCellType()) {
				 case NUMERIC:
					 System.out.print(cell.getNumericCellValue());
					 break; 
				 case STRING:
					 System.out.print(cell.getStringCellValue());
					 break; 
				 case BOOLEAN:
					 System.out.print(cell.getBooleanCellValue());
					 break;
				default:
					break; 
 
				 }
				 
				 
				 System.out.print("| ");
			 }
			 System.out.println();
						 
				 
			 }*/
		 
		 	Iterator<?> iterator = sheet.iterator();
		 	
		 	while(iterator.hasNext()) {
		XSSFRow row = (XSSFRow)iterator.next();
		Iterator<?> cellIterator = row.cellIterator();
			
		while(cellIterator.hasNext()) {
		XSSFCell cell = (XSSFCell)cellIterator.next();
		switch (cell.getCellType()) {
		 case NUMERIC:
			 System.out.print(cell.getNumericCellValue());
			 break; 
		 case STRING:
			 System.out.print(cell.getStringCellValue());
			 break; 
		 case BOOLEAN:
			 System.out.print(cell.getBooleanCellValue());
			 break;
		default:
			break; 

		 }
		 
		 
		 System.out.print("| ");
	 }
	 System.out.println();
	}
	
	
	}
	
	
	}
		 
	

