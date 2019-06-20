package getExcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class simpleLoop {
	public static void main(String[] args) throws IOException {
		
//THis wont pass Numeric values- see getExceldata for the solution
		 File file = new File("C:\\Users\\Hari\\Desktop\\GoogleFormTestdata.xlsx");
		 FileInputStream fis = new FileInputStream(file);
		 System.out.println(file.getName());
		 XSSFWorkbook workbook = new XSSFWorkbook(fis);
	
		 XSSFSheet sheet =  workbook.getSheet("Sheet1");
		 int rowcount = sheet.getLastRowNum();
		 System.out.println(rowcount+"  "+sheet.getRow(0).getLastCellNum());
		 for(int i=0;i<rowcount;i++) {
			 XSSFRow row = sheet.getRow(i);
			 
			 for(int j=0;j<row.getLastCellNum();j++) {
				 System.out.print(row.getCell(j).toString()+" || ");
			 }
			 System.out.println();
		 }
	}}
