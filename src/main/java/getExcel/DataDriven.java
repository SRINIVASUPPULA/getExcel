package getExcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataDriven {

 public void getdata() throws IOException, InvalidFormatException{
	 //Getting file location & to print file name
	 File file = new File("D:/Testing/material/DataDriven.xlsx");
	 //To read the file
	 FileInputStream fis = new FileInputStream(file);
	 System.out.println(file.getName());
	 //Accessing workbook-DataDriven
	 XSSFWorkbook workbook = new XSSFWorkbook(fis);
	 int sheets = workbook.getNumberOfSheets();
	 //Printing no of sheets present in the workbook
	 System.out.println(sheets);
	 
	for(int i=0;i<sheets;i++){
		//verifying Sheet name-String
		if(workbook.getSheetName(i).equalsIgnoreCase("String")){
			System.out.println(workbook.getSheetName(i));
			//Accessing Sheet-String-if matches with given string ,can gets that sheet-i
			XSSFSheet sheet = workbook.getSheetAt(i);
			//Iterating towards down to access Rows
			Iterator<Row> rows = sheet.iterator();
			Row firstrow = rows.next();
			System.out.println(firstrow.getCell(i));
			//initiating the cell iterator for firstrow
			Iterator<Cell> cells = firstrow.cellIterator();
			int k=0,column=0;
			//loops till get empty cell towards right
			while(cells.hasNext()){
				Cell value =cells.next();
				//verifying cell value
				if(value.getStringCellValue().equalsIgnoreCase("Testcases")){
					column = k;
				}
				k++;
			}
			//loops till empty row 
			while(rows.hasNext()){
				Row r = rows.next();
				//verifying cell value of specified row Testcases
				if(r.getCell(column).getStringCellValue().equalsIgnoreCase("purchase")){
					Iterator<Cell> cv = r.cellIterator();
					while(cv.hasNext()){
						System.out.println(cv.next().getStringCellValue());//getString values only.use other methods to get other values
					}
				}
			}
		}
	}
 
 }
}
