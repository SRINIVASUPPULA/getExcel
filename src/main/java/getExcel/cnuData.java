package getExcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class cnuData{
	public void data() throws IOException, InvalidFormatException{
	 File file = new File("D:/Testing/material/DataDriven.xlsx");
	 FileInputStream fis = new FileInputStream(file);
	 System.out.println(file.getName());
	 XSSFWorkbook workbook = new XSSFWorkbook(fis);
	 int sheets = workbook.getNumberOfSheets();
	 System.out.println(sheets);
	 for(int i=0;i<sheets;i++){
		 
			//verifying Sheet name-String
			if(workbook.getSheetName(i).equalsIgnoreCase("String")){
				System.out.println(workbook.getSheetName(i));
				//Accessing Sheet-String-if matches with given string ,can gets that sheet-i
				XSSFSheet sheet = workbook.getSheetAt(i);
			/* Display first Row values
				Iterator<Row> rows = sheet.iterator();

				Iterator<Cell> cells = rows.next().cellIterator();
				
				while(cells.hasNext()){
				System.out.println(cells.next().getStringCellValue());
				} */
			int rowcount = sheet.getLastRowNum();
			int colcount = sheet.getRow(0).getLastCellNum();
			//row count giving "-1" so j should be <=rowcount 
			System.out.println(rowcount+" "+colcount);
			//looping through Rows
			for(int j=0;j<=rowcount;j++){
				//storing row 
				XSSFRow currentrow = sheet.getRow(j);
				//verifying the cell content of that specific row
				if(currentrow.getCell(0).getStringCellValue().equalsIgnoreCase("Purchase")){
				for(int k=0;k<colcount;k++){
					//looping through cell values of required row
					String currentcell = currentrow.getCell(k).toString();//toString return everything as string values
					System.out.print(currentcell+" ");
				}}
				//System.out.println();
			}
			}
				
			}	
			}
	}