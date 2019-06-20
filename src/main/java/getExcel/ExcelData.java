package getExcel;

import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

public class ExcelData{

	public static void main(String[] args) throws InvalidFormatException, IOException {
		DataDriven dd = new DataDriven();
		dd.getdata();
		cnuData cd = new cnuData();
		cd.data();
	}

}
