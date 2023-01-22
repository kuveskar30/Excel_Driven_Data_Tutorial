package dataDrivenDataProvider;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;


public class DataProviderWithExcelData {

	DataFormatter formatter = new DataFormatter();
	
	@Test(dataProvider="dataFromExcel")
	public void driveDataFromExcel(String name, String city, String id) {
		System.out.println(name);
		System.out.println(city);
		System.out.println(id);
	}



	@DataProvider(name = "dataFromExcel")
	public Object[][] getData() throws IOException {

		FileInputStream fis = new FileInputStream(
				"E:\\pratik30\\Software testing\\Udemy_selenium_course\\ExcelDataDriven\\excelDataProvider.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.getSheetAt(0);

		int row_count = sheet.getPhysicalNumberOfRows();
		XSSFRow row = sheet.getRow(0);
		int column_count = row.getPhysicalNumberOfCells();

		Object[][] data = new Object[row_count - 1][column_count];
		for (int r = 0; r < row_count-1; r++) {
			row = sheet.getRow(r+1);
			for (int c = 0; c < column_count; c++) {
				XSSFCell cell = row.getCell(c);
				data[r][c] = formatter.formatCellValue(cell);
			}
		}
		return data;
	}

}
