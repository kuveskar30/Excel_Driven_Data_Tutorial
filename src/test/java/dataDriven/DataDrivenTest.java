package dataDriven;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

//Apache POI API, POI API are java libraries having jar files which have methods
//which help to get data from excel file
public class DataDrivenTest {

	public ArrayList<String> getDataFromExcel(String testCaseName) throws IOException {
		ArrayList<String> testCaseDataArrayList = new ArrayList<String>();

		// FileInputStream is a Class in java which creates an object which has power to
		// read any file provided to it
		FileInputStream fis = new FileInputStream(
				"E:\\pratik30\\Software testing\\Udemy_selenium_course\\ExcelDataDriven\\DemoData.xlsx");
		// By below step we are getting access to excel workbook
		XSSFWorkbook workbook = new XSSFWorkbook(fis);

		// below steps we are selecting required sheet from available sheets
		// from selected sheet we are entering in rows and selecting required row
		// and from selected row we entering in cells and selecting required cell
		int sheet_count = workbook.getNumberOfSheets();
		for (int i = 0; i < sheet_count; i++) {
			if (workbook.getSheetName(i).equalsIgnoreCase("testdata")) {
				// sheet is a collection of rows
				XSSFSheet sheet = workbook.getSheetAt(i);
				// to traverse to each row we are using rowIterator
				Iterator<Row> rows_iterator = sheet.rowIterator();
				// row is a collection of cells
				Row row1 = rows_iterator.next();
				// to traverse to each cell we are using cellIterator
				Iterator<Cell> cells_iterator = row1.cellIterator();
				int column_index_no = 0;
				while (cells_iterator.hasNext()) {
					Cell cell = cells_iterator.next();
					if (cell.getStringCellValue().equalsIgnoreCase("testcases")) {
						// I'm breaking this inner while loop on getting desired column
						break;
					}
					column_index_no++;
				}

					System.out.println(column_index_no);
					
				while (rows_iterator.hasNext()) {
					Row row2 = rows_iterator.next();
					if (row2.getCell(column_index_no).getStringCellValue().equalsIgnoreCase(testCaseName)) {
						Iterator<Cell> cell_iterator2 = row2.cellIterator();
						while (cell_iterator2.hasNext()) {
							Cell cell2 = cell_iterator2.next();
							if(cell2.getCellType() == CellType.STRING) {
							testCaseDataArrayList.add(cell2.getStringCellValue());
							}else {
								String s = NumberToTextConverter.toText(cell2.getNumericCellValue());
//								String s = String.valueOf(cell2.getNumericCellValue());
//								System.out.println(cell2.getNumericCellValue());
								testCaseDataArrayList.add(s);
							}
						}
						break;
					}

				}
			}
		}
		return testCaseDataArrayList;

	}
}
