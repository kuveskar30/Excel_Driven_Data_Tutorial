package dataDriven;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

//apache poi api is a library having jar file which have methods
//which help to get data from excel file
public class DataDrivenTestModifiedOuterLoopCopy {

	public static void main(String[] args) throws IOException {
		// step1. create XSSFWorkbook object

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

				int row_index_no = 0;
				//actually no need of outer while loop as always header of table will
				//be in 1st row of iterator
				while (rows_iterator.hasNext()) {
					// row is a collection of cells
					Row row = rows_iterator.next();
					// to traverse to each cell we are using cellIterator
					Iterator<Cell> cell_iterator = row.cellIterator();
					int coulmn_count = 0;
					int column_index_no = 0;
					boolean break_outer_while_loop = false;
					while (cell_iterator.hasNext()) {
						Cell cell = cell_iterator.next();
						if (cell.getStringCellValue().equalsIgnoreCase("testcases")) {
							column_index_no = coulmn_count;
							break_outer_while_loop = true;
							//I'm breaking this inner while loop on getting desired column
							break;
						}
						coulmn_count++;
					}
					
//					System.out.println(column_index_no);
//					System.out.println(coulmn_count);
					
					//I'm breaking outer while loop on getting desired column
					if(break_outer_while_loop) {
						break;
					}
					row_index_no++;
				}
				System.out.println(row_index_no);
			}
		}

	}

}
