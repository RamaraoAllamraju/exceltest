package exceltest.exceltest;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Hello world!
 *
 */
public class App {
	
	Map<String,String> myMap = new ConcurrentHashMap <String,String>();
	
	
	public static void main(String[] args) throws IOException {
		System.out.println("Hello World!");
		FileInputStream fis = new FileInputStream(".\\Files\\Test.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		System.out.println(workbook.getSheetName(0));
		XSSFSheet sheet = workbook.getSheetAt(0);

		// System.out.println("Row -- " + sheet.getLastRowNum());
		Row row = null;
		Cell cell = null;
		int start = 0,xstart = 0, ystart=0,end = 0, xend = 0, yend=0;
		for (int i = 0; i <= sheet.getLastRowNum(); i++) {
			row = sheet.getRow(i);
			try {
				for (int j = 0; j < row.getLastCellNum(); j++) {
					cell = row.getCell(j);
					System.out.println("Row is :" + i + " and Cell is " + j + " and value in CELL is :" + cell);
					if (cell != null) {
						if (cell.equals("CONT_ID")) {
							System.out.println("Found CONT_ID and value is : " + cell);
							if(start == 0)
							{
								xstart = i;
								ystart = j;
							}
							if(end == 0)
							{
								xend = i;
								yend = j;
							}
						}
					}
				}
			} catch (NullPointerException e) {
				System.out.println("No Line");
			}
		}
	}
	
	public void readingXL()
	{
		
	}
}
