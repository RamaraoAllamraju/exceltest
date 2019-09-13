package exceltest.exceltest;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteIToExcel {
	static String path = "C:\\Users\\rallamr\\Desktop\\Ramarao\\captured\\test.xlsx";
	
	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		System.out.println("Here");
		XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(path));
		XSSFSheet sheet = wb.getSheet("Sheet1");
		Row row = sheet.createRow(0);
		Cell cell= row.createCell(0);
		cell.setCellValue("Ramarao");
		FileOutputStream fos = new FileOutputStream(path);
		wb.write(fos);
		fos.close();
		System.out.println("Done");
	}

}
