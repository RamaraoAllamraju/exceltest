package exceltest.exceltest;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.HashMap;
import java.util.concurrent.ConcurrentHashMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadXL {

	public ReadXL() {
		// TODO Auto-generated constructor stub
	}
	
	public static void main(String args[]) throws IOException
	{
		ConcurrentHashMap<String, String> myMap = new ConcurrentHashMap<String,String>();
		int rowNum = 2;
		
		FileInputStream fis = new FileInputStream(".\\Files\\Test.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.getSheetAt(0);
		System.out.println("Last Row is :"+sheet.getLastRowNum());
		
		Row row = null;
		Cell cell = null;
		row = sheet.getRow(0);
		for(int i=0;i<row.getLastCellNum();i++)
		{
			cell = row.getCell(i);
			myMap.put(sheet.getRow(0).getCell(i).toString(), sheet.getRow(rowNum).getCell(i).toString());
			System.out.println(sheet.getRow(0).getCell(i)+" --> "+sheet.getRow(rowNum).getCell(i));
		}
		System.out.println(myMap);
	}

}
