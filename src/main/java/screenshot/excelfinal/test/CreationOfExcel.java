package screenshot.excelfinal.test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreationOfExcel {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		File f = new File("C:\\Users\\rallamr\\Desktop\\Ramarao\\Personal\\insert_png_xlsx_example2.xlsx");
		if (f.exists()) {
			System.out.println("File present");
		} else {
			System.out.println("No file");
			XSSFWorkbook my_workbook = new XSSFWorkbook();
			XSSFSheet my_sheet = my_workbook.createSheet("MySheet1");
			FileOutputStream file = new FileOutputStream(
					"C:\\Users\\rallamr\\Desktop\\Ramarao\\Personal\\insert_png_xlsx_example2.xlsx");
			my_workbook.write(file);
			file.close();
		}
	}
	
	/*
	 * Not required
	 */
	public static void copyTemplate() throws IOException
	{
		System.out.println(System.getProperty("user.name"));
		File Source = new File(".\\Files\\Sample Excel.xlsx");
		/*File Destination = new File(path);
		FileUtils.copyFile(Source, Destination);*/
	}

}
