package exceltest.exceltest;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class GetNextPosition {
	static String excelPath = "C:\\Users\\rallamr\\Desktop\\Ramarao\\captured\\Position Test.xlsx";

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		//XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(excelPath));
		/*XSSFWorkbook wb = new XSSFWorkbook(excelPath);
		XSSFSheet sheet = wb.getSheetAt(0);*/
				
		XSSFWorkbook my_workbook = new XSSFWorkbook();
        XSSFSheet my_sheet = my_workbook.createSheet("MyLogo");         
        InputStream my_banner_image = new FileInputStream("C:\\Users\\rallamr\\Desktop\\Ramarao\\Personal\\Ramarao.PNG");
        /* Convert picture to be added into a byte array */
        byte[] bytes = org.apache.poi.util.IOUtils.toByteArray(my_banner_image);
        /* Add Picture to Workbook, Specify picture type as PNG and Get an Index */
        int my_picture_id = my_workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
        /* Close the InputStream. We are ready to attach the image to workbook now */
        my_banner_image.close();                
        /* Create the drawing container */
        XSSFDrawing drawing = my_sheet.createDrawingPatriarch();
        /* Create an anchor point */
        XSSFClientAnchor my_anchor = new XSSFClientAnchor();
        /* Define top left corner, and we can resize picture suitable from there */
        my_anchor.setCol1(2);
        my_anchor.setRow1(1);           
        /* Invoke createPicture and pass the anchor point and ID */
        XSSFPicture  my_picture = drawing.createPicture(my_anchor, my_picture_id);
        System.out.println("Size Column :"+my_picture.getPreferredSize().getCol2());
        System.out.println("Size Row :"+my_picture.getPreferredSize().getRow2());

        /* Call resize method, which resizes the image */
        my_picture.resize();            
        /* Write changes to the workbook */
        FileOutputStream out = new FileOutputStream(new File("C:\\Users\\rallamr\\Desktop\\Ramarao\\Personal\\insert_png_xlsx_example.xlsx"));
        my_workbook.write(out);
        out.close();
	}

}
