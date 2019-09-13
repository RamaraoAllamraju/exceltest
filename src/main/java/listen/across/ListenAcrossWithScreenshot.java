package listen.across;

import java.awt.Image;
import java.awt.Toolkit;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.UnsupportedFlavorException;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.image.BufferedImage;
import java.awt.image.RenderedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.logging.Level;
import java.util.logging.LogManager;
import java.util.logging.Logger;

import javax.imageio.ImageIO;
import javax.swing.JFileChooser;
import javax.swing.JTextArea;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jnativehook.GlobalScreen;
import org.jnativehook.NativeHookException;
import org.jnativehook.keyboard.NativeKeyEvent;
import org.jnativehook.keyboard.NativeKeyListener;

public class ListenAcrossWithScreenshot implements ActionListener,NativeKeyListener {

	static JFileChooser chooser = new JFileChooser();
	JTextArea excelLoc = null;
	
	static int i = 0;
	static int row = 1;
	static int column = 1;
	static String excelPath = "C:\\Users\\rallamr\\Desktop\\Ramarao\\test1.xlsx";
	static String imgPath = "C:\\Users\\rallamr\\Desktop\\Ramarao\\test.jpg";

	static int height = 0;
	
	
	static XSSFWorkbook my_workbook = null;
	static XSSFSheet my_sheet = null;
	
	
	
	
	public ListenAcrossWithScreenshot() {
		// TODO Auto-generated constructor stub
	}

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		LogManager.getLogManager().reset();

		// Get the logger for "org.jnativehook" and set the level to off.
		Logger logger = Logger.getLogger(GlobalScreen.class.getPackage().getName());
		logger.setLevel(Level.OFF);
		
		new ListenAcrossWithScreenshot();
		try {
			GlobalScreen.registerNativeHook();
		}
		catch (NativeHookException ex) {
			System.err.println("There was a problem registering the native hook.");
			System.err.println(ex.getMessage());

			System.exit(1);
		}

		GlobalScreen.addNativeKeyListener(new Test());
	}
	/*
	 * Exit code - Deleting the created Image file on disk & closing the UI
	 */

	private void completed() throws IOException {
		File f = new File(imgPath);
		f.deleteOnExit();
		// f.delete();
		System.exit(0);
	}

	/*
	 * Used for getting the information in the Clipboard ( Supports : String and
	 * Image Flavour )
	 */
	private static void copyClipBoardInfo() throws IOException, UnsupportedFlavorException {
		Toolkit toolkit = Toolkit.getDefaultToolkit();
		Clipboard clipboard = toolkit.getSystemClipboard();
		try {
			System.out.println("In String flavor");
			String s = (String) clipboard.getData(DataFlavor.stringFlavor);
			pasteTextInExcel(s);
			System.out.println("String is : " + s);
		} catch (UnsupportedFlavorException exception) {
			System.out.println("It is not a String flavor, checking for image");
			Image SrcFile = (Image) clipboard.getData(DataFlavor.imageFlavor);

			BufferedImage bi = (BufferedImage) SrcFile;
			System.out.println("Height :" + bi.getHeight() + "\n Width :" + bi.getWidth());
			height = bi.getHeight();
			
			
			File DestFile = new File(imgPath);
			ImageIO.write((RenderedImage) SrcFile, "jpg", DestFile);
			pasteImageInExcel();
		}
	}

	/*
	 * Creates the Excel
	 */
	private static void createExcel(String path) throws IOException {
		/*File file = new File(path);
		if (file.exists()) {
			System.out.println("File already Exists");
		} else {*/
			System.out.println("No file");
			my_workbook = new XSSFWorkbook();
			my_sheet = my_workbook.createSheet("TestSheet");

			FileOutputStream fileOS = new FileOutputStream(path);
			my_workbook.write(fileOS);
			fileOS.close();
		//}
	}

	/*
	 * Creates a FileOutPutStream object and writes data to Excel
	 */

	private static void writeToExcel() throws IOException {
		//FileOutputStream out = new FileOutputStream(excelPath);
		my_workbook.write(new FileOutputStream(excelPath));
		//out.close();
		System.out.println("File Closed");
	}

	/*
	 * This function helps for getting the new Anchor Point to paste the image in
	 * Excel
	 */
	private static XSSFClientAnchor getAnchorPoint() {
		XSSFClientAnchor my_anchor = new XSSFClientAnchor();
		my_anchor.setCol1(2);
		my_anchor.setRow1(row);
		System.out.println("Row is :" + row);
		return my_anchor;
	}

	/*
	 * This function will gives the next row where the data needs to be pasted -
	 * Called from pasteTextInExcel()
	 */
	private static void getNextPosition() {
		System.out.println("Last Row is " + my_sheet.getLastRowNum());
		row = my_sheet.getLastRowNum() + 5;
	}

	/*
	 * This function will gives the next row where the data needs to be pasted -
	 * Called from pasteImageInExcel()
	 */
	private static void getNextPosition(XSSFPicture my_picture) {
		int defaultCellHeight = 20;
		System.out.println("Current Row :"+row);
		row = row + (height/defaultCellHeight) +5;
		System.out.println("Latest Row will be :" + row);
	}

	/*
	 * Will open the Excel file
	 */
	private static void openExcel() throws FileNotFoundException, IOException {
		my_workbook = new XSSFWorkbook(new FileInputStream(excelPath));
		my_sheet = my_workbook.getSheet("TestSheet");
	}

	/*
	 * Pastes an Image into excel
	 */

	private static void pasteImageInExcel() throws IOException {
		System.out.println("In PasteImageInExcel");
		openExcel();
		InputStream my_banner_image = new FileInputStream(imgPath);
		byte[] bytes = org.apache.poi.util.IOUtils.toByteArray(my_banner_image);
		int my_picture_id = my_workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
		my_banner_image.close();
		XSSFDrawing drawing = my_sheet.createDrawingPatriarch();
		XSSFPicture my_picture = drawing.createPicture(getAnchorPoint(), my_picture_id);
		getNextPosition(my_picture);
		my_picture.resize();
		System.out.println("Excel Path :" + excelPath);
		writeToExcel();
	}

	/*
	 * Pastes the text into excel
	 */
	private static void pasteTextInExcel(String s) throws FileNotFoundException, IOException {
		openExcel();
		System.out.println("Value :" + s);
		System.out.println("Row :" + row + "\n Column :" + column);
		System.out.println("Sheet is " + my_sheet.getSheetName());

		Cell my_cell = null;
		Row my_row = null;
		
		String myRowsArray[] = s.split("\n");
		int numberOfRows = myRowsArray.length;
		System.out.println("Number of Rows :"+numberOfRows);
		int numberOfColumns = myRowsArray[0].split("\t").length;
		if (numberOfRows > 0 && numberOfColumns > 1) {
			System.out.println("Number of Rows :" + numberOfRows + "\nNumber of Columns:" + numberOfColumns);
			String[][] data = new String[numberOfRows][numberOfColumns];

			for (int rowItr = 0; rowItr < numberOfRows; rowItr++) {
				data[rowItr] = myRowsArray[rowItr].split("\t");
			}
			/*
			 * Just for Log purpose
			 */
			for (int j = 0; j < numberOfRows; j++)
				for (int k = 0; k < numberOfColumns; k++)
					System.out.println("Data is data[" + j + "][" + k + "]:" + data[j][k]);
			for (int rowItr = 0; rowItr < numberOfRows; rowItr++, row++) {
				my_row = my_sheet.createRow(row);
				for (int columnItr = 1; columnItr < numberOfColumns; columnItr++) {
					my_cell = my_row.createCell(columnItr);
					System.out.println("Row ITR - " + row + "\nColumn ITR - " + columnItr);
					my_cell.setCellValue(data[rowItr][columnItr]);
				}
			}
		}
		else
		{
			my_row = my_sheet.getRow(row);
			my_cell = my_row.createCell(column);
			my_cell.setCellValue(s);
		}
		System.out.println("Array is :" + myRowsArray.length);
		getNextPosition();
		writeToExcel();
	}

	public void nativeKeyPressed(NativeKeyEvent e) {
		// TODO Auto-generated method stub
		//System.out.println("Key Pressed: " + NativeKeyEvent.getKeyText(e.getKeyCode()));
		if(NativeKeyEvent.getKeyText(e.getKeyCode()).equals("Print Screen"))
		{
			System.out.println("PrintScreen");
			System.out.println("Here");
			Toolkit toolkit = Toolkit.getDefaultToolkit();
			Clipboard clipboard = toolkit.getSystemClipboard();
			
			try {
				System.out.println("In String flavor");
				String s = (String) clipboard.getData(DataFlavor.stringFlavor);
				pasteTextInExcel(s);
				System.out.println("String is : " + s);
			} catch (UnsupportedFlavorException exception) {
				System.out.println("It is not a String flavor, checking for image");
				Image SrcFile=null;
				try {
					SrcFile = (Image) clipboard.getData(DataFlavor.imageFlavor);
				} catch (UnsupportedFlavorException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				} catch (IOException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}

				BufferedImage bi = (BufferedImage) SrcFile;
				System.out.println("Height :" + bi.getHeight() + "\n Width :" + bi.getWidth());
				height = bi.getHeight();
				
				
				File DestFile = new File(imgPath);
				try {
					ImageIO.write((RenderedImage) SrcFile, "jpg", DestFile);
				} catch (IOException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
				//pasteImageInExcel();
			} catch (IOException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}
			
			
			
			
			
			/*if(! new File(excelPath).exists())
				try {
					createExcel(excelPath);
				} catch (IOException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
			System.out.println("Copied");
			try {
				copyClipBoardInfo();
			} catch (IOException ex) {
				System.out.println("In IOExcepetion");
				ex.printStackTrace();
			} catch (UnsupportedFlavorException ex) {
				System.out.println("In UnsupportedFlavourException");
				ex.printStackTrace();
			}*/
		
		}
	}

	public void nativeKeyReleased(NativeKeyEvent e) {
		// TODO Auto-generated method stub
		System.out.println("Key Released: " + NativeKeyEvent.getKeyText(e.getKeyCode()));

	}

	public void nativeKeyTyped(NativeKeyEvent e) {
		// TODO Auto-generated method stub
		System.out.println("Key Typed: " + NativeKeyEvent.getKeyText(e.getKeyCode()));

	}

	public void actionPerformed(ActionEvent e) {
		// TODO Auto-generated method stub
		//System.out.println("Key Pressed: " + NativeKeyEvent.getKeyText(e.getKeyCode()));

	}

}
