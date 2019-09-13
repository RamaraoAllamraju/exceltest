package screenshot.excelfinal.test;

import java.awt.Button;
import java.awt.Dimension;
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

import javax.imageio.ImageIO;
import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JPanel;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class MergedCode implements ActionListener {

	static int i = 0;
	static int row = 1;
	static int column = 1;
	static String excelPath = "C:\\Users\\rallamr\\Desktop\\Ramarao\\captured\\test1.xlsx";
	static String imgPath = "C:\\Users\\rallamr\\Desktop\\Ramarao\\captured\\test.jpg";

	MergedCode() throws IOException {
		createExcel(excelPath);
		JFrame frame = new JFrame("Insert Screenshot");
		JPanel panel = new JPanel();
		Button copyButton = new Button("Copy");
		Button doneButton = new Button("Done");

		copyButton.setPreferredSize(new Dimension(100, 30));
		doneButton.setPreferredSize(new Dimension(100, 30));

		panel.add(copyButton);
		panel.add(doneButton);

		frame.setSize(300, 200);
		frame.setContentPane(panel);
		frame.show();

		copyButton.addActionListener(this);
		doneButton.addActionListener(this);
	}

	public static void main(String[] args) throws IOException, UnsupportedFlavorException {
		// TODO Auto-generated method stub
		/*
		 * String path = "C:\\Users\\rallamr\\Desktop\\Ramarao\\captured\\test.xlsx";
		 * createExcel(path);
		 * 
		 * copyClipBoardInfo();
		 */
		new MergedCode();
	}

	private static void createExcel(String path) throws IOException {
		File file = new File(path);
		if (file.exists()) {
			System.out.println("File already Exists");
		} else {
			System.out.println("No file");
			XSSFWorkbook my_workbook = new XSSFWorkbook();
			XSSFSheet my_sheet = my_workbook.createSheet("TestSheet");
			
			
			FileOutputStream fileOS = new FileOutputStream(path);
			my_workbook.write(fileOS);
			fileOS.close();
		}
	}

	private static void copyClipBoardInfo() throws IOException, UnsupportedFlavorException {
		Toolkit toolkit = Toolkit.getDefaultToolkit();
		Clipboard clipboard = toolkit.getSystemClipboard();
		try {
			System.out.println("In String flavor");
			String s = (String) clipboard.getData(DataFlavor.stringFlavor);
			System.out.println("String is : " + s);
		} catch (UnsupportedFlavorException exception) {
			System.out.println("It is not a String flavor, checking for image");
			Image SrcFile = (Image) clipboard.getData(DataFlavor.imageFlavor);

			BufferedImage bi = (BufferedImage) SrcFile;
			System.out.println("Height :" + bi.getHeight() + "\n Width :" + bi.getWidth());
			// return bi;

			File DestFile = new File(imgPath);
			ImageIO.write((RenderedImage) SrcFile, "jpg", DestFile);
			pasteImageInExcel(false);
		}
		// return null;
	}

	private static void pasteImageInExcel(boolean flag) throws IOException {

		if (!flag) {
			
			XSSFWorkbook my_workbook = new XSSFWorkbook();
			XSSFSheet my_sheet = my_workbook.createSheet("MyLogo");
			InputStream my_banner_image = new FileInputStream(imgPath);
			byte[] bytes = org.apache.poi.util.IOUtils.toByteArray(my_banner_image);
			int my_picture_id = my_workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
			my_banner_image.close();
			XSSFDrawing drawing = my_sheet.createDrawingPatriarch();
			XSSFClientAnchor my_anchor = new XSSFClientAnchor();

			/*
			 * Get position
			 */

			my_anchor.setCol1(2);
			my_anchor.setRow1(1);
			XSSFPicture my_picture = drawing.createPicture(my_anchor, my_picture_id);
			my_picture.resize();
		}

		FileOutputStream out = new FileOutputStream(new File(excelPath));
		my_workbook.write(out);
		out.close();
	}

	private void pasteTextInExcel() {

	}

	private void calculateLength() {

	}

	public void actionPerformed(ActionEvent paramActionEvent) {
		// TODO Auto-generated method stub
		if (paramActionEvent.getActionCommand().toString().equals("Copy")) {
			System.out.println("Copied");
			try {
				copyClipBoardInfo();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (UnsupportedFlavorException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		} else if (paramActionEvent.getActionCommand().toString().equals("Done")) {
			System.out.println("Clicked on DONE");
			completed();
		}

	}

	private void completed() {
		// TODO Auto-generated method stub
		pasteImageInExcel(true);
	}
}
