package screenshot.excelfinal.test;

import java.awt.Image;
import java.awt.Toolkit;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.UnsupportedFlavorException;
import java.awt.image.BufferedImage;
import java.awt.image.RenderedImage;
import java.io.File;
import java.io.IOException;

import javax.imageio.ImageIO;

public class CopyClipBoardContent {
static int i = 0;
	public static void main(String[] args) throws IOException, UnsupportedFlavorException {
		// TODO Auto-generated method stub
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
			//return bi;
			
			File DestFile = new File("C:\\Users\\rallamr\\Desktop\\Ramarao\\captured\\test" + i + ".jpg");
			ImageIO.write((RenderedImage) SrcFile, "jpg", DestFile);
			i++;
			
		}
		//return null;
	}

}
