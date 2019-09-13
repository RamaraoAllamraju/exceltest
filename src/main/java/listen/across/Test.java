package listen.across;

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
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.LogManager;
import java.util.logging.Logger;

import javax.imageio.ImageIO;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.JTextArea;

import org.jnativehook.GlobalScreen;
import org.jnativehook.NativeHookException;
import org.jnativehook.keyboard.NativeKeyEvent;
import org.jnativehook.keyboard.NativeKeyListener;

public class Test implements NativeKeyListener, ActionListener {

	JTextArea excelLoc = null;
	static JFileChooser chooser = new JFileChooser();
	static String excelPath = chooser.getCurrentDirectory().toString() + "\\test1.xlsx";

	static String imgPath = "C:\\Users\\rallamr\\Desktop\\Ramarao\\test.jpg";

	Test() {
		JFrame frame = new JFrame("Insert Screenshot");
		JPanel panel = new JPanel();

		excelLoc = new JTextArea(excelPath);

		JButton copyButton = new JButton("Copy");
		JButton doneButton = new JButton("Done");
		JButton selectExcelLocation = new JButton("Select Excel Location");

		excelLoc.setPreferredSize(new Dimension(250, 20));
		selectExcelLocation.setPreferredSize(new Dimension(160, 30));
		copyButton.setPreferredSize(new Dimension(75, 30));
		doneButton.setPreferredSize(new Dimension(75, 30));

		panel.add(excelLoc);
		panel.add(selectExcelLocation);

		panel.add(copyButton);
		panel.add(doneButton);

		frame.setSize(300, 200);
		frame.setContentPane(panel);
		frame.setAlwaysOnTop(true);
		frame.show();

		copyButton.addActionListener(this);
		doneButton.addActionListener(this);
		selectExcelLocation.addActionListener(this);
	}

	public static void main(String args[]) {
		LogManager.getLogManager().reset();
		Logger logger = Logger.getLogger(GlobalScreen.class.getPackage().getName());
		logger.setLevel(Level.OFF);

		try {
			GlobalScreen.registerNativeHook();
		} catch (NativeHookException ex) {
			System.err.println("There was a problem registering the native hook.");
			System.err.println(ex.getMessage());
			System.exit(1);
		}
		GlobalScreen.addNativeKeyListener(new Test());
		new Test();
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

	void copyImgFromClip() {
		Toolkit toolkit = Toolkit.getDefaultToolkit();
		Clipboard clipboard = toolkit.getSystemClipboard();
		Image SrcFile = null;
		try {
			SrcFile = (Image) clipboard.getData(DataFlavor.imageFlavor);
		} catch (Exception e2) {
			e2.printStackTrace();
		}
		BufferedImage bi = (BufferedImage) SrcFile;
		File DestFile = new File(imgPath);
		try {
			ImageIO.write((RenderedImage) SrcFile, "jpg", DestFile);
		} catch (IOException e1) {
			e1.printStackTrace();
		}
	}

	void copyStringFromClip() throws UnsupportedFlavorException, IOException {
		Toolkit toolkit = Toolkit.getDefaultToolkit();
		Clipboard clipboard = toolkit.getSystemClipboard();
		String s = (String) clipboard.getData(DataFlavor.stringFlavor);
		System.out.println("String is : " + s);
	}

	/*
	 * (non-Javadoc)
	 * @see
	 * org.jnativehook.keyboard.NativeKeyListener#nativeKeyPressed(org.jnativehook.
	 * keyboard.NativeKeyEvent) It GLOBALLY listens to the KeyBoard events
	 */
	public void nativeKeyPressed(NativeKeyEvent e) {
		// TODO Auto-generated method stub
		try {
			Thread.sleep(1000);
		} catch (InterruptedException e3) {
			// TODO Auto-generated catch block
			e3.printStackTrace();
		}
		if (NativeKeyEvent.getKeyText(e.getKeyCode()).equals("Print Screen")) {
			System.out.println("PrintScreen");
			copyImgFromClip();
		}
	}

	public void nativeKeyReleased(NativeKeyEvent arg0) {
		// TODO Auto-generated method stub

	}

	public void nativeKeyTyped(NativeKeyEvent arg0) {
		// TODO Auto-generated method stub

	}

	public void actionPerformed(ActionEvent paramActionEvent) {
		// TODO Auto-generated method stub
		if (paramActionEvent.getActionCommand().toString().equals("Copy")) {
			try {
				copyStringFromClip();
			} catch (UnsupportedFlavorException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
		else if (paramActionEvent.getActionCommand().toString().equals("Done")) {
			System.out.println("Clicked on DONE");
			try {
				completed();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		
		else if(paramActionEvent.getActionCommand().toString().equals("Select Excel Location"))
		{       
		    chooser.setCurrentDirectory(new java.io.File("."));
		    chooser.setDialogTitle("Select Excel");
		    chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
		    chooser.setAcceptAllFileFilterUsed(false); 
		    if (chooser.showOpenDialog(chooser) == JFileChooser.APPROVE_OPTION) {
		    	excelPath = chooser.getSelectedFile().toString() + "\\test1.xlsx";
		    	excelLoc.setText(excelPath);
		      System.out.println("getCurrentDirectory(): " 
		         +  chooser.getCurrentDirectory());
		      System.out.println("getSelectedFile() : " 
		         +  chooser.getSelectedFile());
		      }
		    else {
		      System.out.println("No Selection ");
		      }
		}
	}
}