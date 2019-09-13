package listen.across;

import java.awt.Dimension;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.util.EventListener;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.JTextArea;

public class ListenAcross implements EventListener, ActionListener  {
	JTextArea excelLoc = null;
	static JFileChooser chooser = new JFileChooser();
	static String excelPath = chooser.getCurrentDirectory().toString()+"\\test1.xlsx";

	
	public ListenAcross() {
		// TODO Auto-generated constructor stub
		
		JFrame frame = new JFrame("Insert Screenshot");
		JPanel panel = new JPanel();
				
		excelLoc = new JTextArea(excelPath);
		
		JButton copyButton = new JButton("Copy");
		JButton doneButton = new JButton("Done");
		JButton selectExcelLocation = new JButton("Select Excel Location");

		excelLoc.setPreferredSize(new Dimension(250,20));
		selectExcelLocation.setPreferredSize(new Dimension(160,30));
		copyButton.setPreferredSize(new Dimension(75, 30));
		doneButton.setPreferredSize(new Dimension(75, 30));
		
		//excelLoc.setBounds(200, 20, 100, 30);

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

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		new ListenAcross();
	}

	public void actionPerformed(ActionEvent arg0) {
		// TODO Auto-generated method stub
		
	}

}
