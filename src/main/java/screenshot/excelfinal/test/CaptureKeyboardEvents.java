package screenshot.excelfinal.test;

import java.awt.event.ActionEvent;

import javax.swing.AbstractAction;
import javax.swing.Action;
import javax.swing.ActionMap;
import javax.swing.InputMap;
import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.KeyStroke;

public class CaptureKeyboardEvents {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		/*
		 * JFrame frame = new JFrame("Test");
		 * frame.setSize(500, 500);
		 * JPanel panel = new JPanel();
		 * frame.add(panel);
		 * frame.show();
		 */

		String ACTION_KEY = "theAction";
		JFrame frame = new JFrame("KeyStroke Sample");
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		JButton buttonA = new JButton("Press 'control alt 7'");
		Action actionListener = new AbstractAction() {
			public void actionPerformed(ActionEvent actionEvent) {
				JButton source = (JButton) actionEvent.getSource();
				System.out.println(source.getText());
			}
		};
		KeyStroke controlAlt7 = KeyStroke.getKeyStroke("control alt 7");
		InputMap inputMap = buttonA.getInputMap();
		inputMap.put(controlAlt7, ACTION_KEY);
		ActionMap actionMap = buttonA.getActionMap();
		actionMap.put(ACTION_KEY, actionListener);
		frame.add(buttonA);
		frame.setSize(400, 200);
		frame.setVisible(true);
	}

}
