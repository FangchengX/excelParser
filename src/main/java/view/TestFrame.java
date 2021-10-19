package view;

import javax.swing.*;
import java.awt.*;

/**
 * @author fangcheng
 * @since 10/19/21
 */
public class TestFrame extends JFrame {
    JTextArea jTextArea;

    public TestFrame(JTextArea jTextArea) throws HeadlessException {
        this.jTextArea = jTextArea;
        MyPrintStream myPrintStream = new MyPrintStream(System.out, jTextArea);
        System.setOut(myPrintStream);
        JFrame frame = new JFrame("Testing");
        frame.add(logButton(), BorderLayout.WEST);
        frame.add(initLogArea(), BorderLayout.EAST);
        frame.setBounds(300, 200, 600, 300);
        frame.setVisible(true);
        frame.setDefaultCloseOperation(EXIT_ON_CLOSE);
    }

    private JScrollPane initLogArea() {
        JScrollPane jScrollPane = new JScrollPane();
        jScrollPane.setPreferredSize(new Dimension(300, 300));
        jScrollPane.setBounds(20, 20, 100, 50);
        jTextArea.setEditable(false);
        jScrollPane.setViewportView(jTextArea);
        return jScrollPane;
    }

    private JButton logButton() {
        JButton button = new JButton("summary");
        button.setBounds(150, 100, 100, 20);
        button.addActionListener(e -> {
            System.out.println("this is a log");
        });
        return button;
    }
}
