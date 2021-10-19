package view;

import javax.swing.*;
import java.io.OutputStream;
import java.io.PrintStream;
import java.util.Objects;

/**
 * @author fangcheng
 * @since 10/19/21
 */
public class MyPrintStream extends PrintStream {
    private JTextArea jTextArea;

    public MyPrintStream(OutputStream outputStream, JTextArea jTextArea) {
        super(outputStream);
        this.jTextArea = jTextArea;
    }

    @Override
    public void println(String message) {

        if (Objects.isNull(message)) {
            return;
        }
        SwingUtilities.invokeLater(() -> jTextArea.append(message + "\r\n"));
    }
}
