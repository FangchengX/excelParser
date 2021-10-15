package view;

import com.google.common.collect.Lists;
import service.formygirl.SummaryService;

import javax.swing.*;
import java.awt.*;
import java.util.List;

/**
 * @author fangcheng
 * @since 10/15/21
 */
public class SummaryFrame extends JFrame {

    public static final List<String> CLASS_NAMES = Lists.newArrayList("1", "2");
    JComboBox<String> cmb = new JComboBox<>();    //创建JComboBox
    String className;
    String filePath;

    public SummaryFrame() throws HeadlessException {
        JFrame frame = new JFrame("Java下拉列表组件示例");
        JPanel jp = new JPanel();
        jp.setLayout(null);
        addClassSelector(jp);
        addFileSelector(jp);
        addSummaryButton(jp);
        frame.add(jp);
        frame.setBounds(300, 200, 600, 300);
        frame.setVisible(true);
        frame.setDefaultCloseOperation(EXIT_ON_CLOSE);
    }

    /**
     * @return
     */
    private void addClassSelector(JPanel jp) {
        JLabel label = new JLabel("class name：");    //创建标签
        label.setBounds(30, 60, 150, 30);

        cmb.addItem("--请选择--");    //向下拉列表中添加一项
        cmb.setEditable(true);
        cmb.setPreferredSize(new Dimension(400, 20));
        cmb.setBounds(130, 65, 400, 20);
        CLASS_NAMES.forEach(cmb::addItem);
        jp.add(label);
        jp.add(cmb);
    }

    private void addFileSelector(JPanel panel) {
        JLabel label = new JLabel("file path:");
        label.setBounds(30, 100, 150, 30);
        JTextField jtf = new JTextField(25);
        jtf.setBounds(130, 105, 350, 20);
        JButton button = new JButton("浏览");
        button.setBounds(480, 105, 80, 20);
        button.addActionListener(e -> {
            JFileChooser fc = new JFileChooser();
            int val = fc.showOpenDialog(null);    //文件打开对话框
            if (val == JFileChooser.APPROVE_OPTION) {
                //正常选择文件
                jtf.setText(fc.getSelectedFile().toString());
                filePath = jtf.getText();
            } else {
                //未正常选择文件，如选择取消按钮
                jtf.setText("未选择文件");
                filePath = null;
            }
        });
        panel.add(label);
        panel.add(jtf);
        panel.add(button);
    }

    private void addSummaryButton(JPanel panel) {
        JButton button = new JButton("summary");
        button.setBounds(250, 150, 100, 20);
        panel.add(button);
        button.addActionListener(e -> {
            className = cmb.getSelectedItem().toString();
            SummaryService summaryService = new SummaryService();
            try {
                summaryService.doDingdingSummary(className, filePath);
            } catch (InterruptedException ex) {
                ex.printStackTrace();
            }
        });
    }
}
