package view;

import com.google.common.collect.Lists;
import service.formygirl.SummaryService;

import javax.swing.*;
import java.awt.*;
import java.io.File;
import java.util.List;

/**
 * @author fangcheng
 * @since 10/15/21
 */
public class SummaryFrame extends JFrame {
    JTextArea jTextArea;

    SummaryService summaryService;

    protected static final List<String> CLASS_NAMES = Lists.newArrayList(
        "消毒灭菌那些事儿",
        "环境卫生学检测",
        "医疗机构门急诊医院感染管理规范",
        "医疗机构预防与感染控制基本制度",
        "软式内镜清洗消毒",
        "医疗机构环境表面清洁与消毒",
        "生物安全柜",
        "病区感染规范",
        "实验室生物安全管理和防护",
        "临床微生物检测标本采集和运送规范",
        "实验室生物安全相关法律、法规及要点",
        "2021实验室生物安全防护培训",
        "手卫生",
        "多重耐药菌医院感染预防与控制措施",
        "2021年新冠肺炎诊疗与防控培训",
        "三大导管的预防与控制",
        "手术部位感染预防与控制",
        "医疗废物管理"
    );
    JComboBox<String> cmb = new JComboBox<>();    //创建JComboBox
    String className;
    String filePath;
    String outputFolder;

    public SummaryFrame(JTextArea jTextArea) throws HeadlessException {
        summaryService = new SummaryService();
        this.jTextArea = jTextArea;
        JFrame frame = new JFrame("FOR MY GIRL");
        JPanel jp = new JPanel();
        jp.setLayout(null);
        addClassSelector(jp);
        addFileSelector(jp);
        addOutputFolderSelector(jp);
        addSummaryButton(jp);
        frame.add(jp);
        frame.add(initLogArea(), BorderLayout.SOUTH);
        frame.setBounds(300, 200, 600, 400);
        frame.setVisible(true);
        frame.setDefaultCloseOperation(EXIT_ON_CLOSE);
    }

    private JScrollPane initLogArea() {
        JScrollPane jScrollPane = new JScrollPane();
        jScrollPane.setPreferredSize(new Dimension(600, 150));
        jScrollPane.setBounds(20, 20, 100, 50);
        jTextArea.setEditable(false);
        jScrollPane.setViewportView(jTextArea);
        return jScrollPane;
    }

    /**
     * @return
     */
    private void addClassSelector(JPanel jp) {
        JLabel label = new JLabel("课程名：");    //创建标签
        label.setBounds(30, 60, 150, 30);

        cmb.addItem("--请选择--");    //向下拉列表中添加一项
        cmb.setEditable(true);
        cmb.setPreferredSize(new Dimension(400, 20));
        cmb.setBounds(130, 65, 400, 20);
        CLASS_NAMES.forEach(cmb::addItem);
        jp.add(label);
        jp.add(cmb);
    }

    private void addOutputFolderSelector(JPanel panel) {
        JLabel label = new JLabel("结果存放文件夹:");
        label.setBounds(30, 140, 150, 30);
        JTextField jtf = new JTextField(25);
        jtf.setBounds(130, 145, 350, 20);
        JButton button = new JButton("浏览");
        button.setBounds(480, 145, 80, 20);
        button.addActionListener(e -> {
            JFileChooser fc = new JFileChooser();
            fc.setCurrentDirectory(new File("C:\\Users\\kq644\\Desktop\\lyy"));
            fc.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
            int val = fc.showOpenDialog(null);    //文件打开对话框
            if (val == JFileChooser.APPROVE_OPTION) {
                //正常选择文件
                jtf.setText(fc.getSelectedFile().toString());
                outputFolder = jtf.getText();
            } else {
                //未正常选择文件，如选择取消按钮
                jtf.setText("请选择文件夹");
                outputFolder = null;
            }
        });
        panel.add(label);
        panel.add(jtf);
        panel.add(button);
    }


    private void addFileSelector(JPanel panel) {
        JLabel label = new JLabel("输入文件:");
        label.setBounds(30, 100, 150, 30);
        JTextField jtf = new JTextField(25);
        jtf.setBounds(130, 105, 350, 20);
        JButton button = new JButton("浏览");
        button.setBounds(480, 105, 80, 20);
        button.addActionListener(e -> {
            JFileChooser fc = new JFileChooser();
            fc.setCurrentDirectory(new File("C:\\Users\\kq644\\Desktop\\lyy"));
            int val = fc.showOpenDialog(null);    //文件打开对话框
            if (val == JFileChooser.APPROVE_OPTION) {
                //正常选择文件
                jtf.setText(fc.getSelectedFile().toString());
                filePath = jtf.getText();
            } else {
                //未正常选择文件，如选择取消按钮
                jtf.setText("请选择钉钉结果文件");
                filePath = null;
            }
        });
        panel.add(label);
        panel.add(jtf);
        panel.add(button);
    }

    private void addSummaryButton(JPanel panel) {
        JButton button = new JButton("开始统计");
        button.setBounds(250, 180, 100, 20);
        panel.add(button);
        button.addActionListener(e -> {
            className = cmb.getSelectedItem().toString();
            try {
                summaryService.doDingdingSummary(className, filePath, outputFolder);
            } catch (Exception ex) {
                ex.printStackTrace();
            }
        });
    }
}
