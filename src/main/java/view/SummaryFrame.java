package view;

import com.google.common.collect.Lists;
import service.formygirl.SummaryService;

import javax.swing.*;
import java.awt.*;
import java.io.IOException;
import java.util.List;

/**
 * @author fangcheng
 * @since 10/15/21
 */
public class SummaryFrame extends JFrame {

    public static final List<String> CLASS_NAMES = Lists.newArrayList(
            "消毒灭菌那些事儿",
            "环境卫生学检测",
            "医疗机构门急诊医院感染管理规范",
            "医疗机构预防与感染控制基本制度",
            "软式内镜清洗消毒",
            "医疗机构环境表面清洁与消毒",
            "生物安全柜",
            "病区感染规范",
            "实验室生物安全管理和保护",
            "临床微生物检测标本采集和运送规范",
            "实验室生物安全相关法律、法规及要点",
            "手卫生",
            "多重耐药菌医院感染预防与控制措施",
            "2021年新馆肺炎诊疗与防控培训",
            "三大导管的预防与控制",
            "手术部位感染预防与控制",
            "医疗废物管理"
    );
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
            } catch (IOException ex) {
                ex.printStackTrace();
            }
        });
    }
}
