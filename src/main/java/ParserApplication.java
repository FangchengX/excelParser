import com.google.common.collect.Lists;
import lombok.AllArgsConstructor;
import lombok.Data;
import service.exam.ExamService;
import view.MyPrintStream;
import view.SummaryFrame;

import javax.swing.*;
import java.awt.*;
import java.util.List;

public class ParserApplication {
    public static void main(String[] args) throws Exception {
        forMyGirl();
    }

    public static void forMySister() throws Exception {
        String studentSheet = "C:\\Users\\kq644\\Desktop\\qiandao\\2021\\考场签到表\\在校生名单20-21-1.xlsx";
        String examSheet = "C:\\Users\\kq644\\Desktop\\qiandao\\2021\\考场签到表\\2021-2022-1学期考试安排.xls";
        String outputPath = "C:\\Users\\kq644\\Desktop\\qiandao\\2021\\考场签到表\\result.xlsx";
        ExamService examService = new ExamService();
        examService.parseData(studentSheet, examSheet, outputPath);
    }

    public static void forMyGirl() {
        @Data
        @AllArgsConstructor
        class AppClassInfo {
            String name;
            String resultPath;
            String progressPath;
        }

        String resultFolder = "C:\\Users\\kq644\\Desktop\\lyy\\1026学习强院\\考核成绩\\";
        String progressFolder = "C:\\Users\\kq644\\Desktop\\lyy\\1026学习强院\\学习进度通过\\";

        List<AppClassInfo> appClasses = Lists.newArrayList(
            new AppClassInfo("消毒灭菌那些事儿", resultFolder + "消毒灭菌考核试题-所有学员.xls", progressFolder + "课程《消毒灭菌的那些事儿》的完成人员名单.xls"),
//                new AppClassInfo("环境卫生学检测", resultFolder + "", progressFolder+""),
            new AppClassInfo("医疗机构门急诊医院感染管理规范", resultFolder + "医疗机构门急诊医院感染管理规范-所有学员.xls", progressFolder + "课程《医疗机构门急诊医院感染管理规范》的完成人员名单.xls"),
            new AppClassInfo("医疗机构预防与感染控制基本制度", resultFolder + "医疗机构预防与感染控制基本制度-所有学员.xls", progressFolder + "课程《医疗机构预防与感染控制基本制度》的完成人员名单.xls"),
            new AppClassInfo("软式内镜清洗消毒", resultFolder + "软式内镜清洗消毒考核-所有学员.xls", progressFolder + "课程《软式内镜清洁消毒培训》的完成人员名单.xls"),
            new AppClassInfo("医疗机构环境表面清洁与消毒", resultFolder + "医疗机构环境清洁与消毒试题-所有学员.xls", progressFolder + "课程《医疗机构环境表面清洁与消毒》的完成人员名单.xls"),
            new AppClassInfo("生物安全柜", null, progressFolder + "课程《生物安全柜使用》的完成人员名单.xls"),
            new AppClassInfo("病区感染规范", resultFolder + "病区医院感染规范考题-所有学员.xls", progressFolder + "课程《病区医院感染规范》的完成人员名单.xls"),
            new AppClassInfo("实验室生物安全管理和防护", resultFolder + "实验室生物安全管理和防护-所有学员.xls", progressFolder + "课程《实验室生物安全管理和防护》的完成人员名单.xls"),
            new AppClassInfo("临床微生物检测标本采集和运送规范", resultFolder + "临床微生物检测标本采集和运送规范-所有学员.xls", progressFolder + "课程《临床微生物检测标本采集和运送规范》的完成人员名单.xls"),
            new AppClassInfo("实验室生物安全相关法律、法规及要点", resultFolder + "实验室生物安全相关法律、法规及要点-所有学员.xls", progressFolder + "课程《实验室生物安全相关法律、法规及要点》的完成人员名单.xls"),
            new AppClassInfo("手卫生", resultFolder + "医务人员手卫生培训考核-所有学员.xls", progressFolder + "课程《医务人员手卫生培训》的完成人员名单.xls"),
            new AppClassInfo("多重耐药菌医院感染预防与控制措施", resultFolder + "多重耐药医院感染预防与控制措施-所有学员.xls", progressFolder + "课程《多重耐药菌医院感染预防与控制措施》的完成人员名单.xls"),
            new AppClassInfo("2021年新馆肺炎诊疗与防控培训", resultFolder + "2021年新冠肺炎诊疗与防控培训试题-所有学员.xls", progressFolder + "课程《2021年新冠疫情诊疗与防控培训》的完成人员名单.xls"),
            new AppClassInfo("医疗废物管理", resultFolder + "医疗废弃物管理要求-所有学员.xls", progressFolder + "课程《医疗废弃物管理要求》的完成人员名单.xls"),
            new AppClassInfo("2021实验室生物安全防护培训", resultFolder + "实验室生物安全防护培训考核试题-所有学员.xls", progressFolder + "课程《实验室生物安全防护培训》的完成人员名单.xls")
        );
//        SummaryService summaryService = new SummaryService();
//        summaryService.parseMembers();
//        String filePath = "C:\\Users\\kq644\\Desktop\\lyy\\线上课_1.消毒灭菌那些事儿.xlsx";
//        summaryService.doDingdingSummary("消毒灭菌那些事儿", filePath);
//        appClasses.forEach(data -> {
//            try {
//                summaryService.doExamSummary(data.getName(), data.getResultPath(), data.getProgressPath());
//            } catch (Exception e) {
//                e.printStackTrace();
//            }
//        });
        JTextArea jTextArea = new JTextArea();
        MyPrintStream myPrintStream = new MyPrintStream(System.out, jTextArea);
        System.setOut(myPrintStream);
        Frame frame = new SummaryFrame(jTextArea);
    }

}
