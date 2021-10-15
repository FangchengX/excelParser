import service.formygirl.SummaryService;

public class ParserApplication {
    public static void main(String[] args) throws Exception {

        String path = "C:\\Users\\kq644\\Desktop\\lyy\\2021.9.8学习强院考试成绩\\消毒灭菌考核试题-考试成绩.xls";
        String progressPath = "C:\\Users\\kq644\\Desktop\\lyy\\2021.9.8学习强院学习合格名单\\课程《消毒灭菌的那些事儿》的完成人员名单.xls";
        SummaryService summaryService = new SummaryService();
        summaryService.doExamSummary("消毒灭菌那些事儿", path, progressPath);
    }
}
