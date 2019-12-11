
import service.Process;

public class ParserApplication {
    public static void main(String[] args) throws Exception{

        String path = "201911食堂出入库明细 11月（上交） .xls";
        Process process = new Process();
        process.doProcess(path);
        process.output();
    }
}
