import org.apache.poi.ss.usermodel.*;
import service.Process;

import java.io.FileInputStream;
import java.util.Objects;

public class ParserApplication {
    public static void main(String[] args) throws Exception{

        String path = "201911食堂出入库明细 11月（上交） .xls";
        Process process = new Process();
        process.doProcessSheet(path);
    }
}
