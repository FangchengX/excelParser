import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.FileInputStream;
import java.util.Objects;

public class ParserApplication {
    public static void main(String[] args) throws Exception{
        FileInputStream fis = new FileInputStream("201911食堂出入库明细 11月（上交） .xls");
        Workbook workbook = WorkbookFactory.create(fis);
        int sheetNumber = workbook.getNumberOfSheets();
        for (int i=0;i<sheetNumber;i++) {
            Sheet sheet = workbook.getSheetAt(i);
            int rowNumber = sheet.getLastRowNum()+1;
            for (int j=0;j<rowNumber;j++) {
                Row row = sheet.getRow(j);
                int columnNumber = row.getLastCellNum()+1;
                for (int k=0;i<columnNumber;k++) {
                    if (Objects.nonNull(row.getCell(k))) {
                        System.out.print(row.getCell(k).getStringCellValue());
                    }
                }
                System.out.println();
            }
        }
    }
}
