package service;

import data.RowDTO;
import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * Created by kq644 on 2019/12/9.
 */
public class Process {

    public static SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd");

    public static Map<String, List<RowDTO>> studentMap = new HashMap<>();
    public static Map<String, List<RowDTO>> techerMap = new HashMap<>();

    public void doProcess(String filePath) throws Exception {
        FileInputStream fis = new FileInputStream(filePath);
        Workbook workbook = WorkbookFactory.create(fis);
        int sheetNumber = workbook.getNumberOfSheets();
        for (int i = 1; i < sheetNumber; i++) {
            Sheet sheet = workbook.getSheetAt(i);
            int rowNumber = sheet.getLastRowNum() + 1;
            for (int j = 2; j < rowNumber; j++) {
                Row row = sheet.getRow(j);
                int columnNumber = row.getLastCellNum() + 1;
                for (int k = 0; k < columnNumber; k++) {
                    if (Objects.nonNull(row.getCell(k))) {
                        Cell cell = row.getCell(k);
                        switch (cell.getCellTypeEnum()) {
                            case NUMERIC:
                                System.out.print(cell.getNumericCellValue());
                                break;
                            case FORMULA:
                            default:
                                System.out.print(row.getCell(k).getStringCellValue());
                                break;
                        }
                    }
                }
                System.out.println();
            }
        }
    }

    public void doProcessSheet(String filePath) throws Exception{
        FileInputStream fis = new FileInputStream(filePath);
        Workbook workbook = WorkbookFactory.create(fis);
        processSheet(workbook.getSheetAt(1));
    }

    public void processSheet(Sheet sheet) {
        for (int i = 0; i < sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            processRow(row);
        }
    }

    public void processRow(Row row) {
        //TODO 新建一个rowDTO， 根据日期放入对应的map中
        for (int i = 0; i < row.getLastCellNum(); i ++) {
            String temp = getCellValue(row.getCell(i), i);
        }
        System.out.println();
    }




    public String getCellValue(Cell cell, int i) {
        if (cell != null) {
            switch (cell.getCellTypeEnum()) {
                case BOOLEAN:
                    System.out.print(cell.getBooleanCellValue());
                    break;
                case NUMERIC:
                    if (i==0) {
                        System.out.print(format.format(cell.getDateCellValue()));
                    } else {
                        System.out.print(cell.getNumericCellValue());
                    }
                    break;
                case STRING:
                    System.out.print(cell.getStringCellValue());
                    break;
                case BLANK:
                    break;
                case ERROR:
                    System.out.print(cell.getErrorCellValue());
                    break;

                // CELL_TYPE_FORMULA will never occur
                case FORMULA:
                    break;
                default:
                    break;
            }
        }
        System.out.print(",");
        return "123";
    }


}
