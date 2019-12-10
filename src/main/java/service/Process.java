package service;

import data.RowDTO;
import org.apache.commons.compress.utils.Lists;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
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

    /**
     * 7，8，9 学生
     * 10，11，12 老师
     * 3 unit
     * 1/2 name
     *
     * @param row
     */
    public void processRow(Row row) {
        //TODO 新建一个rowDTO， 根据日期放入对应的map中

        String date;
        if (row.getLastCellNum() != 0 && row.getCell(0).getCellType() == CellType.NUMERIC) {
            date = format.format(row.getCell(0).getDateCellValue());
        } else {
            return;
        }
        RowDTO studentDTO = new RowDTO();
        RowDTO teacherDTO = new RowDTO();
        List<RowDTO> studentList = studentMap.computeIfAbsent(date, (unused) -> Lists.newArrayList());
        List<RowDTO> teacherList = techerMap.computeIfAbsent(date, (unused) -> Lists.newArrayList());
        int number = row.getRowNum();
        String name;
        if (Objects.isNull(row.getCell(1)) || Objects.equals(row.getCell(1).getStringCellValue(), "")) {
            name = row.getCell(2).getStringCellValue();
        } else {
            name = row.getCell(1).getStringCellValue();
        }
        studentDTO.setName(name);
        teacherDTO.setName(name);

        String unit = row.getCell(3).getStringCellValue();
        studentDTO.setUnit(unit);
        teacherDTO.setUnit(unit);

        String num = String.valueOf(row.getCell(7).getNumericCellValue());
        studentDTO.setNumber(Integer.valueOf(num.substring(0, num.indexOf("."))));
        studentDTO.setUnitPrice(String.valueOf(row.getCell(8).getNumericCellValue()));
        studentDTO.setTotalPrice(String.valueOf(row.getCell(9).getNumericCellValue()));
        studentDTO.setNo(studentList.size() + 1);
        studentList.add(studentDTO);

        if (number > 13) {
            num = String.valueOf(row.getCell(10).getNumericCellValue());
            teacherDTO.setNumber(Integer.valueOf(num.substring(0, num.indexOf("."))));
            teacherDTO.setUnitPrice(String.valueOf(row.getCell(11).getNumericCellValue()));
            teacherDTO.setTotalPrice(String.valueOf(row.getCell(12).getNumericCellValue()));
            teacherDTO.setNo(teacherList.size() + 1);
            teacherList.add(teacherDTO);
        }
    }

    public void output() throws IOException {
        Workbook studentBook = new XSSFWorkbook();
        Set<Map.Entry<String, List<RowDTO>>> entrySet = studentMap.entrySet();
        System.out.println(entrySet.size());
        for (Map.Entry<String, List<RowDTO>> entry : entrySet) {
            String date = entry.getKey();
            Sheet sheet = studentBook.createSheet(date.substring(date.lastIndexOf("-") + 1));
            printRow(entry.getValue(), sheet);
        }
        FileOutputStream fileOutputStream = new FileOutputStream("temp.xlsx");
        studentBook.write(fileOutputStream);
        fileOutputStream.close();
    }

    public void printRow(List<RowDTO> list, Sheet sheet) {
        System.out.println(list.size());
        int rowNum = 1;
        for (RowDTO rowDTO : list) {
            Row row = sheet.createRow(rowNum++);

            row.createCell(0)
                    .setCellValue(rowDTO.getNo());

            row.createCell(1)
                    .setCellValue(rowDTO.getName());
            row.createCell(2)
                    .setCellValue(rowDTO.getUnit());

            row.createCell(3)
                    .setCellValue(rowDTO.getUnitPrice());

            row.createCell(4).setCellValue(rowDTO.getNumber());

            row.createCell(5)
                    .setCellValue(rowDTO.getTotalPrice());

            row.createCell(6)
                    .setCellValue(rowDTO.getComment());
        }
    }
}
