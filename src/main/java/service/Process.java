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
import java.util.regex.Pattern;

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
            if (!Pattern.matches("\\d+\\S+", sheet.getSheetName())) {
                return;
            }
            System.out.println("start process sheet:" + sheet.getSheetName());
            processSheet(sheet);
            System.out.println("complete process sheet:" + sheet.getSheetName());
        }
    }

    public void doProcessSheet(String filePath) throws Exception{
        FileInputStream fis = new FileInputStream(filePath);
        Workbook workbook = WorkbookFactory.create(fis);
        processSheet(workbook.getSheetAt(3));
    }

    public void processSheet(Sheet sheet) {
        for (int i = 0; i < sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
//            try {
            processRow(row, sheet);
//            } catch (Exception e) {
//                System.out.println(sheet.getSheetName()+","+row.getRowNum()+","+e.getMessage());
////                throw new RuntimeException();
////            }
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
    public void processRow(Row row, Sheet sheet) {
        //TODO 新建一个rowDTO， 根据日期放入对应的map中
        String date;
        boolean firstIsNull = Objects.isNull(row) || Objects.isNull(row.getCell(0));
        if (firstIsNull) {
            return;
        } else if (row.getCell(0).getCellType() == CellType.STRING && Objects.equals(row.getCell(0).getStringCellValue(), "")) {
            return;
        }
        if (row.getLastCellNum() != 0 && row.getCell(0).getCellType() == CellType.NUMERIC) {
            Date dateTime = row.getCell(0).getDateCellValue();
            if (dateTime.getMonth() != 10) {
                return;
            }
            date = format.format(dateTime);
        } else if (row.getLastCellNum() != 0 && row.getCell(0).getCellType() == CellType.STRING) {
            date = row.getCell(0).getStringCellValue();
        } else {
            return;
        }
        if (!Pattern.matches("\\d+-\\d+-\\d+", date)) {
            return;
        }
        RowDTO studentDTO = new RowDTO();
        RowDTO teacherDTO = new RowDTO();
        studentDTO.setSheetName(sheet.getSheetName());
        teacherDTO.setSheetName(sheet.getSheetName());
        List<RowDTO> studentList = studentMap.computeIfAbsent(date, (unused) -> Lists.newArrayList());
        List<RowDTO> teacherList = techerMap.computeIfAbsent(date, (unused) -> Lists.newArrayList());
        int number = row.getLastCellNum();
        String name;
        if (Objects.isNull(row.getCell(1)) || Objects.equals(row.getCell(1).getStringCellValue(), "")) {
            name = row.getCell(2).getStringCellValue();
        } else {
            name = row.getCell(1).getStringCellValue();
        }
        int studentOrderNum;
        int studentOrderUnitPrice;
        int studentOrderTotalPrice;
        int teacherOrderNum;
        int teacherOrderUnitPrice;
        int teacherOrderTotalPrice;
        int orderUnit;
        if (number == 13 || number == 16 || number == 14) {
            orderUnit = 3;
            studentOrderNum = 7;
            studentOrderUnitPrice = 8;
            studentOrderTotalPrice = 9;
            teacherOrderNum = 10;
            teacherOrderUnitPrice = 11;
            teacherOrderTotalPrice = 12;
        } else if (number == 15 || number == 12) {
            orderUnit = 2;
            studentOrderNum = 6;
            studentOrderUnitPrice = 7;
            studentOrderTotalPrice = 8;
            teacherOrderNum = 9;
            teacherOrderUnitPrice = 10;
            teacherOrderTotalPrice = 11;
        } else {
            System.out.println(number);
            return;
        }
        studentDTO.setName(name);
        teacherDTO.setName(name);

        String unit = getCellString(row, orderUnit);
        studentDTO.setUnit(unit);
        teacherDTO.setUnit(unit);

        addRowDTO(studentList, studentDTO, row, studentOrderNum, studentOrderUnitPrice, studentOrderTotalPrice);

        if (number > 13) {
            addRowDTO(teacherList, teacherDTO, row, teacherOrderNum, teacherOrderUnitPrice, teacherOrderTotalPrice);
        }
    }

    private void addRowDTO(List<RowDTO> rowDTOList, RowDTO rowDTO, Row row, Integer orderNum, Integer orderUnitPrice, Integer orderTotalPrice) {
        String num = getCellString(row, orderNum);
        if (Objects.equals(num, "")) {
            return;
        }
        rowDTO.setNumber(Integer.valueOf(num.substring(0, num.indexOf("."))));
        rowDTO.setUnitPrice(getCellDouble(row, orderUnitPrice));
        rowDTO.setTotalPrice(getCellDouble(row, orderTotalPrice));
        rowDTO.setNo(rowDTOList.size() + 1);
        rowDTOList.add(rowDTO);
    }

    private Double getCellDouble(Row row, Integer i) {
        Cell cell = row.getCell(i);
        if (Objects.isNull(cell)) {
            return 0.0;
        }
        switch (cell.getCellType()) {
            case NUMERIC:
                return cell.getNumericCellValue();
            case FORMULA:
                return cell.getNumericCellValue();
            default:
                return 0.0;
        }
    }

    private String getCellString(Row row, Integer i) {
        Cell cell = row.getCell(i);
        if (Objects.isNull(cell)) {
            return "0";
        }
        switch (cell.getCellType()) {
            case NUMERIC:
                return String.format("%.2f", cell.getNumericCellValue());
            case FORMULA:
                return String.format("%.2f", cell.getNumericCellValue());
            default:
                return cell.getStringCellValue();
        }
    }

    public void output() throws IOException {
        Workbook studentBook = new XSSFWorkbook();
        Set<Map.Entry<String, List<RowDTO>>> entrySet = studentMap.entrySet();
        for (Map.Entry<String, List<RowDTO>> entry : entrySet) {
            String date = entry.getKey();
            Sheet sheet = studentBook.createSheet(date.substring(date.lastIndexOf("-") + 1));
            printSheetTitle(sheet);
            printRow(entry.getValue(), sheet);
        }
        FileOutputStream fileOutputStream = new FileOutputStream("temp.xlsx");
        studentBook.write(fileOutputStream);
        fileOutputStream.close();
    }

    public void printSheetTitle(Sheet sheet) {
        Row row = sheet.createRow(0);
        row.createCell(0)
                .setCellValue("编号");

        row.createCell(1)
                .setCellValue("名");
        row.createCell(2)
                .setCellValue("单位");

        row.createCell(3)
                .setCellValue("单价");

        row.createCell(4).setCellValue("数量");

        row.createCell(5)
                .setCellValue("总价");

        row.createCell(6)
                .setCellValue("备注");
    }

    public void printRow(List<RowDTO> list, Sheet sheet) {
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
            row.createCell(7)
                    .setCellValue(rowDTO.getSheetName());
        }
    }
}
