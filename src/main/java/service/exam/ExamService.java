package service.exam;

import com.google.common.collect.Lists;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * @author kq644
 * @since 2020-11-02 12:46
 */
public class ExamService {

    private XSSFCellStyle style;
    private XSSFCellStyle infoStyle;

    public void parseData(String studentSheet, String examSheet, String outputPath) throws Exception {
        Map<String, List<StudentDTO>> studentMap = readStudentInfo(studentSheet);
        Map<String, List<String>> examMap = readExamInfo(examSheet);
        printResult(outputPath, studentMap, examMap);
    }

    private void printResult(String outputPath, Map<String, List<StudentDTO>> studentMap,
                             Map<String, List<String>> examMap) throws Exception {
        XSSFWorkbook examResult = new XSSFWorkbook();
        infoStyle = examResult.createCellStyle();
        infoStyle.setBorderBottom(BorderStyle.THIN);
        infoStyle.setBorderLeft(BorderStyle.THIN);
        infoStyle.setBorderRight(BorderStyle.THIN);
        infoStyle.setBorderTop(BorderStyle.THIN);
        style = examResult.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        XSSFFont font = examResult.createFont();
        font.setBold(true);
        font.setFontName("Adobe 黑体 Std R");
        style.setFont(font);
        style.setWrapText(true);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        for (Map.Entry<String, List<String>> entry : examMap.entrySet()) {
            String key = entry.getKey();
            if (!studentMap.containsKey(key)) {
                System.out.println("Class not found:" + key);
                continue;
            }
            List<StudentDTO> students = studentMap.get(key);
            List<String> exams = entry.getValue();
            exams.forEach(exam -> {
                Sheet sheet = examResult.createSheet(exam + "_" + key);
                printSheetTitle(sheet);
                printRows(exam, key, students, sheet);
                sheet.setColumnWidth(1, 15 * 256);
                sheet.setColumnWidth(2, 15 * 256);
                sheet.setColumnWidth(3, 10 * 256);
                sheet.setColumnWidth(4, 15 * 256);
                sheet.setColumnWidth(5, 10 * 256);
                sheet.setColumnWidth(6, 20 * 180);
            });
        }
        FileOutputStream fileOutputStream = new FileOutputStream(outputPath);
        examResult.write(fileOutputStream);
        fileOutputStream.close();
    }

    private void printRows(String exam, String clazz, List<StudentDTO> students, Sheet sheet) {
        for (StudentDTO student : students) {
            Row row = sheet.createRow(student.getNumber());
            row.createCell(0).setCellValue(student.getNumber());
            row.createCell(1).setCellValue(clazz);
            row.createCell(2).setCellValue(student.getCode());
            row.createCell(3).setCellValue(student.getName());
            row.createCell(4).setCellValue(exam);
            row.createCell(5);
            row.createCell(6);
            for (int i = 0; i < 7; i++) {
                row.getCell(i).setCellStyle(infoStyle);
            }
        }
    }

    private void printSheetTitle(Sheet sheet) {
        Row row = sheet.createRow(0);
        Cell cell1 = row.createCell(0);
        Cell cell2 = row.createCell(1);

        Cell cell3 = row.createCell(2);

        Cell cell4 = row.createCell(3);

        Cell cell5 = row.createCell(4);

        Cell cell6 = row.createCell(5);

        Cell cell = row.createCell(6);

        List<Cell> cells = Lists.newArrayList(cell, cell1, cell2, cell3, cell4, cell5, cell6);
        cells.forEach(aCell -> {
            aCell.setCellStyle(style);
        });
        cell1.setCellValue("序号");
        cell2.setCellValue("班级");
        cell3.setCellValue("学号");
        cell4.setCellValue("姓名");
        cell5.setCellValue("科目");
        cell6.setCellValue("签名");
        cell.setCellValue("备注（重修提前修请注明）");

    }

    private Map<String, List<String>> readExamInfo(String examSheet) throws Exception {
        Map<String, List<String>> map = new HashMap<>();
        FileInputStream fis = new FileInputStream(examSheet);
        try (Workbook workbook = WorkbookFactory.create(fis)) {
            Sheet sheet = workbook.getSheetAt(0);
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                String clazz = row.getCell(5).getStringCellValue();
                String exam = row.getCell(0).getStringCellValue();
                List<String> list = map.computeIfAbsent(clazz, unused -> Lists.newArrayList());
                list.add(exam);
            }
        }
        return map;
    }

    private Map<String, List<StudentDTO>> readStudentInfo(String studentSheetPath) throws Exception {
        Map<String, List<StudentDTO>> map = new HashMap<>();
        FileInputStream fis = new FileInputStream(studentSheetPath);
        try (Workbook workbook = WorkbookFactory.create(fis)) {
            Sheet sheet = workbook.getSheetAt(0);
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                StudentDTO studentDTO = new StudentDTO();
                studentDTO.setCode(row.getCell(0).getStringCellValue());
                studentDTO.setClazz(row.getCell(2).getStringCellValue());
                studentDTO.setName(row.getCell(1).getStringCellValue());
                List<StudentDTO> students = map.computeIfAbsent(studentDTO.getClazz(), unused -> Lists.newArrayList());
                studentDTO.setNumber(students.size() + 1);
                students.add(studentDTO);
            }
        }
        return map;
    }
}
