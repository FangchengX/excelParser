package service.exam;

import com.google.common.collect.Lists;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
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

    public List<String> exams = new ArrayList<>();
    public List<List<String>> classes = new ArrayList<>();

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
        for (int i = 0; i < exams.size(); i++) {
            String exam = exams.get(i);
            List<String> classesForExam = classes.get(i);
            classesForExam.forEach(clazz -> {
                Sheet sheet = examResult.createSheet(exam + "_" + clazz);
                printSheetTitle(sheet);
                if (!studentMap.containsKey(clazz)) {
                    System.out.println("Class not found:" + clazz);
                    return;
                }
                List<StudentDTO> students = studentMap.get(clazz);
                printRows(exam, clazz, students, sheet);
                sheet.setColumnWidth(1, 15 * 256);
                sheet.setColumnWidth(2, 15 * 256);
                sheet.setColumnWidth(3, 15 * 256);
                sheet.setColumnWidth(4, 10 * 256);
                sheet.setColumnWidth(5, 20 * 256);
                sheet.setColumnWidth(6, 10 * 256);
                sheet.setColumnWidth(7, 20 * 180);
            });
        }
//        for (Map.Entry<String, List<String>> entry : examMap.entrySet()) {
//            String key = entry.getKey();
//            if (!studentMap.containsKey(key)) {
//                System.out.println("Class not found:" + key);
//                continue;
//            }
//            List<StudentDTO> students = studentMap.get(key);
//            List<String> exams = entry.getValue();
//            exams.forEach(exam -> {
//                Sheet sheet = examResult.createSheet(exam + "_" + key);
//                printSheetTitle(sheet);
//                printRows(exam, key, students, sheet);
//                sheet.setColumnWidth(1, 15 * 256);
//                sheet.setColumnWidth(2, 15 * 256);
//                sheet.setColumnWidth(3, 10 * 256);
//                sheet.setColumnWidth(4, 15 * 256);
//                sheet.setColumnWidth(5, 10 * 256);
//                sheet.setColumnWidth(6, 20 * 180);
//            });
//        }
        FileOutputStream fileOutputStream = new FileOutputStream(outputPath);
        examResult.write(fileOutputStream);
        fileOutputStream.close();
    }

    private void printRows(String exam, String clazz, List<StudentDTO> students, Sheet sheet) {
        for (StudentDTO student : students) {
            Row row = sheet.createRow(student.getNumber());
            row.createCell(0).setCellValue(student.getNumber());
            row.createCell(1).setCellValue(student.getCollege());
            row.createCell(2).setCellValue(clazz);
            row.createCell(3).setCellValue(student.getCode());
            row.createCell(4).setCellValue(student.getName());
            row.createCell(5).setCellValue(exam);
            row.createCell(6);
            row.createCell(7);
            for (int i = 0; i < 8; i++) {
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

        Cell cell7 = row.createCell(6);

        Cell cell8 = row.createCell(7);

        List<Cell> cells = Lists.newArrayList(cell8, cell7, cell1, cell2, cell3, cell4, cell5, cell6);
        cells.forEach(aCell -> {
            aCell.setCellStyle(style);
        });
        cell1.setCellValue("序号");
        cell2.setCellValue("学院");
        cell3.setCellValue("班级");
        cell4.setCellValue("学号");
        cell5.setCellValue("姓名");
        cell6.setCellValue("科目");
        cell7.setCellValue("签名");
        cell8.setCellValue("备注（重修提前修请注明）");

    }

    private Map<String, List<String>> readExamInfo(String examSheet) throws Exception {
        Map<String, List<String>> map = new HashMap<>();
        FileInputStream fis = new FileInputStream(examSheet);
        try (Workbook workbook = WorkbookFactory.create(fis)) {
            Sheet sheet = workbook.getSheetAt(0);
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                Cell cell = row.getCell(4);
                String clazz;
                if (Objects.equals(cell.getCellType(), CellType.NUMERIC)) {
                    clazz = (int) (cell.getNumericCellValue()) + "";
                } else {
                    clazz = cell.getStringCellValue();
                }
                String exam = row.getCell(0).getStringCellValue();
                String last = exams.isEmpty() ? "" : exams.get(exams.size() - 1);
                List<String> classesForExam;
                if (!last.equals(exam)) {
                    exams.add(exam);
                    classesForExam = new ArrayList<>();
                    classes.add(classesForExam);
                } else {
                    classesForExam = classes.get(classes.size() - 1);
                }
                classesForExam.add(clazz);
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
                studentDTO.setCode(getCode(row));
                studentDTO.setClazz(row.getCell(2).getStringCellValue());
                studentDTO.setName(row.getCell(1).getStringCellValue());
                studentDTO.setCollege(row.getCell(3).getStringCellValue());
                List<StudentDTO> students = map.computeIfAbsent(studentDTO.getClazz(), unused -> Lists.newArrayList());
                studentDTO.setNumber(students.size() + 1);
                students.add(studentDTO);
            }
        }
        return map;
    }

    private String getCode(Row row) {
        Cell cell = row.getCell(0);
        if (Objects.equals(cell.getCellType(), CellType.STRING)) {
            return cell.getStringCellValue();
        } else {
            return String.valueOf((long) row.getCell(0).getNumericCellValue());
        }
    }
}
