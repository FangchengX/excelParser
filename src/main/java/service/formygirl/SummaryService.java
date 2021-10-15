package service.formygirl;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import service.formygirl.dto.ResultDTO;

/**
 * @author kq644
 * @since 2021-10-14 12:42
 */
public class SummaryService {

    public void doExamSummary(String examName, String appResultPath, String appProgressPath) throws Exception {
        Map<String, ResultDTO> gradeMap = readAppResult(appResultPath);
        readAppProgress(appProgressPath, gradeMap);
        writeResult(examName, gradeMap);
    }

    private void writeResult(String examName, Map<String, ResultDTO> gradeMap) throws IOException {
        Path resultPath = Paths.get("result");
        if (!resultPath.toFile().exists()) {
            resultPath.toFile().mkdir();
        }
        Path resultFilePath = resultPath.resolve(examName + ".txt");
        try (BufferedWriter bufferedWriter = new BufferedWriter(new FileWriter(resultFilePath.toFile()))) {
            for (ResultDTO resultDTO : gradeMap.values()) {
                bufferedWriter.write(resultDTO.toString());
                bufferedWriter.write("\n");
            }
        }
    }

    private void readAppProgress(String appProgressPath, Map<String, ResultDTO> gradeMap) throws IOException {
        Set<String> nameSet = new HashSet<>();
        try (FileInputStream fis = new FileInputStream(new File(appProgressPath));
             Workbook workbook = WorkbookFactory.create(fis)) {
            Sheet sheet = workbook.getSheetAt(0);
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                String name = row.getCell(1).getStringCellValue();
                String progress = row.getCell(7).getStringCellValue();
                String code = row.getCell(0).getStringCellValue();
                String phone = row.getCell(5).getStringCellValue();
                ResultDTO resultDTO = gradeMap.computeIfAbsent(code, unused -> new ResultDTO());
                resultDTO.setName(name);
                resultDTO.setProgress(progress);
                resultDTO.setCode(code);
                resultDTO.setPhone(phone);
                if (!nameSet.add(name)) {
                    System.out.println("warning: 发现重复的名字：" + name);
                }
            }
        } catch (Exception e) {
            System.out.println("read app progess file failed");
            throw e;
        }
    }

    private Map<String, ResultDTO> readAppResult(String appResultPath) throws IOException {
        Map<String, ResultDTO> map = new HashMap<>();
        Set<String> nameSet = new HashSet<>();
        try (FileInputStream fis = new FileInputStream(appResultPath);
             Workbook workbook = WorkbookFactory.create(fis)) {
            Sheet sheet = workbook.getSheetAt(0);
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                String name = row.getCell(1).getStringCellValue();
                String grade = row.getCell(5).getStringCellValue();
                String code = row.getCell(2).getStringCellValue();
                if (!nameSet.add(name)) {
                    System.out.println("warning: 发现重复的名字：" + name);
                }
                ResultDTO resultDTO = new ResultDTO();
                resultDTO.setName(name);
                resultDTO.setCode(code);
                resultDTO.setGrade(parseGrade(grade));
                map.put(code, resultDTO);
            }
        } catch (Exception e) {
            System.out.println("read app result failed");
            throw e;
        }
        return map;
    }

    private int parseGrade(String grade) {
        try {
            return Integer.parseInt(grade);
        } catch (Exception e) {
            System.out.println("parse grade failed, input: grade");
            return 0;
        }
    }
}
