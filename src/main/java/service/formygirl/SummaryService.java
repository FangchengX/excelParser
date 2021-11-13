package service.formygirl;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.google.common.collect.Lists;
import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import service.formygirl.dto.DepartDTO;
import service.formygirl.dto.MemberDTO;
import service.formygirl.dto.ResultDTO;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.Collator;
import java.util.*;
import java.util.stream.Collectors;

/**
 * @author kq644
 * @since 2021-10-14 12:42
 */
public class SummaryService {

    public static final Map<String, MemberDTO> ID_ACCOUNT_MAP;
    public static final Map<String, MemberDTO> NAME_ACCOUNT_MAP;
    public static final Set<String> REPEAT_NAMES;
    public static final Map<String, MemberDTO> ACCOUNT_MAP;

    //
    static {
        Map<String, MemberDTO> tempMap = new HashMap<>();
        Map<String, MemberDTO> nameMap = new HashMap<>();
        Map<String, MemberDTO> accountMap = new HashMap<>();

        Set<String> repeatNames = new HashSet<>();
        try {
            String memberInfo = FileUtils.readFileToString(new File("member.json"), StandardCharsets.UTF_8);
            List<MemberDTO> members = JSON.parseArray(memberInfo, MemberDTO.class);
            for (MemberDTO memberDTO : members) {
                if (Objects.isNull(memberDTO.getId()) || Objects.isNull(memberDTO.getAccount())) {
                    System.out.println("Missing id, detail:" + JSON.toJSONString(memberDTO));
                } else {
                    tempMap.put(memberDTO.getId(), memberDTO);
                }
                if (nameMap.containsKey(memberDTO.getName())) {
                    if (!nameMap.get(memberDTO.getName()).getId().equals(memberDTO.getId())) {
//                        System.out.println("Warning, 医院系统存在重名。 姓名：" + memberDTO.getName());
                        repeatNames.add(memberDTO.getName());
                    }
                } else {
                    nameMap.put(memberDTO.getName(), memberDTO);
                }
                if (Objects.nonNull(memberDTO.getAccount())) {
                    accountMap.put(memberDTO.getAccount(), memberDTO);
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
            tempMap = new HashMap<>();
        }
        ID_ACCOUNT_MAP = tempMap;
        NAME_ACCOUNT_MAP = nameMap;
        REPEAT_NAMES = repeatNames;
        ACCOUNT_MAP = accountMap;
    }

    public void parseMembers() throws IOException {
        String memberInfoPath = "C:\\Users\\kq644\\Desktop\\lyy\\users.json";
        String memberInfo = FileUtils.readFileToString(new File(memberInfoPath), StandardCharsets.UTF_8);
        JSONArray array = JSON.parseArray(memberInfo);
        List<MemberDTO> members = array.stream().map(o -> {
            JSONObject object = (JSONObject) o;
            String code = object.getString("ZSDM");
            String id = object.getString("ID");
            String account = object.getString("YHZH");
            if (Objects.isNull(account)) {
                account = object.getString("PY");
            }
            String depart = object.getString("BMMC");
            String name = object.getString("XM");
            return new MemberDTO(account, code, id, depart, name);
        }).collect(Collectors.toList());
        FileUtils.writeStringToFile(new File("member.json"), JSON.toJSONString(members), StandardCharsets.UTF_8);
    }

    public void doDingdingSummary(String examName, String dingdingResultPath, String outputFolder) throws IOException
        , InvocationTargetException, IllegalAccessException {
        Map<String, ResultDTO> appResult = readAppResult(examName);
        List<ResultDTO> results = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(new File(dingdingResultPath));
             Workbook workbook = WorkbookFactory.create(fis)) {
            Sheet sheet = workbook.getSheetAt(0);
            for (int i = 2; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                String id = row.getCell(2).getStringCellValue();
                String name = row.getCell(0).getStringCellValue();
                String grade = readGrade(row.getCell(13));
                String reGrade = readGrade(row.getCell(14));
                String progress = row.getCell(12).getStringCellValue();
                String depart = row.getCell(5).getStringCellValue();
                String subDepart = row.getCell(6).getStringCellValue();
                String type = row.getCell(8).getStringCellValue();
                ResultDTO resultDTO = getResultData(id, appResult, name);
                resultDTO.setGrade(Math.max(parseGrade(grade), parseGrade(reGrade)));
                resultDTO.setName(name);
                resultDTO.setProgress(progress);
                resultDTO.setDepart(depart);
                resultDTO.setSubDepart(subDepart);
                resultDTO.setType(type);
                MemberDTO memberDTO = ID_ACCOUNT_MAP.getOrDefault(id, null);
                if (Objects.nonNull(memberDTO)) {
                    resultDTO.setCode(memberDTO.getCode());
                    resultDTO.setAccount(memberDTO.getAccount());
                }
                resultDTO.setId(id);
                results.add(resultDTO);
            }
//            Collection<ResultDTO> resultDTOS = appResult.values();
//            parseAppResult(resultDTOS);
//            results.addAll(resultDTOS);
        } catch (Exception e) {
            System.out.println("Read ding ding result failed");
            throw e;
        }
        writeDingDingResult(examName, results, outputFolder);
    }

    private String readGrade(Cell cell) {
        if (Objects.isNull(cell)) {
            return "0";
        }
        String grade;
        try {
            grade = cell.getStringCellValue();
        } catch (Exception e) {
            grade = String.valueOf(cell.getNumericCellValue());
        }
        return grade;
    }

    private Map<String, ResultDTO> readAppResult(String examName) throws IOException {
        Path resultFilePath = getAppResultPath(examName);
        if (!resultFilePath.toFile().exists()) {
            System.out.println("App数据不存在，路径:" + resultFilePath);
            return new HashMap<>();
        }
        List<ResultDTO> results = JSON.parseArray(FileUtils.readFileToString(resultFilePath.toFile(), StandardCharsets.UTF_8), ResultDTO.class);
        return results.stream().collect(Collectors.toMap(ResultDTO::getAccount, resultDTO -> resultDTO));
    }

    private ResultDTO getResultData(String id, Map<String, ResultDTO> appDataMap, String name) throws InvocationTargetException, IllegalAccessException {
        if (!ID_ACCOUNT_MAP.containsKey(id) && (REPEAT_NAMES.contains(name) || !NAME_ACCOUNT_MAP.containsKey(name))) {
            System.out.println("医院系统中未找到该用户：" + name + "， 用户ID为：" + id + "。请核查");
            return new ResultDTO();
        }
        MemberDTO memberDTO = ID_ACCOUNT_MAP.getOrDefault(id, NAME_ACCOUNT_MAP.get(name));
        String account = memberDTO.getAccount();
        ResultDTO resultDTO = new ResultDTO();
        if (appDataMap.containsKey(account)) {
            ResultDTO appResult = appDataMap.get(account);
            resultDTO.setName(name);
            resultDTO.setProgress(appResult.getProgress());
            resultDTO.setCode(appResult.getCode());
            resultDTO.setPhone(appResult.getPhone());
            resultDTO.setAccount(appResult.getAccount());
            resultDTO.setGrade(appResult.getGrade());
        }
        resultDTO.setDepart(memberDTO.getDepart());
        return resultDTO;
    }

    private void parseAppResult(Collection<ResultDTO> resultDTOS) {
        for (ResultDTO result : resultDTOS) {
            if (!ACCOUNT_MAP.containsKey(result.getAccount()) && (REPEAT_NAMES.contains(result.getName()) || !NAME_ACCOUNT_MAP.containsKey(result.getName()))) {
                continue;
            }
            MemberDTO memberDTO = ACCOUNT_MAP.getOrDefault(result.getAccount(), NAME_ACCOUNT_MAP.get(result.getName()));
            result.setDepart(memberDTO.getDepart());
            result.setId(memberDTO.getId());
        }
    }

    public void doExamSummary(String examName, String appResultPath, String appProgressPath) throws Exception {
        Map<String, ResultDTO> gradeMap = new HashMap<>();
        if (Objects.nonNull(appResultPath)) {
            readAppResult(appResultPath, gradeMap);
        }
        if (Objects.nonNull(appProgressPath)) {
            readAppProgress(appProgressPath, gradeMap);
        }
        writeResult(examName, gradeMap.values());
    }

    private void writeResult(String examName, Collection<ResultDTO> results) throws IOException {
        Path resultFilePath = getAppResultPath(examName);
        FileUtils.writeStringToFile(resultFilePath.toFile(), JSON.toJSONString(results), StandardCharsets.UTF_8);
    }

    private Path getAppResultPath(String examName) {
        Path resultPath = Paths.get("result");
        if (!resultPath.toFile().exists()) {
            resultPath.toFile().mkdir();
        }
        return resultPath.resolve(examName + ".json");
    }

    private void writeDingDingResult(String examName, List<ResultDTO> results, String folderName) throws IOException {
        Path resultPath = Paths.get(folderName);
        Path resultFilePath = resultPath.resolve(examName + ".xlsx");
        XSSFWorkbook examResult = new XSSFWorkbook();
        Sheet sheet = examResult.createSheet("考核结果");
        Map<String, DepartDTO> summaryMap = printExamSheet(results, sheet);
        Sheet summary = examResult.createSheet("统计结果");
        printSummarySheet(summary, summaryMap);
        FileOutputStream fileOutputStream = new FileOutputStream(resultFilePath.toFile());
        examResult.write(fileOutputStream);
        fileOutputStream.close();
        System.out.println("成绩统计已完成，请查看：" + resultFilePath);
    }

    private void printSummarySheet(Sheet summary, Map<String, DepartDTO> summaryMap) {
        printSummaryTitle(summary);
        List<DepartDTO> departs = Lists.newArrayList(summaryMap.values());
        departs.sort(((o1, o2) -> o2.getFailNumber() - o1.getFailNumber()));
        for (int i = 1; i <= departs.size(); i++) {
            Row row = summary.createRow(i);
            DepartDTO result = departs.get(i - 1);
            row.createCell(0).setCellValue(result.getDepart());
            row.createCell(1).setCellValue(result.getTotalNumber());
            row.createCell(2).setCellValue(result.getPassNumber());
            row.createCell(3).setCellValue(result.getFailNumber());
            row.createCell(4).setCellValue(String.format("%.2f", result.getPassPercent() * 100) + "%");
        }
    }

    private void printSummaryTitle(Sheet sheet) {
        Row row = sheet.createRow(0);
        Cell cell1 = row.createCell(0);
        Cell cell2 = row.createCell(1);
        Cell cell3 = row.createCell(2);
        Cell cell4 = row.createCell(3);
        Cell cell5 = row.createCell(4);
        cell1.setCellValue("科室");
        cell2.setCellValue("总考核人数");
        cell3.setCellValue("通过人数");
        cell4.setCellValue("不通过人数");
        cell5.setCellValue("通过率");
    }

    private Map<String, DepartDTO> printExamSheet(List<ResultDTO> results, Sheet sheet) {
        Map<String, DepartDTO> map = new HashMap<>();
        printResultTitle(sheet);
        Comparator comparator = Collator.getInstance(Locale.CHINA);
        results.sort((o1, o2) -> {
            if (Objects.equals(o1.getDepart(), o2.getDepart())) {
                return comparator.compare(o1.getSubDepart(), o2.getSubDepart());
            } else {
                return comparator.compare(o1.getDepart(), o2.getDepart());
            }
        });
        for (int i = 1; i <= results.size(); i++) {
            Row row = sheet.createRow(i);
            ResultDTO result = results.get(i - 1);
            row.createCell(0).setCellValue(result.getName());
            row.createCell(1).setCellValue(result.getAccount());
            row.createCell(2).setCellValue(result.getCode());
            row.createCell(3).setCellValue(result.getDepart());
            row.createCell(4).setCellValue(result.getSubDepart());
            row.createCell(5).setCellValue(result.getId());
            row.createCell(6).setCellValue(result.getType());
            row.createCell(7).setCellValue(result.getProgress());
            row.createCell(8).setCellValue(result.getGrade());
            row.createCell(9).setCellValue(result.findCheckMessage());
            String depart = result.getDepart();
            if (Objects.equals(result.getType().trim(), "必修")) {
                DepartDTO departDTO = map.computeIfAbsent(depart, unused -> new DepartDTO(depart));
                departDTO.addNumber(result.pass());
            }
        }
        return map;
    }

    public void printResultTitle(Sheet sheet) {
        Row row = sheet.createRow(0);
        Cell cell1 = row.createCell(0);
        Cell cell2 = row.createCell(1);
        Cell cell3 = row.createCell(2);
        Cell cell4 = row.createCell(3);
        Cell cell5 = row.createCell(4);
        Cell cell6 = row.createCell(5);
        Cell cell7 = row.createCell(6);
        Cell cell8 = row.createCell(7);
        Cell cell9 = row.createCell(8);
        Cell cell10 = row.createCell(9);

        cell1.setCellValue("姓名");
        cell2.setCellValue("账号");
        cell3.setCellValue("终生代码");
        cell4.setCellValue("科室");
        cell5.setCellValue("部门");
        cell6.setCellValue("工号");
        cell7.setCellValue("类型");
        cell8.setCellValue("课程进度");
        cell9.setCellValue("考核结果");
        cell10.setCellValue("是否通过");
    }

    private void readAppProgress(String appProgressPath, Map<String, ResultDTO> gradeMap) throws IOException {
        System.out.println("FilePath:" + appProgressPath);
        try (FileInputStream fis = new FileInputStream(new File(appProgressPath));
             Workbook workbook = WorkbookFactory.create(fis)) {
            Sheet sheet = workbook.getSheetAt(0);
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                String name = row.getCell(1).getStringCellValue();
                String progress = row.getCell(7).getStringCellValue();
                String account = row.getCell(0).getStringCellValue();
                String code = row.getCell(2).getStringCellValue();
                String phone = row.getCell(5).getStringCellValue();
                ResultDTO resultDTO = gradeMap.computeIfAbsent(account, unused -> new ResultDTO());
                resultDTO.setName(name);
                resultDTO.setProgress(progress);
                resultDTO.setCode(code);
                resultDTO.setPhone(phone);
                resultDTO.setAccount(account);

            }
        } catch (Exception e) {
            System.out.println("read app progess file failed");
            throw e;
        }
    }

    private Map<String, ResultDTO> readAppResult(String appResultPath, Map<String, ResultDTO> map) throws IOException {
        System.out.println("FilePath:" + appResultPath);
        try (FileInputStream fis = new FileInputStream(appResultPath);
             Workbook workbook = WorkbookFactory.create(fis)) {
            Sheet sheet = workbook.getSheetAt(0);
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                String name = row.getCell(1).getStringCellValue();
                String grade = row.getCell(5).getStringCellValue();
                String account = row.getCell(2).getStringCellValue();
                String code = row.getCell(4).getStringCellValue();

                ResultDTO resultDTO = new ResultDTO();
                resultDTO.setName(name);
                resultDTO.setAccount(account);
                resultDTO.setGrade(parseGrade(grade));
                resultDTO.setCode(code);
                map.put(account, resultDTO);
            }
        } catch (Exception e) {
            System.out.println("read app result failed");
            throw e;
        }
        return map;
    }

    private double parseGrade(String grade) {
        if (Objects.equals(grade, "未开始") || Objects.equals(grade, "未参与")) {
            return 0.0;
        }
        grade = dealGrade(grade);
        try {
            return Double.parseDouble(grade);
        } catch (Exception e) {
            System.out.println("parse grade failed, input:" + grade);
            return 0.0;
        }
    }

    private String dealGrade(String grade) {
        int start = Math.max(grade.indexOf("("), grade.indexOf("（"));
        int end = Math.max(grade.indexOf(")"), grade.indexOf("）"));
        if (start < 0 || end < 0) {
            return grade;
        }
        return grade.substring(start + 1, end).trim();
    }
}
