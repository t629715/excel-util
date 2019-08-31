package com.study.poi.util;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.util.StringUtils;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ReadExcelUtil {
    // 总行数
    private int totalRows = 0;
    // 总条数
    private int totalCells = 0;
    // 错误信息接收器
    private String errorMsg;

    // 构造方法
    public ReadExcelUtil() {
    }

    // 获取总行数
    public int getTotalRows() {
        return totalRows;
    }

    // 获取总列数
    public int getTotalCells() {
        return totalCells;
    }

    // 获取错误信息
    public String getErrorInfo() {
        return errorMsg;
    }

    /**
     * 读EXCEL文件，获取信息集合
     *
     * @param
     * @return
     */
    public List<Map<String, String>> getExcelInfo(File file) {
        String fileName = file.getName();// 获取文件名
        try {
            if (!validateExcel(fileName)) {// 验证文件名是否合格
                return null;
            }
            boolean isExcel2003 = true;// 根据文件名判断文件是2003版本还是2007版本
            if (isExcel2007(fileName)) {
                isExcel2003 = false;
            }
            return createExcel(new FileInputStream(file), isExcel2003);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    /**
     * 根据excel里面的内容读取信息
     *
     * @param is          输入流
     * @param isExcel2003 excel是2003还是2007版本
     * @return
     * @throws
     */
    public List<Map<String, String>> createExcel(InputStream is, boolean isExcel2003) {
        try {
            Workbook wb = null;
            if (isExcel2003) {// 当excel是2003时,创建excel2003
                wb = new HSSFWorkbook(is);
            } else {// 当excel是2007时,创建excel2007
                wb = new HSSFWorkbook(is);
            }
            return readExcelValue(wb);// 读取Excel里面的信息
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }

    /**
     * 读取Excel里面的信息
     *
     * @param wb
     * @return
     */
    private List<Map<String, String>> readExcelValue(Workbook wb) {
        // 得到第一个shell
        Sheet sheet = wb.getSheetAt(0);
        // 得到Excel的行数
        this.totalRows = sheet.getPhysicalNumberOfRows();
        // 得到Excel的列数(前提是有行数)
        if (totalRows > 1 && sheet.getRow(0) != null) {
            this.totalCells = sheet.getRow(0).getPhysicalNumberOfCells();
        }
        List<Map<String, String>> dataList = new ArrayList<Map<String, String>>();
        // 循环Excel行数
        List<String> tableNameList = new ArrayList<>();

        int total = 0;
        for (int r = 2; r < totalRows; r++) {
            Row row = sheet.getRow(r);
            if (row == null) {
                total ++;
                continue;
            }


            String nameTmp  = row.getCell(3).getStringCellValue().trim();
            if (StringUtils.isEmpty(nameTmp)){
                nameTmp = "";
            }else {
                String[] nameArr = nameTmp.split("\\.");
                if (nameArr.length > 1){
                    nameTmp = nameArr[1].toLowerCase();
                }else {
                    nameTmp = nameArr[0].toLowerCase();
                }
            }
            if (tableNameList.size() > 0 && tableNameList.contains(nameTmp)){
                total ++;
                continue;
            }
            // 循环Excel的列
            Map<String, String> map = new HashMap<String, String>();
            for (int c = 0; c < this.totalCells; c++) {
                Cell cell = row.getCell(c);
                if (null != cell) {
                    if (c == 1) {
                        String schemaName = cell.getStringCellValue().trim();
                        if ("p".equals(schemaName)){
                            cell = row.getCell(3);
                            String[] arr = cell.getStringCellValue().split("\\.");
                            if (arr.length > 1){
                                schemaName = arr[0].toLowerCase();
                            }else {
                                schemaName = "";
                            }

                        }else {
                            schemaName = cell.getStringCellValue().toLowerCase();
                        }
                        map.put("schemaName", schemaName);// schema名
                    } else if (c == 2) {
                        map.put("controlSchema", cell.getStringCellValue().toLowerCase());// 控制schema
                    } else if (c == 3) {
                        String tableName = cell.getStringCellValue().trim();
                        if (StringUtils.isEmpty(tableName)){
                            tableName = "";
                        }else {
                            String[] nameArr = tableName.split("\\.");
                            if (nameArr.length > 1){
                                tableName = nameArr[1].toLowerCase();

                            }else {
                                tableName = nameArr[0].toLowerCase();
                            }

                        }
                        tableNameList.add(tableName);
                        map.put("tableName", tableName);// 表名称
                    } else if (c == 4) {
                        String type = cell.getStringCellValue().trim();
                        if (type.contains("行")){
                            type = "0-行存";
                        }else {
                            type = "1-列存";
                        }
                        map.put("storeModel",type);
                    } else if (c == 5) {
                        String synModel = cell.getStringCellValue().trim();
                        if ("轮询".equals(synModel)){
                            synModel = "0-轮询";
                        }else if (synModel.contains("定时")){
                            synModel = "1-定时";
                        }
                        map.put("synModel",synModel);
                    }  else if (c == 7) {
                        String addCondition = cell.getStringCellValue().trim();
                        if (StringUtils.isEmpty(addCondition)){
                            addCondition = "";
                        }else {
                            addCondition = "where "+addCondition+"='@ETL_DATE@'";
                        }
                        map.put("addCondition", addCondition);//增量条件
                    } else if (c == 8) {
                        String initCondition = cell.getStringCellValue().trim();
                        if (StringUtils.isEmpty(initCondition)){
                            initCondition = "";
                        }else if ("全部数据".equals(initCondition)){
                            initCondition = "";
                        }else {
                            int indexOf = initCondition.indexOf("in");
                            if (indexOf > 0){
                                String column = initCondition.substring(0,indexOf).trim();
                                char[] arr = column.toCharArray();
                                int index = arr.length-1;
                                for (int i=arr.length-1; i>=0;i--){
                                    String str = arr[i]+"";
                                    if ((str).matches("[a-zA-Z]+")){
                                        break;
                                    }else{
                                        index --;
                                    }
                                }
                                column = column.substring(0,index);
                                String condition = initCondition.substring(indexOf);
                                initCondition = column+" "+condition;
                            }
                            if (initCondition.contains(">")){
                                DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyyMMdd");
                                String currentDate = LocalDateTime.now().format(dtf);
                                initCondition = initCondition.replace("where"," ");
                                String colName = initCondition.substring(0,initCondition.indexOf(">"));
                                colName = "  and " + colName +"<='" + currentDate + "'";
                                int startIndex = initCondition.indexOf("or");
                                if (initCondition.contains("or") || initCondition.contains("OR") ||initCondition.contains("Or") || initCondition.contains("oR")){

                                    if (startIndex < 0){
                                        startIndex = initCondition.indexOf("OR");
                                    }
                                    if (startIndex > 0){
                                        startIndex += 2;
                                    }
                                    colName = " (" + initCondition.substring(startIndex) + " and " + initCondition.substring(startIndex,initCondition.indexOf(">"))
                                    + "<='" + currentDate + "')";
                                }
                                if (startIndex > 0){
                                    initCondition = initCondition.substring(0,startIndex);
                                }
                                initCondition = "where "+initCondition + colName;
                            }else{
                                initCondition = initCondition.replace("where"," ");
                                initCondition = "where "+initCondition;
                            }
                        }
                        map.put("initCondition", initCondition);//初始化条件
                    }
                }
                map.put("targetTableUser", "");//目标表归属用户
                map.put("targetTableSpace", "");// 目标表空间
                map.put("priorityLevel","");//优先级
            }

            dataList.add(map);
        }
        System.out.println(total);
        return dataList;
    }

    /**
     * 验证EXCEL文件
     *
     * @param filePath
     * @return
     */
    public boolean validateExcel(String filePath) {
        if (filePath == null || !(isExcel2003(filePath) || isExcel2007(filePath))) {
            errorMsg = "文件名不是excel格式";
            return false;
        }
        return true;
    }
    // @描述：是否是2003的excel，返回true是2003
    public static boolean isExcel2003(String filePath) {
        return filePath.matches("^.+\\.(?i)(xls)$");
    }
    // @描述：是否是2007的excel，返回true是2007
    public static boolean isExcel2007(String filePath) {
        return filePath.matches("^.+\\.(?i)(xlsx)$");
    }

}
