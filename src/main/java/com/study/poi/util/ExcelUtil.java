package com.study.poi.util;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.*;
import java.util.List;
import java.util.Map;

public class ExcelUtil {

    public static List<Map<String,String>> convertExcel(String filePath, String fileName) throws IOException {
        File file = new File(filePath+"/"+fileName);
        ReadExcelUtil readExcelUtil = new ReadExcelUtil();
        return readExcelUtil.getExcelInfo(file);
    }
    public static String[][] convertListToValuesForExcel(int titleLen, List<Map<String, String>> dataList){
        String[][] arr = new String[dataList.size()][titleLen];
        for (int i=0; i<dataList.size(); i++) {
            arr[i][0] = dataList.get(i).get("schemaName");
            arr[i][1] = dataList.get(i).get("controlSchema");
            arr[i][2] = dataList.get(i).get("tableName");
            arr[i][3] = dataList.get(i).get("targetTableUser");
            arr[i][4] = dataList.get(i).get("targetTableSpace");
            arr[i][5] = dataList.get(i).get("storeModel");
            arr[i][6] = dataList.get(i).get("synModel");
            arr[i][7] = dataList.get(i).get("priorityLevel");
            arr[i][8] = dataList.get(i).get("addCondition");
            arr[i][9] = dataList.get(i).get("initCondition");
        }
        return arr;
    }

    public static String[][] convertListToValuesForCsv(int titleLen, List<Map<String, String>> dataList){
        String[][] arr = new String[dataList.size()][titleLen];
        for (int i=0; i<dataList.size(); i++) {
            arr[i][0] = dataList.get(i).get("schemaName");
            arr[i][1] = dataList.get(i).get("tableName");
        }
        return arr;
    }
    public static HSSFWorkbook getHSSFWorkbook(String sheetName,String[] title, String[][] values){
        //创建一个HSSFWorkbook对象 对应一个Excel文件
        HSSFWorkbook wb = new HSSFWorkbook();
        //在Workbook对象中添加一个sheet，对应excel中的而sheet
        HSSFSheet sheet = wb.createSheet(sheetName);

        HSSFRow headRow = sheet.createRow(0);
        HSSFCell headCell = headRow.createCell(0);
        headCell.setCellValue("百川数据同步");
        CellRangeAddress cra = new CellRangeAddress(0, 0, 0, title.length-1);
        sheet.addMergedRegion(cra);
        HSSFCellStyle headStyle = wb.createCellStyle();
        headStyle.setAlignment(HorizontalAlignment.CENTER);
        headStyle.setBorderTop(BorderStyle.THIN);
        headStyle.setBorderBottom(BorderStyle.THIN);
        headStyle.setBorderLeft(BorderStyle.THIN);
        headStyle.setBorderRight(BorderStyle.THIN);
        headCell.setCellStyle(headStyle);
        //设置表头
        HSSFRow tableHeader = sheet.createRow(1);
        //创建单元格，设置表头，设置表头居中
        HSSFCellStyle tableHeaderStyle = wb.createCellStyle();
        tableHeaderStyle.setAlignment(HorizontalAlignment.CENTER);
        tableHeaderStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        tableHeaderStyle.setWrapText(true);
        tableHeaderStyle.setBorderTop(BorderStyle.THIN);
        tableHeaderStyle.setBorderBottom(BorderStyle.THIN);
        tableHeaderStyle.setBorderLeft(BorderStyle.THIN);
        tableHeaderStyle.setBorderRight(BorderStyle.THIN);
        //声明列对象
        HSSFCell cell = null;
        //创建标题
        for (int i = 0; i<title.length; i++){
            cell = tableHeader.createCell(i);
            cell.setCellValue(title[i]);
            tableHeaderStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            tableHeaderStyle.setFillForegroundColor(IndexedColors.GREEN.index);
            cell.setCellStyle(tableHeaderStyle);
        }
        //创建内容
        HSSFRow row = null;
        for(int i = 0; i<values.length; i++){
            row = sheet.createRow(i+2);
            for(int j=0; j<values[i].length; j++){
                row.createCell(j).setCellValue(values[i][j]);
            }
        }
        setAutoColumnWidth(sheet,title.length);
        //返回
        return wb;
    }

    private static HSSFSheet setAutoColumnWidth(HSSFSheet sheet, int columnTotal){
        for (int i=0;i<columnTotal; i++){
            sheet.autoSizeColumn(i);
        }

//        for (int columnNum = 0; columnNum < columnTotal; columnNum++) {
//            int columnWidth = sheet.getColumnWidth(columnNum) / 256;
//            for (int rowNum = 0; rowNum < sheet.getLastRowNum(); rowNum++) {
//                HSSFRow currentRow;
//                //当前行未被使用过
//                if (sheet.getRow(rowNum) == null) {
//                    currentRow = sheet.createRow(rowNum);
//                } else {
//                    currentRow = sheet.getRow(rowNum);
//                }
//
//                if (currentRow.getCell(columnNum) != null) {
//                    HSSFCell currentCell = currentRow.getCell(columnNum);
//                    if (currentCell.getCellType() == HSSFCell.CELL_TYPE_STRING) {
//                        int length = currentCell.getStringCellValue().getBytes().length;
//                        if (columnWidth < length) {
//                            columnWidth = length;
//                        }
//                    }
//                }
//            }
//            sheet.setColumnWidth(columnNum, columnWidth * 256);
//        }
        return sheet;
    }

    public static void createExcelFile(HSSFWorkbook workbook, String targetPath, String fileName) throws IOException {
        FileOutputStream fileOutputStream = new FileOutputStream(targetPath+"/"+fileName);
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        workbook.close();
    }


}
