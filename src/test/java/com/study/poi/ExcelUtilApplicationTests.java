package com.study.poi;

import com.study.poi.util.ExcelUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.io.IOException;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.List;
import java.util.Map;

@RunWith(SpringRunner.class)
@SpringBootTest
public class ExcelUtilApplicationTests {

    @Test
    public void contextLoads() {
    }

    @Test
    public void testCreateFile(){
        List<Map<String,String>> dataList = null;
        DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyyMMddHHmmss");
        String currentDateTime = LocalDateTime.now().format(dtf);
        try {
            dataList = ExcelUtil.convertExcel("C:\\Users\\meng\\Desktop","按小时跑批.xls");
        } catch (IOException e) {
            e.printStackTrace();
        }
        String[] titlesForExcel = new String[]{"schemaName", "控制schema", "表名称", "目标表归属用户", "目标表空间", "存储模式", "同步方式", "优先级", "增量条件", "初始条件"};

        String[][] valuesForExcel = ExcelUtil.convertListToValuesForExcel(titlesForExcel.length,dataList);
        HSSFWorkbook workbook = ExcelUtil.getHSSFWorkbook("1",titlesForExcel,valuesForExcel);
        try {
            ExcelUtil.createExcelFile(workbook,"C:\\Users\\meng\\Desktop","按小时跑批"+currentDateTime+".xls");
        } catch (IOException e) {
            e.printStackTrace();
        }
        String[] titlesForCsv = new String[]{"schemaName",  "表名称"};
        String[][] valuesForCsv = ExcelUtil.convertListToValuesForCsv(titlesForCsv.length,dataList);
        workbook = ExcelUtil.getHSSFWorkbook("1",titlesForCsv,valuesForCsv);
        try {
            ExcelUtil.createExcelFile(workbook,"C:\\Users\\meng\\Desktop","按小时跑批"+currentDateTime+".csv");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
