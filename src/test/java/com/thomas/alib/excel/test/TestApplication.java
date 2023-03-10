package com.thomas.alib.excel.test;

import com.thomas.alib.excel.exporter.ExcelExportSimple;
import com.thomas.alib.excel.importer.ExcelImportSimple;
import com.thomas.alib.excel.importer.SafetyResult;

import java.io.File;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;

public class TestApplication {
    public static void main(String[] args) {
        generateTemplate();
        exportTest();
        importTest();
        importTestSafety();
    }

    public static void generateTemplate() {
        ExcelExportSimple.with(new File("target/output/template.xlsx"))
                .createEmptySheet(YourDataBean.class)
                .export();
        System.out.println("生成模板示例执行完毕");
    }

    public static void exportTest() {
        LocalDate now = LocalDate.now();
        List<YourDataBean> exportList = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            YourDataBean item = new YourDataBean("王大头" + i, now.plusDays(i), i % 2,
                    "https://img2.woyaogexing.com/2023/03/06/1cfcb6418f13f881269aae6cff8bce16.jpg");
            exportList.add(item);
        }
        ExcelExportSimple.with(new File("target/output/exportTest.xlsx"))
                .addSheet(exportList)
                .export();
        System.out.println("导出示例执行完毕");
    }

    public static void importTest() {
        File file = new File("target/output/exportTest.xlsx");
        if (!file.exists()) throw new RuntimeException("导入文件不存在");
        List<YourDataBean> readList = ExcelImportSimple.readList(file, YourDataBean.class);
        System.out.println("导入示例执行完毕" + readList.size());
    }

    public static void importTestSafety() {
        File file = new File("target/output/exportTest.xlsx");
        if (!file.exists()) throw new RuntimeException("导入文件不存在");
        SafetyResult<List<SafetyResult<YourDataBean>>> readList = ExcelImportSimple.readListSafety(file, YourDataBean.class);
        System.out.println("安全模式导入示例执行完毕" + readList.getData().size());
    }
}
