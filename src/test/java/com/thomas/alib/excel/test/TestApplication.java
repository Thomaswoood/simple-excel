package com.thomas.alib.excel.test;

import com.thomas.alib.excel.exporter.ExcelExportSimple;

import java.io.File;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;

public class TestApplication {
    public static void main(String[] args) {
        exportTest();
    }

    public static void generateTemplate() {
        ExcelExportSimple.with(new File("target/output/template.xlsx"))
                .createEmptySheet(YourDataBean.class)
                .export();
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
    }
}
