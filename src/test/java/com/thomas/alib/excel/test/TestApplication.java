package com.thomas.alib.excel.test;

import com.thomas.alib.excel.exporter.provider.ExcelExport2File;
import com.thomas.alib.excel.exporter.provider.ExcelExport2Path;
import com.thomas.alib.excel.importer.ExcelImportSimple;
import com.thomas.alib.excel.importer.SafetyResult;
import org.apache.logging.log4j.Level;
import org.apache.logging.log4j.core.config.Configurator;

import java.io.File;
import java.io.FileInputStream;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;

/**
 * 示例程序
 */
public class TestApplication {
    public static void main(String[] args) {
        // 设置日志级别为DEBUG
        Configurator.setRootLevel(Level.DEBUG);
        // 将日志输出到控制台
        System.setProperty("org.slf4j.simpleLogger.defaultLogLevel", "debug");
        System.setProperty("org.slf4j.simpleLogger.showDateTime", "true");
        System.setProperty("org.slf4j.simpleLogger.dateTimeFormat", "yyyy-MM-dd HH:mm:ss.SSS");
        System.setProperty("org.slf4j.simpleLogger.showThreadName", "false");
        System.setProperty("org.slf4j.simpleLogger.showLogName", "false");
        System.setProperty("org.slf4j.simpleLogger.showShortLogName", "true");
        //示例
        generateTemplate();
        exportTest();
        exportTest2();
        importTest();
        importTest1_0_2();
        importTestSafety();
        importTestSafety1_0_2();
    }

    /**
     * 使用导出工具，根据实体类生成一个模板实例
     */
    public static void generateTemplate() {
        ExcelExport2File.with(new File("target/output/template.xlsx"))
                .createEmptySheet(YourDataBean.class)
                .export();
        System.out.println("生成模板示例执行完毕");
    }

    /**
     * 创建测试数据
     *
     * @return 测试数据列表
     */
    public static List<YourDataBean> makeTestData() {
        LocalDate now = LocalDate.now();
        List<YourDataBean> exportList = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            YourDataBean item = new YourDataBean("王大头" + i, now.plusDays(i), i % 2,
                    "https://img2.woyaogexing.com/2023/03/06/1cfcb6418f13f881269aae6cff8bce16.jpg");
            exportList.add(item);
        }
        return exportList;
    }

    /**
     * 导出示例
     */
    public static void exportTest() {
        ExcelExport2File.with(new File("target/output/exportTest.xlsx"))
                .addSheet(makeTestData())
                .export();
        System.out.println("导出示例执行完毕");
    }

    /**
     * 导出示例2
     */
    public static void exportTest2() {
        ExcelExport2Path.with("target/output/exportTest.xlsx")
                .addSheet(makeTestData())
                .export();
        System.out.println("导出示例执行完毕");
    }

    /**
     * 普通模式导入读取示例
     */
    public static void importTest() {
        File file = new File("target/output/exportTest.xlsx");
        if (!file.exists()) throw new RuntimeException("导入文件不存在");
        List<YourDataBean> readList = ExcelImportSimple.readList(file, YourDataBean.class);
        System.out.println("导入示例执行完毕" + readList.size());
    }

    /**
     * 普通模式导入，1.0.2新功能读取示例
     */
    public static void importTest1_0_2() {
        File file = new File("target/output/exportTest.xlsx");
        if (!file.exists()) throw new RuntimeException("导入文件不存在");
        List<YourDataBean> readList = ExcelImportSimple.readList(FileInputStream::new, file, YourDataBean.class);
        System.out.println("导入示例执行完毕" + readList.size());
    }

    /**
     * 安全模式导入读取示例
     */
    public static void importTestSafety() {
        File file = new File("target/output/exportTest.xlsx");
        if (!file.exists()) throw new RuntimeException("导入文件不存在");
        SafetyResult<List<SafetyResult<YourDataBean>>> readList = ExcelImportSimple.readListSafety(file, YourDataBean.class);
        System.out.println("安全模式导入示例执行完毕" + readList.getData().size());
    }

    /**
     * 安全模式导入，1.0.2新功能读取示例
     */
    public static void importTestSafety1_0_2() {
        File file = new File("target/output/exportTest.xlsx");
        if (!file.exists()) throw new RuntimeException("导入文件不存在");
        SafetyResult<List<SafetyResult<YourDataBean>>> readList = ExcelImportSimple.readListSafety(FileInputStream::new, file, YourDataBean.class);
        System.out.println("安全模式导入示例执行完毕" + readList.getData().size());
    }
}
