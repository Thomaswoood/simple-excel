package com.thomas.alib.excel.exporter;


import javax.servlet.http.HttpServletResponse;
import java.io.File;

/**
 * Excel导出工具类
 */
public class ExcelExportSimple {

    /**
     * 构造方法
     *
     * @param response 返回对象
     * @return 初始化完毕的本类对象
     */
    public static ExcelExporterResponse with(HttpServletResponse response) {
        return new ExcelExporterResponse(response);
    }

    /**
     * 构造方法
     *
     * @param file 导出文件对象
     * @return 初始化完毕的本类对象
     */
    public static ExcelExporterFile with(File file) {
        return new ExcelExporterFile(file);
    }
}
