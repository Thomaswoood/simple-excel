package com.thomas.alib.excel.exporter;


import com.thomas.alib.excel.exporter.provider.ExcelExporterByFile;
import com.thomas.alib.excel.exporter.provider.ExcelExporterByPath;
import com.thomas.alib.excel.exporter.provider.ExcelExporterByResponse;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.nio.file.Path;

/**
 * Excel导出工具类
 */
public class ExcelExportSimple {

    /**
     * 根据HttpServletResponse导出，直接导出到返回输出流中
     *
     * @param response 返回对象
     * @return 导出执行对象
     */
    public static ExcelExporterByResponse with(HttpServletResponse response) {
        return new ExcelExporterByResponse(response);
    }

    /**
     * 根据File对象导出，直接导出到对应文件中
     *
     * @param file 导出文件对象
     * @return 导出执行对象
     */
    public static ExcelExporterByFile with(File file) {
        return new ExcelExporterByFile(file);
    }

    /**
     * 根据Path对象导出，直接导出到对应Path路径中
     *
     * @param path 导出文件路径
     * @return 导出执行对象
     */
    public static ExcelExporterByPath with(Path path) {
        return new ExcelExporterByPath(path);
    }

    /**
     * 根据Path对象导出，直接导出到对应Path路径中
     *
     * @param path 导出文件路径
     * @return 导出执行对象
     */
    public static ExcelExporterByPath with(String path) {
        return new ExcelExporterByPath(path);
    }
}
