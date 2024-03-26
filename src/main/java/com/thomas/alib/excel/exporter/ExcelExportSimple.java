package com.thomas.alib.excel.exporter;


import com.thomas.alib.excel.exporter.provider.ExcelExporterByFile;
import com.thomas.alib.excel.exporter.provider.ExcelExporterByJakartaResponse;
import com.thomas.alib.excel.exporter.provider.ExcelExporterByPath;
import com.thomas.alib.excel.exporter.provider.ExcelExporterByResponse;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.nio.file.Path;

/**
 * Excel导出工具类
 */
public class ExcelExportSimple {
    private static Logger logger = LoggerFactory.getLogger(ExcelExportSimple.class);

    /**
     * 根据HttpServletResponse导出，直接导出到返回输出流中
     *
     * @param response 返回对象
     * @return 导出执行对象
     */
    public static ExcelExporterByResponse with(javax.servlet.http.HttpServletResponse response) {
        logger.debug("基于response的导出开始");
        return new ExcelExporterByResponse(response);
    }

    /**
     * 根据HttpServletResponse导出，直接导出到返回输出流中
     *
     * @param response 返回对象
     * @return 导出执行对象
     */
    public static ExcelExporterByJakartaResponse with(jakarta.servlet.http.HttpServletResponse response) {
        logger.debug("基于response的导出开始");
        return new ExcelExporterByJakartaResponse(response);
    }

    /**
     * 根据File对象导出，直接导出到对应文件中
     *
     * @param file 导出文件对象
     * @return 导出执行对象
     */
    public static ExcelExporterByFile with(File file) {
        logger.debug("基于file的导出开始");
        return new ExcelExporterByFile(file);
    }

    /**
     * 根据Path对象导出，直接导出到对应Path路径中
     *
     * @param path 导出文件路径
     * @return 导出执行对象
     */
    public static ExcelExporterByPath with(Path path) {
        logger.debug("基于path的导出开始");
        return new ExcelExporterByPath(path);
    }

    /**
     * 根据Path对象导出，直接导出到对应Path路径中
     *
     * @param path 导出文件路径
     * @return 导出执行对象
     */
    public static ExcelExporterByPath with(String path) {
        logger.debug("基于路径的导出开始");
        return new ExcelExporterByPath(path);
    }
}
