package com.thomas.alib.excel.exporter;


import com.thomas.alib.excel.exporter.provider.*;
import com.thomas.alib.excel.interfaces.EFunction;

import java.io.File;
import java.io.OutputStream;
import java.nio.file.Path;

/**
 * Excel导出工具类
 */
public class ExcelExportSimple {

    /**
     * 根据HttpServletResponse导出，直接导出到返回输出流中
     * 此方法为尽量兼容历史版本的过渡方法，未来会废弃
     *
     * @param response 返回对象
     * @return 导出执行对象
     */
    @Deprecated
    public static ExcelExport2JavaxResponse withJavaxResponse(Object response) {
        return ExcelExport2JavaxResponse.withJavaxResponse(response);
    }

    /**
     * 根据HttpServletResponse导出，直接导出到返回输出流中
     * 此方法为尽量兼容历史版本的过渡方法，未来会废弃
     *
     * @param response 返回对象
     * @return 导出执行对象
     */
    @Deprecated
    public static ExcelExport2JakartaResponse withJakartaResponse(Object response) {
        return ExcelExport2JakartaResponse.withJakartaResponse(response);
    }

    /**
     * 根据File对象导出，直接导出到对应文件中
     *
     * @param file 导出文件对象
     * @return 导出执行对象
     */
    public static ExcelExport2File with(File file) {
        return ExcelExport2File.with(file);
    }

    /**
     * 根据Path对象导出，直接导出到对应Path路径中
     *
     * @param path 导出文件路径
     * @return 导出执行对象
     */
    public static ExcelExport2Path with(Path path) {
        return ExcelExport2Path.with(path);
    }

    /**
     * 根据Path对象导出，直接导出到对应Path路径中
     *
     * @param path 导出文件路径
     * @return 导出执行对象
     */
    public static ExcelExport2Path with(String path) {
        return ExcelExport2Path.with(path);
    }

    /**
     * 根据输出流提供者导出，导出到提供的输出流中
     *
     * @param source               输出流来源对象
     * @param outputStreamProvider 输出流提供者
     * @return 导出执行对象
     */
    public static <E> ExcelExport2OutputStream<E> with(E source, EFunction<E, OutputStream> outputStreamProvider) {
        return ExcelExport2OutputStream.with(source, outputStreamProvider);
    }
}
