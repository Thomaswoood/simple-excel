package com.thomas.alib.excel.exporter;


import java.io.File;
import java.io.OutputStream;
import java.nio.file.Files;

/**
 * Excel导出者：针对直接导出为File文件的版本，主要负责处理File类和输出流的交互
 */
public class ExcelExporterFile extends ExcelExporterBase<ExcelExporterFile> {
    private File file;

    /**
     * 构造方法
     *
     * @param file 本次导出对应输出的文件
     */
    ExcelExporterFile(File file) {
        super();
        this.file = file;
        this.child = this;
    }

    @Override
    OutputStream getOutputStream() throws Exception {
        if (file == null) {
            file = new File(System.getProperty("java.io.tmpdir"), System.currentTimeMillis() + ".xlsx");
        }
        return Files.newOutputStream(file.toPath());
    }
}
