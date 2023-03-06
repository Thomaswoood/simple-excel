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
        //没传入文件，默认在临时文件夹处理
        if (file == null) {
            file = new File(System.getProperty("java.io.tmpdir"), System.currentTimeMillis() + ".xlsx");
        }
        //检查指定文件目录是否存在，不存在则创建
        File parent = file.getParentFile();
        if (!parent.exists()) {
            parent.mkdirs();
        }
        //检查文件是否存在，存在先删除，再生成
        if (file.exists()) {
            file.delete();
        }
        file.createNewFile();
        return Files.newOutputStream(file.toPath());
    }
}
