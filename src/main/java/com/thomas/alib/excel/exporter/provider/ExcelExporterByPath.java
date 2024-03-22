package com.thomas.alib.excel.exporter.provider;

import com.thomas.alib.excel.exporter.ExcelExporterBase;
import com.thomas.alib.excel.utils.StringUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

/**
 * Excel导出者：针对基于Path直接导出为文件的版本，主要负责处理Path对象和文件输出流的交互
 */
public class ExcelExporterByPath extends ExcelExporterBase<ExcelExporterByPath> {
    private static Logger logger = LoggerFactory.getLogger(ExcelExporterByPath.class);
    private Path path;

    /**
     * 构造方法
     *
     * @param filePath 本次导出对应输出的文件路径字符串
     */
    public ExcelExporterByPath(String filePath) {
        this(StringUtils.isEmpty(filePath) ? null : Paths.get(filePath));
    }

    /**
     * 构造方法
     *
     * @param path 本次导出对应输出的文件路径
     */
    public ExcelExporterByPath(Path path) {
        super();
        this.path = path;
        if (this.path == null) {
            this.path = Paths.get(System.getProperty("java.io.tmpdir") + File.separator + System.currentTimeMillis() + ".xlsx");
        }
        this.child = this;
    }

    @Override
    protected OutputStream getOutputStream() throws Exception {
        logger.debug("文件将被导出到：" + path.toAbsolutePath());
        // 检查目录是否存在
        Path parent = path.getParent();
        if (parent != null && !Files.exists(parent)) {
            // 如果目录不存在，则创建目录
            Files.createDirectories(parent);
        } else {
            // 目录存在(或无法获取到目录)，检查指定文件是否存在，存在先删
            if (Files.exists(path)) {
                Files.delete(path);
            }
        }
        // 生成导出文件，如果文件所在目录不存在，会自动创建
        return Files.newOutputStream(path);
    }
}
