package com.thomas.alib.excel.exporter.provider;


import com.thomas.alib.excel.exporter.ExcelExporterBase;
import com.thomas.alib.excel.utils.StringUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.servlet.http.HttpServletResponse;
import java.io.OutputStream;

/**
 * Excel导出者：针对http请求的response版本，主要负责处理response类和输出流的交互
 */
public class ExcelExport2JavaxResponse extends ExcelExporterBase<ExcelExport2JavaxResponse> {
    protected static final Logger logger = LoggerFactory.getLogger(ExcelExport2JavaxResponse.class);
    /**
     * 导出到response中时，在返回头中的文件名
     */
    protected String fileName;
    /**
     * 导出源response
     */
    protected HttpServletResponse response;

    /**
     * 构造方法
     *
     * @param response 本次导出请求对应的response
     */
    protected ExcelExport2JavaxResponse(HttpServletResponse response) {
        super();
        this.response = response;
        this.child = this;
        logger.debug("基于response的导出开始");
    }

    /**
     * 设置文件名
     *
     * @param fileName 文件名
     * @return 链式调用，返回自身
     */
    public ExcelExport2JavaxResponse fileName(String fileName) {
        this.fileName = fileName;
        return this;
    }

    /**
     * 提供具体导出的输出流
     *
     * @return 输出流对象
     * @throws Exception 获取response输出流时可能产生异常
     */
    @Override
    protected OutputStream getOutputStream() throws Exception {
        if (StringUtils.isEmpty(fileName)) {
            fileName = System.currentTimeMillis() + ".xlsx";
        }
        response.reset();
        response.addHeader("Content-disposition", "attachment; filename=" + fileName);
        response.setContentType("application/vnd.ms-excel;charset=utf-8");
        return response.getOutputStream();
    }

    /**
     * 销毁
     */
    @Override
    public void destroy() {
        super.destroy();
        fileName = null;
        response = null;
    }

    /**
     * 根据HttpServletResponse导出，直接导出到返回输出流中
     *
     * @param response 返回对象
     * @return 导出执行对象
     */
    public static ExcelExport2JavaxResponse with(HttpServletResponse response) {
        return new ExcelExport2JavaxResponse(response);
    }

    /**
     * 根据HttpServletResponse导出，直接导出到返回输出流中
     * 此方法为尽量兼容历史版本的过渡方法，未来会废弃
     *
     * @param response 返回对象
     * @return 导出执行对象
     */
    @Deprecated
    public static ExcelExport2JavaxResponse withJavaxResponse(Object response) {
        return new ExcelExport2JavaxResponse((HttpServletResponse) response);
    }
}
