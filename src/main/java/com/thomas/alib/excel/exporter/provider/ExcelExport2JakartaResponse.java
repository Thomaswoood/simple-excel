package com.thomas.alib.excel.exporter.provider;


import com.thomas.alib.excel.exporter.ExcelExporterBase;
import com.thomas.alib.excel.utils.StringUtils;
import jakarta.servlet.http.HttpServletResponse;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.OutputStream;

/**
 * Excel导出者：针对http请求的response版本，主要负责处理response类和输出流的交互
 */
public class ExcelExport2JakartaResponse extends ExcelExporterBase<ExcelExport2JakartaResponse> {
    private static Logger logger = LoggerFactory.getLogger(ExcelExport2JakartaResponse.class);
    /**
     * 导出到response中时，在返回头中的文件名
     */
    String fileName;
    /**
     * 导出源response
     */
    private final HttpServletResponse response;

    /**
     * 构造方法
     *
     * @param response 本次导出请求对应的response
     */
    private ExcelExport2JakartaResponse(HttpServletResponse response) {
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
    public ExcelExport2JakartaResponse fileName(String fileName) {
        this.fileName = fileName;
        return this;
    }

    @Override
    protected OutputStream getOutputStream() throws Exception {
        if (StringUtils.isEmpty(fileName)) {
            fileName = System.currentTimeMillis() + ".xlsx";
        }
        response.reset();
        response.addHeader("Access-Control-Allow-Origin", "*");
        response.addHeader("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE, OPTIONS");
        response.addHeader("Access-Control-Expose-Headers", "Content-Disposition");
        response.addHeader("Content-disposition", "attachment; filename=" + fileName);
        response.addHeader("Access-Control-Allow-Credentials", "true");
        response.addHeader("Access-Control-Allow-Headers", "access-control-allow-origin, authority, content-type, version-info, X-Requested-With, x-access-token");
        response.addHeader("Access-Control-Max-Age", "3600");
        response.setContentType("application/octet-stream");
        return response.getOutputStream();
    }

    /**
     * 根据HttpServletResponse导出，直接导出到返回输出流中
     *
     * @param response 返回对象
     * @return 导出执行对象
     */
    public static ExcelExport2JakartaResponse with(HttpServletResponse response) {
        return new ExcelExport2JakartaResponse(response);
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
        return new ExcelExport2JakartaResponse((HttpServletResponse) response);
    }
}
