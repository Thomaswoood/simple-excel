package com.thomas.alib.excel.exporter.provider;



import com.thomas.alib.excel.exporter.ExcelExporterBase;
import com.thomas.alib.excel.utils.StringUtils;

import javax.servlet.http.HttpServletResponse;
import java.io.OutputStream;

/**
 * Excel导出者：针对http请求的response版本，主要负责处理response类和输出流的交互
 */
public class ExcelExporterByResponse extends ExcelExporterBase<ExcelExporterByResponse> {
    String fileName;
    private final HttpServletResponse response;

    /**
     * 构造方法
     *
     * @param response 本次导出请求对应的response
     */
    public ExcelExporterByResponse(HttpServletResponse response) {
        super();
        this.response = response;
        this.child = this;
    }

    /**
     * 设置文件名
     *
     * @param fileName 文件名
     * @return 链式调用，返回自身
     */
    public ExcelExporterByResponse fileName(String fileName) {
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
}
