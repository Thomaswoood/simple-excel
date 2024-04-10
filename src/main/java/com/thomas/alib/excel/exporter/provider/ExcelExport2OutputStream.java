package com.thomas.alib.excel.exporter.provider;


import com.thomas.alib.excel.exporter.ExcelExporterBase;
import com.thomas.alib.excel.interfaces.EFunction;

import java.io.OutputStream;

/**
 * Excel导出者：针对直接提供输出流的版本
 *
 * @param <T> 输出流来源对象泛型
 */
public class ExcelExport2OutputStream<T> extends ExcelExporterBase<ExcelExport2OutputStream<T>> {
    /**
     * 输出流来源对象
     */
    protected T source;
    /**
     * 输出流提供者
     */
    protected EFunction<T, OutputStream> outputStreamProvider;

    /**
     * 构造方法
     *
     * @param source               输出流来源对象
     * @param outputStreamProvider 输出流提供者
     */
    protected ExcelExport2OutputStream(T source, EFunction<T, OutputStream> outputStreamProvider) {
        super();
        this.source = source;
        this.outputStreamProvider = outputStreamProvider;
        this.child = this;
    }

    /**
     * 提供具体导出的输出流
     *
     * @return 输出流对象
     * @throws Exception 从导出源获取输出流时可能产生异常
     */
    @Override
    protected OutputStream getOutputStream() throws Exception {
        if (source == null || outputStreamProvider == null)
            throw new NullPointerException("需要提供导出源");
        return outputStreamProvider.apply(source);
    }

    /**
     * 销毁
     */
    @Override
    public void destroy() {
        super.destroy();
        source = null;
        outputStreamProvider = null;
    }

    /**
     * 根据导出源创建导出执行对象
     *
     * @param source               输出流来源对象
     * @param outputStreamProvider 输出流提供者
     * @return 导出执行对象
     */
    public static <E> ExcelExport2OutputStream<E> with(E source, EFunction<E, OutputStream> outputStreamProvider) {
        return new ExcelExport2OutputStream<>(source, outputStreamProvider);
    }
}
