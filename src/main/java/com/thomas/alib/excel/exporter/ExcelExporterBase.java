package com.thomas.alib.excel.exporter;


import com.thomas.alib.excel.utils.CollectionUtils;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * Excel导出者基类，主要负责处理输出流和poi之间的交互
 *
 * @param <C> 子类泛型，用于链式调用
 */
public abstract class ExcelExporterBase<C extends ExcelExporterBase<C>> {
    private static Logger logger = LoggerFactory.getLogger(ExcelExporterBase.class);
    final SXSSFWorkbook sxssfWorkbook;
    final List<ExcelExportSheetItem<?, C>> sheetItemList;
    protected ExcelExportSheetItem<?, C> currentSheetItem;
    protected C child;

    /**
     * 构造方法
     */
    protected ExcelExporterBase() {
        this.sxssfWorkbook = new SXSSFWorkbook();
        this.sheetItemList = new ArrayList<>();
    }

    /**
     * 添加一个列表，作为一个sheet
     *
     * @param dataList 数据源列表
     * @param <T>      数据类型泛型
     * @return 链式调用，返回自身
     */
    public <T> C addSheet(List<T> dataList) {
        return addSheet(dataList, null);
    }

    /**
     * 添加一个列表，作为一个sheet
     *
     * @param dataList   数据源列表
     * @param show_index 是否显示序号
     * @param <T>        数据类型泛型
     * @return 链式调用，返回自身
     */
    public <T> C addSheet(List<T> dataList, Boolean show_index) {
        return addSheet(dataList, show_index, null);
    }

    /**
     * 添加一个列表，作为一个sheet
     *
     * @param dataList   数据源列表
     * @param show_index 是否显示序号
     * @param sheet_name 指定sheet名称
     * @param <T>        数据类型泛型
     * @return 链式调用，返回自身
     */
    public <T> C addSheet(List<T> dataList, Boolean show_index, String sheet_name) {
        //空数据不操作
        if (CollectionUtils.isEmpty(dataList)) return child;
        //获取数据类型
        Class<T> dataClazz = null;
        try {
            for (T item : dataList) {
                dataClazz = (Class<T>) item.getClass();
                if (dataClazz != null) break;
            }
        } catch (Throwable ignored) {
        }
        return createSheet(dataList, dataClazz, show_index, sheet_name);
    }

    /**
     * 添加一个仅有表头的空sheet
     *
     * @param dataClazz 数据源类型
     * @param <T>       数据类型泛型
     * @return 链式调用，返回自身
     */
    public <T> C createEmptySheet(Class<T> dataClazz) {
        return createSheet(null, dataClazz);
    }

    /**
     * 添加一个仅有表头的空sheet
     *
     * @param dataClazz  数据源类型
     * @param show_index 是否显示序号
     * @param <T>        数据类型泛型
     * @return 链式调用，返回自身
     */
    public <T> C createEmptySheet(Class<T> dataClazz, Boolean show_index) {
        return createSheet(null, dataClazz, show_index);
    }

    /**
     * 添加一个仅有表头的空sheet
     *
     * @param dataClazz  数据源类型
     * @param show_index 是否显示序号
     * @param sheet_name 指定sheet名称
     * @param <T>        数据类型泛型
     * @return 链式调用，返回自身
     */
    public <T> C createEmptySheet(Class<T> dataClazz, Boolean show_index, String sheet_name) {
        return createSheet(null, dataClazz, show_index, sheet_name);
    }

    /**
     * 添加一个列表，作为一个sheet
     *
     * @param dataList  数据源列表
     * @param dataClazz 数据源类型
     * @param <T>       数据类型泛型
     * @return 链式调用，返回自身
     */
    public <T> C createSheet(List<T> dataList, Class<T> dataClazz) {
        return createSheet(dataList, dataClazz, null);
    }

    /**
     * 添加一个列表，作为一个sheet
     *
     * @param dataList   数据源列表
     * @param dataClazz  数据源类型
     * @param show_index 是否显示序号
     * @param <T>        数据类型泛型
     * @return 链式调用，返回自身
     */
    public <T> C createSheet(List<T> dataList, Class<T> dataClazz, Boolean show_index) {
        return createSheet(dataList, dataClazz, show_index, null);
    }

    /**
     * 添加一个列表，作为一个sheet
     *
     * @param dataList   数据源列表
     * @param dataClazz  数据源类型
     * @param show_index 是否显示序号
     * @param sheet_name 指定sheet名称
     * @param <T>        数据类型泛型
     * @return 链式调用，返回自身
     */
    public <T> C createSheet(List<T> dataList, Class<T> dataClazz, Boolean show_index, String sheet_name) {
        currentSheetItem = new ExcelExportSheetItem<>(child, dataList, dataClazz, show_index, sheet_name);
        currentSheetItem.writeData();
        sheetItemList.add(currentSheetItem);
        return child;
    }

    /**
     * 取得当前创建的sheet页签对象，可以调用sheet中的某些方法
     * （目前暂无实际业务场景，为未来扩展预留的方法）
     *
     * @return 当前创建的sheet页签对象
     */
    public ExcelExportSheetItem<?, C> currentSheet() {
        return currentSheetItem;
    }

    /**
     * 由子类提供具体导出的输出流
     *
     * @return 输出流对象
     * @throws Exception 创建输出流时可能产生的异常，子类可直接抛出，由父类处理
     */
    protected abstract OutputStream getOutputStream() throws Exception;

    /**
     * 开始生成并输出返回
     */
    public void export() {
        if (CollectionUtils.isEmpty(sheetItemList))
            throw new RuntimeException("导出数据为空");
        logger.debug("导出数据处理完毕，开始输出到指定的流中。");
        OutputStream os = null;
        try {
            os = getOutputStream();
            sxssfWorkbook.write(os);
        } catch (Exception e) {
            throw new RuntimeException("导出失败", e);
        } finally {
            try {
                if (os != null) {
                    os.flush();
                    os.close();
                }
                logger.debug("导出数据输出完毕。");
            } catch (IOException e) {
                logger.error("表格export关闭输出流时发生错误:", e);
                System.err.println(e.getMessage());
            }

        }
    }
}
