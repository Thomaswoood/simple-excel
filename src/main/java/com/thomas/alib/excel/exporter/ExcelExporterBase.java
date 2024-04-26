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
    protected static Logger logger = LoggerFactory.getLogger(ExcelExporterBase.class);
    protected SXSSFWorkbook sxssfWorkbook;
    protected List<ExcelExportSheetItem<?, C>> sheetItemList;
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
     * @param data_list 数据源列表
     * @param <T>       数据类型泛型
     * @return 链式调用，返回自身
     */
    public <T> C addSheet(List<T> data_list) {
        return addSheet(data_list, null);
    }

    /**
     * 添加一个列表，作为一个sheet
     *
     * @param data_list  数据源列表
     * @param show_index 是否显示序号
     * @param <T>        数据类型泛型
     * @return 链式调用，返回自身
     */
    public <T> C addSheet(List<T> data_list, Boolean show_index) {
        return addSheet(data_list, show_index, null);
    }

    /**
     * 添加一个列表，作为一个sheet
     *
     * @param data_list  数据源列表
     * @param show_index 是否显示序号
     * @param sheet_name 指定sheet名称
     * @param <T>        数据类型泛型
     * @return 链式调用，返回自身
     */
    public <T> C addSheet(List<T> data_list, Boolean show_index, String sheet_name) {
        //空数据不操作
        if (CollectionUtils.isEmpty(data_list)) return child;
        //获取数据类型
        Class<T> data_clazz = null;
        for (T item : data_list) {
            if (item != null) {
                try {
                    data_clazz = (Class<T>) item.getClass();
                    break;
                } catch (Throwable ignored) {
                }
            }
        }
        return createSheet(data_list, data_clazz, show_index, sheet_name);
    }

    /**
     * 添加一个仅有表头的空sheet
     *
     * @param data_clazz 数据源类型
     * @param <T>        数据类型泛型
     * @return 链式调用，返回自身
     */
    public <T> C createEmptySheet(Class<T> data_clazz) {
        return createSheet(null, data_clazz);
    }

    /**
     * 添加一个仅有表头的空sheet
     *
     * @param data_clazz 数据源类型
     * @param show_index 是否显示序号
     * @param <T>        数据类型泛型
     * @return 链式调用，返回自身
     */
    public <T> C createEmptySheet(Class<T> data_clazz, Boolean show_index) {
        return createSheet(null, data_clazz, show_index);
    }

    /**
     * 添加一个仅有表头的空sheet
     *
     * @param data_clazz 数据源类型
     * @param show_index 是否显示序号
     * @param sheet_name 指定sheet名称
     * @param <T>        数据类型泛型
     * @return 链式调用，返回自身
     */
    public <T> C createEmptySheet(Class<T> data_clazz, Boolean show_index, String sheet_name) {
        return createSheet(null, data_clazz, show_index, sheet_name);
    }

    /**
     * 添加一个列表，作为一个sheet
     *
     * @param data_list  数据源列表
     * @param data_clazz 数据源类型
     * @param <T>        数据类型泛型
     * @return 链式调用，返回自身
     */
    public <T> C createSheet(List<T> data_list, Class<T> data_clazz) {
        return createSheet(data_list, data_clazz, null);
    }

    /**
     * 添加一个列表，作为一个sheet
     *
     * @param data_list  数据源列表
     * @param data_clazz 数据源类型
     * @param show_index 是否显示序号
     * @param <T>        数据类型泛型
     * @return 链式调用，返回自身
     */
    public <T> C createSheet(List<T> data_list, Class<T> data_clazz, Boolean show_index) {
        return createSheet(data_list, data_clazz, show_index, null);
    }

    /**
     * 添加一个列表，作为一个sheet
     *
     * @param data_list  数据源列表
     * @param data_clazz 数据源类型
     * @param show_index 是否显示序号
     * @param sheet_name 指定sheet名称
     * @param <T>        数据类型泛型
     * @return 链式调用，返回自身
     */
    public <T> C createSheet(List<T> data_list, Class<T> data_clazz, Boolean show_index, String sheet_name) {
        currentSheetItem = new ExcelExportSheetItem<>(child, data_list, data_clazz, show_index, sheet_name);
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
     * 开始生成并输出返回，导出后将自动销毁
     */
    public void export() {
        export(true);
    }

    /**
     * 开始生成并输出返回
     *
     * @param auto_destroy 是否自动销毁
     */
    public void export(boolean auto_destroy) {
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
            } catch (IOException e) {
                logger.error("表格export关闭输出流时发生错误:", e);
            }
            if (auto_destroy) destroy();
            logger.debug("导出数据输出完毕。");
        }
    }

    /**
     * 销毁
     */
    public void destroy() {
        try {
            sxssfWorkbook.close();
        } catch (Throwable e) {
            logger.error("表格关闭Workbook时发生错误:", e);
        }
        sxssfWorkbook = null;
        sheetItemList.clear();
        sheetItemList = null;
        currentSheetItem = null;
        child = null;
    }
}
