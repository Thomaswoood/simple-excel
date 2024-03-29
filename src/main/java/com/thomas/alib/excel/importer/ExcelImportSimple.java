package com.thomas.alib.excel.importer;

import com.thomas.alib.excel.interfaces.EFunction;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.function.BiConsumer;

/**
 * Excel导入工具类
 */
public class ExcelImportSimple {
    private static Logger logger = LoggerFactory.getLogger(ExcelImportSimple.class);

    /**
     * 读取excel文件总行数
     *
     * @param inputStreamGetFunc 获取excel文件流的方法
     * @param inputStreamSource  获取excel文件流的源对象
     * @param <S>                获取excel文件流的源对象的泛型
     * @return 总行数
     */
    public static <S> long getTotalLineCount(EFunction<S, InputStream> inputStreamGetFunc, S inputStreamSource) {
        Workbook wb;
        try {
            wb = WorkbookFactory.create(inputStreamGetFunc.apply(inputStreamSource));
        } catch (Exception e) {
            throw new RuntimeException("文件读取失败", e);
        }
        return getTotalLineCount(wb);
    }

    /**
     * 读取excel文件总行数
     *
     * @param inputStream excel的文件流
     * @return 总行数
     */
    public static long getTotalLineCount(InputStream inputStream) {
        Workbook wb;
        try {
            wb = WorkbookFactory.create(inputStream);
        } catch (Exception e) {
            throw new RuntimeException("文件读取失败", e);
        }
        return getTotalLineCount(wb);
    }

    /**
     * 读取excel文件总行数
     *
     * @param file excel文件
     * @return 总行数
     */
    public static long getTotalLineCount(File file) {
        Workbook wb;
        try {
            wb = WorkbookFactory.create(file);
        } catch (Exception e) {
            throw new RuntimeException("文件读取失败", e);
        }
        return getTotalLineCount(wb);
    }

    /**
     * 读取excel文件总行数
     *
     * @param wb excel文件
     * @return 总行数
     */
    private static long getTotalLineCount(Workbook wb) {
        long totalLine = 0L;
        for (int i = 0; i < wb.getNumberOfSheets(); ++i) {//遍历sheet
            Sheet sheet = wb.getSheetAt(i);
            int first_row = sheet.getFirstRowNum(), last_row = sheet.getLastRowNum();
            totalLine = totalLine + last_row - first_row;
        }
        try {
            wb.close();
        } catch (Throwable e) {
            logger.error("表格关闭Workbook时发生错误:", e);
        }
        return totalLine;
    }

    /**
     * 读取excel文件为指定类型的数据列表
     *
     * @param inputStreamGetFunc 获取excel文件流的方法
     * @param inputStreamSource  获取excel文件流的源对象
     * @param clazz              指定类型
     * @param <S>                获取excel文件流的源对象的泛型
     * @param <E>                指定类型泛型
     * @return 数据列表
     */
    public static <S, E> List<E> readList(EFunction<S, InputStream> inputStreamGetFunc, S inputStreamSource, Class<E> clazz) {
        Workbook wb;
        try {
            wb = WorkbookFactory.create(inputStreamGetFunc.apply(inputStreamSource));
        } catch (Exception e) {
            throw new RuntimeException("文件读取失败", e);
        }
        return readList(wb, clazz);
    }

    /**
     * 读取excel文件为指定类型的数据列表
     *
     * @param inputStream excel的文件流
     * @param clazz       指定类型
     * @param <E>         指定类型泛型
     * @return 数据列表
     */
    public static <E> List<E> readList(InputStream inputStream, Class<E> clazz) {
        Workbook wb;
        try {
            wb = WorkbookFactory.create(inputStream);
        } catch (Exception e) {
            throw new RuntimeException("文件读取失败", e);
        }
        return readList(wb, clazz);
    }

    /**
     * 读取excel文件为指定类型的数据列表
     *
     * @param file  excel文件
     * @param clazz 指定类型
     * @param <E>   指定类型泛型
     * @return 数据列表
     */
    public static <E> List<E> readList(File file, Class<E> clazz) {
        Workbook wb;
        try {
            wb = WorkbookFactory.create(file);
        } catch (Exception e) {
            throw new RuntimeException("文件读取失败", e);
        }
        return readList(wb, clazz);
    }

    /**
     * 读取excel文件为指定类型的数据列表
     *
     * @param wb    excel工作簿对象
     * @param clazz 指定类型
     * @param <E>   指定类型泛型
     * @return 数据列表
     */
    private static <E> List<E> readList(Workbook wb, Class<E> clazz) {
        List<E> result = new ArrayList<>();
        for (int i = 0; i < wb.getNumberOfSheets(); ++i) {//遍历sheet
            ExcelImportSheetItem<E> sheet = new ExcelImportSheetItem<>(wb, wb.getSheetAt(i), clazz);
            result.addAll(sheet.readList());
            sheet.close();
        }
        try {
            wb.close();
        } catch (Throwable e) {
            logger.error("表格关闭Workbook时发生错误:", e);
        }
        return result;
    }

    /**
     * 读取excel文件为指定类型的数据列表，安全模式，会把整个文件读完，并把错误信息返回而不是抛出异常
     *
     * @param inputStreamGetFunc 获取excel文件流的方法
     * @param inputStreamSource  获取excel文件流的源对象
     * @param clazz              指定类型
     * @param <S>                获取excel文件流的源对象的泛型
     * @param <E>                指定类型泛型
     * @return 数据列表
     */
    public static <S, E> SafetyResult<List<SafetyResult<E>>> readListSafety(EFunction<S, InputStream> inputStreamGetFunc, S inputStreamSource, Class<E> clazz) {
        return readListSafety(inputStreamGetFunc, inputStreamSource, clazz, null);
    }

    /**
     * 读取excel文件为指定类型的数据列表，安全模式，会把整个文件读完，并把错误信息返回而不是抛出异常
     *
     * @param inputStream excel的文件流
     * @param clazz       指定类型
     * @param <E>         指定类型泛型
     * @return 数据列表
     */
    public static <E> SafetyResult<List<SafetyResult<E>>> readListSafety(InputStream inputStream, Class<E> clazz) {
        return readListSafety(inputStream, clazz, null);
    }

    /**
     * 读取excel文件为指定类型的数据列表，安全模式，会把整个文件读完，并把错误信息返回而不是抛出异常
     *
     * @param file  excel文件
     * @param clazz 指定类型
     * @param <E>   指定类型泛型
     * @return 数据列表
     */
    public static <E> SafetyResult<List<SafetyResult<E>>> readListSafety(File file, Class<E> clazz) {
        return readListSafety(file, clazz, null);
    }

    /**
     * 读取excel文件为指定类型的数据列表，安全模式，会把整个文件读完，并把错误信息返回而不是抛出异常
     *
     * @param inputStreamGetFunc 获取excel文件流的方法
     * @param inputStreamSource  获取excel文件流的源对象
     * @param clazz              指定类型
     * @param consumer           读取到每一行数据事件的消费者
     * @param <S>                获取excel文件流的源对象的泛型
     * @param <E>                指定类型泛型
     * @return 数据列表
     */
    public static <S, E> SafetyResult<List<SafetyResult<E>>> readListSafety(EFunction<S, InputStream> inputStreamGetFunc, S inputStreamSource, Class<E> clazz, BiConsumer<Long, SafetyResult<E>> consumer) {
        SafetyResult<List<SafetyResult<E>>> result = new SafetyResult<>();
        result.setData(new ArrayList<>());
        Workbook wb;
        try {
            wb = WorkbookFactory.create(inputStreamGetFunc.apply(inputStreamSource));
        } catch (Exception e) {
            logger.error("表格安全模式下，根据输入信息读取Workbook对象时发生错误:", e);
            result.appendMsg("文件读取失败;" + e.getLocalizedMessage());
            result.setReadSuccess(false);
            return result;
        }
        return readListSafety(wb, clazz, consumer, result);
    }

    /**
     * 读取excel文件为指定类型的数据列表，安全模式，会把整个文件读完，并把错误信息返回而不是抛出异常
     *
     * @param inputStream excel的文件流
     * @param clazz       指定类型
     * @param consumer    读取到每一行数据事件的消费者
     * @param <E>         指定类型泛型
     * @return 数据列表
     */
    public static <E> SafetyResult<List<SafetyResult<E>>> readListSafety(InputStream inputStream, Class<E> clazz, BiConsumer<Long, SafetyResult<E>> consumer) {
        SafetyResult<List<SafetyResult<E>>> result = new SafetyResult<>();
        result.setData(new ArrayList<>());
        Workbook wb;
        try {
            wb = WorkbookFactory.create(inputStream);
        } catch (Exception e) {
            logger.error("表格安全模式下，根据输入流读取Workbook对象时发生错误:", e);
            result.appendMsg("文件读取失败;" + e.getLocalizedMessage());
            result.setReadSuccess(false);
            return result;
        }
        return readListSafety(wb, clazz, consumer, result);
    }

    /**
     * 读取excel文件为指定类型的数据列表，安全模式，会把整个文件读完，并把错误信息返回而不是抛出异常
     *
     * @param file     excel文件
     * @param clazz    指定类型
     * @param consumer 读取到每一行数据事件的消费者
     * @param <E>      指定类型泛型
     * @return 数据列表
     */
    public static <E> SafetyResult<List<SafetyResult<E>>> readListSafety(File file, Class<E> clazz, BiConsumer<Long, SafetyResult<E>> consumer) {
        SafetyResult<List<SafetyResult<E>>> result = new SafetyResult<>();
        result.setData(new ArrayList<>());
        Workbook wb;
        try {
            wb = WorkbookFactory.create(file);
        } catch (Exception e) {
            logger.error("表格安全模式下，根据输入文件读取Workbook对象时发生错误:", e);
            result.appendMsg("文件读取失败;" + e.getLocalizedMessage());
            result.setReadSuccess(false);
            return result;
        }
        return readListSafety(wb, clazz, consumer, result);
    }

    /**
     * 读取excel文件为指定类型的数据列表，安全模式，会把整个文件读完，并把错误信息返回而不是抛出异常
     *
     * @param wb       excel工作簿对象
     * @param clazz    指定类型
     * @param consumer 读取到每一行数据事件的消费者
     * @param <E>      指定类型泛型
     * @return 数据列表
     */
    private static <E> SafetyResult<List<SafetyResult<E>>> readListSafety(Workbook wb, Class<E> clazz, BiConsumer<Long, SafetyResult<E>> consumer, SafetyResult<List<SafetyResult<E>>> result) {
        List<ExcelImportSheetItem<E>> sheetItemList = new ArrayList<>();
        long totalLine = 0L;
        for (int i = 0; i < wb.getNumberOfSheets(); ++i) {//遍历sheet
            ExcelImportSheetItem<E> sheet = new ExcelImportSheetItem<>(wb, wb.getSheetAt(i), clazz);
            sheetItemList.add(sheet);
            totalLine += sheet.lineCount;
        }
        for (ExcelImportSheetItem<E> sheet : sheetItemList) {
            long finalTotalLine = totalLine;
            result.getData().addAll(sheet.readListSafety(eSafetyResult -> consumer.accept(finalTotalLine, eSafetyResult)));
            sheet.close();
        }
        try {
            wb.close();
        } catch (Throwable e) {
            logger.error("表格关闭Workbook时发生错误:", e);
        }
        return result;
    }
}
