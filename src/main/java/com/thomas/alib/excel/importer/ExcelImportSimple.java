package com.thomas.alib.excel.importer;

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
    static Logger logger = LoggerFactory.getLogger(ExcelImportSimple.class);

    /**
     * 读取excel文件总行数
     *
     * @param inputStream 上传的excel的文件流
     * @return 总行数
     */
    public static long getTotalLineCount(InputStream inputStream) {
        Workbook wb;
        try {
            wb = WorkbookFactory.create(inputStream);
        } catch (Exception e) {
            logger.error("文件读取失败", e);
            throw new RuntimeException("文件读取失败");
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
            logger.error("文件读取失败", e);
            throw new RuntimeException("文件读取失败");
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
        } catch (Throwable ignore) {
        }
        return totalLine;
    }

    /**
     * 读取excel文件为指定类型的数据列表
     *
     * @param inputStream 上传的excel的文件流
     * @param clazz       指定类型
     * @param <E>         指定类型泛型
     * @return 数据列表
     */
    public static <E> List<E> readList(InputStream inputStream, Class<E> clazz) {
        Workbook wb;
        try {
            wb = WorkbookFactory.create(inputStream);
        } catch (Exception e) {
            logger.error("文件读取失败", e);
            throw new RuntimeException("文件读取失败");
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
            logger.error("文件读取失败", e);
            throw new RuntimeException("文件读取失败");
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
        }
        try {
            wb.close();
        } catch (Throwable ignore) {
        }
        return result;
    }

    /**
     * 读取excel文件为指定类型的数据列表，安全模式，会把整个文件读完，并把错误信息返回而不是抛出异常
     *
     * @param inputStream 上传的excel的文件流
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
     * @param inputStream 上传的excel的文件流
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
            result.appendMsg("文件读取失败;");
            result.setReadSuccess(false);
            logger.error("文件读取失败", e);
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
            result.appendMsg("文件读取失败;");
            result.setReadSuccess(false);
            logger.error("文件读取失败", e);
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
        }
        try {
            wb.close();
        } catch (Throwable ignore) {
        }
        return result;
    }
}
