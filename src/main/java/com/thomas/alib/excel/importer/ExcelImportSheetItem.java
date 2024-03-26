package com.thomas.alib.excel.importer;

import com.thomas.alib.excel.exception.AnalysisException;
import com.thomas.alib.excel.importer.validation.ImportValidator;
import com.thomas.alib.excel.importer.validation.ImportValidatorFactory;
import com.thomas.alib.excel.utils.CollectionUtils;
import com.thomas.alib.excel.utils.ReflectUtil;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Set;
import java.util.function.Consumer;

/**
 * Excel导入时处理每个工作表(Sheet)的解析工具
 *
 * @param <T> 数据源对象类型泛型
 */
class ExcelImportSheetItem<T> implements AutoCloseable {
    private static Logger logger = LoggerFactory.getLogger(ExcelImportSheetItem.class);
    /**
     * excel工作簿对象
     */
    Workbook mWorkbook;
    /**
     * excel公式转换器
     */
    FormulaEvaluator formulaEvaluator;
    /**
     * 本次处理的sheet对象
     */
    Sheet mSheet;
    /**
     * 本次读取对应数据类的泛型
     */
    Class<T> dataClazz;
    /**
     * 根据sheet表格的表头解析的数据列Map
     */
    HashMap<String, ExcelImportColumnItem> columnFieldItemMap;
    /**
     * 本次处理的sheet所包含的总行数（包括所谓的表头行）
     */
    int lineCount;
    /**
     * 导入时数据验证器
     */
    ImportValidator importValidator;

    ExcelImportSheetItem(Workbook workbook, Sheet sheet, Class<T> clazz) {
        this.mWorkbook = workbook;
        this.formulaEvaluator = this.mWorkbook.getCreationHelper().createFormulaEvaluator();
        this.mSheet = sheet;
        this.dataClazz = clazz;
        this.columnFieldItemMap = new HashMap<>();
        this.lineCount = this.mSheet.getLastRowNum() - this.mSheet.getFirstRowNum();
        this.importValidator = ImportValidatorFactory.build();
    }

    /**
     * 读取sheet中的数据列表
     *
     * @return 数据列表
     */
    List<T> readList() {
        List<T> result = new ArrayList<>();
        int first_row = mSheet.getFirstRowNum(), last_row = mSheet.getLastRowNum();
        if (first_row == last_row) return result;//空sheet
        readHeadRow(mSheet.getRow(first_row++));
        for (int row_index = first_row; row_index <= last_row; ++row_index) {//遍历行
            Row row = mSheet.getRow(row_index);
            if (row == null) continue;//getRow是空说明这一行没有任何编辑项，没有被修改过。跳过
            T t;
            try {
                t = dataClazz.newInstance();
                boolean allValueIsNull = true;
                for (ExcelImportColumnItem column : columnFieldItemMap.values()) {
                    if (column.getColumnIndex() < 0) continue;
                    try {
                        Cell cell = row.getCell(column.getColumnIndex());
                        Object value = column.setColumnValueFromCell(t, cell, formulaEvaluator);
                        if (value != null) allValueIsNull = false;
                    } catch (Throwable e) {
                        throw new RuntimeException("表格\"" + column.getHeadName() + "\"列-第" + row_index + "行-解析时发生错误:", e);
                    }
                }
                if (allValueIsNull) {
                    continue;//解析所有字段都是空，这行数据跳过
                }
            } catch (Throwable e) {
                throw new RuntimeException("表格第" + row_index + "行-解析时发生错误:", e);
            }
            //验证器存在才进行验证判断
            if (importValidator != null) {
                Set<String> validateSet = importValidator.validate(t);
                if (!CollectionUtils.isEmpty(validateSet)) {
                    throw new RuntimeException("表格第" + row_index + "行:" + validateSet.iterator().next());
                }
            }
            result.add(t);
        }
        return result;
    }

    /**
     * 读取sheet中的数据列表，安全模式，会把整个文件读完，并把错误信息返回而不是抛出异常
     *
     * @param consumer 读取一条的回调
     * @return 数据列表
     */
    List<SafetyResult<T>> readListSafety(Consumer<SafetyResult<T>> consumer) {
        List<SafetyResult<T>> result = new ArrayList<>();
        int first_row = mSheet.getFirstRowNum(), last_row = mSheet.getLastRowNum();
        if (first_row == last_row) return result;//空sheet
        readHeadRow(mSheet.getRow(first_row++));
        for (int row_index = first_row; row_index <= last_row; ++row_index) {//遍历行
            Row row = mSheet.getRow(row_index);
            if (row == null) continue;//getRow是空说明这一行没有任何编辑项，没有被修改过。跳过
            SafetyResult<T> safety = new SafetyResult<>(row_index);
            T t;
            try {
                t = dataClazz.newInstance();
                boolean allValueIsNull = true;
                for (ExcelImportColumnItem column : columnFieldItemMap.values()) {
                    if (column.getColumnIndex() < 0) continue;
                    try {
                        Cell cell = row.getCell(column.getColumnIndex());
                        Object value = column.setColumnValueFromCell(t, cell, formulaEvaluator);
                        if (value != null) allValueIsNull = false;
                    } catch (Throwable e) {
                        String e_label = "表格\"" + column.getHeadName() + "\"列-第" + row_index + "行-解析时发生错误:";
                        logger.error(e_label, e);
                        safety.setReadSuccess(false);
                        if (e instanceof AnalysisException) {
                            safety.appendMsg(e_label + e.getLocalizedMessage() + ";");
                        } else {
                            safety.appendMsg(e_label + "未知错误，请联系管理员;");
                        }
                    }
                }
                if (allValueIsNull) {
                    continue;//解析所有字段都是空，这行数据跳过
                }
            } catch (Throwable e) {
                String e_label = "表格第" + row_index + "行-解析时发生错误:";
                logger.error(e_label, e);
                t = null;
                safety.setReadSuccess(false);
                safety.appendMsg(e_label + "未知错误，请联系管理员;");
            }
            safety.setData(t);
            //验证对象不为空，且验证器存在才进行验证判断
            if (t != null && importValidator != null) {
                Set<String> validateSet = importValidator.validate(t);
                if (!CollectionUtils.isEmpty(validateSet)) {
                    safety.setReadSuccess(false);
                    if (safety.getMsg().isEmpty()) {
                        safety.appendMsg("表格第" + row_index + "行:");
                    }
                    for (String validateMsg : validateSet) {
                        safety.appendMsg(validateMsg + ";");
                    }
                }
            }
            if (consumer != null) {
                try {
                    consumer.accept(safety);
                } catch (Throwable e) {
                    logger.error("表格第" + row_index + "行-回调时发生错误:", e);
                }
            }
            result.add(safety);
        }
        return result;
    }

    /**
     * 读取表头行（第一行）
     *
     * @param row 表头行
     */
    private void readHeadRow(Row row) {
        //全部的成员列表
        List<Field> total_field_list = ReflectUtil.getAccessibleFieldIncludeSuper(dataClazz);
        //遍历寻找所有有效column成员
        for (Field item_field : total_field_list) {
            ExcelImportColumnItem column = new ExcelImportColumnItem(item_field);
            if (column.isValid() && !column.isPicture()) {//有效且不按图片处理
                columnFieldItemMap.put(column.getHeadName(), column);
            }
        }
        int first_cell = row.getFirstCellNum(), last_cell = row.getLastCellNum();
        for (int cell_index = first_cell; cell_index < last_cell; ++cell_index) {//遍历表头行
            Cell cell = row.getCell(cell_index);
            ExcelImportColumnItem column = columnFieldItemMap.get(readHeadCell(cell));
            if (column != null) {
                column.setColumnIndex(cell_index);
            }
        }
    }

    /**
     * 读取表头项的值
     *
     * @param cell 表头项
     * @return 表头项的值
     */
    private String readHeadCell(Cell cell) {
        String column = "";
        CellType cellType = cell.getCellType();
        switch (cellType) {
            case NUMERIC:
                column = String.valueOf(cell.getDateCellValue());
                break;
            case STRING:
                column = cell.getStringCellValue();
                break;
            case FORMULA:
                column = cell.getCellFormula();
                break;
            case BOOLEAN:
                column = String.valueOf(cell.getBooleanCellValue());
                break;
            case ERROR:
            case BLANK:
                column = " ";
                break;
        }
        return column;
    }

    /**
     * 关闭，目前仅用于关闭验证器工厂
     */
    @Override
    public void close() {
        //如果验证器存在，关闭他
        if (importValidator != null) {
            try {
                importValidator.close();
            } catch (Throwable e) {
                logger.error("验证器工厂关闭时发生错误:", e);
            }
        }
    }
}
