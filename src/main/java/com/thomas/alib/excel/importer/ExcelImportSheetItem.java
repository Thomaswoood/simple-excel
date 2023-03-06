package com.thomas.alib.excel.importer;

import com.thomas.alib.excel.exception.AnalysisException;
import com.thomas.alib.excel.utils.CollectionUtils;
import com.thomas.alib.excel.utils.ReflectUtil;
import org.apache.poi.ss.usermodel.*;

import javax.validation.ConstraintViolation;
import javax.validation.Valid;
import javax.validation.Validation;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Set;
import java.util.function.Consumer;

class ExcelImportSheetItem<T> {
    Workbook mWorkbook;
    FormulaEvaluator formulaEvaluator;
    Sheet mSheet;
    Class<T> dataClazz;
    HashMap<String, ExcelImportColumnItem> columnFieldItemMap;
    int lineCount;

    ExcelImportSheetItem(Workbook workbook, Sheet sheet, Class<T> clazz) {
        this.mWorkbook = workbook;
        this.formulaEvaluator = this.mWorkbook.getCreationHelper().createFormulaEvaluator();
        this.mSheet = sheet;
        this.dataClazz = clazz;
        this.columnFieldItemMap = new HashMap<>();
        this.lineCount = this.mSheet.getLastRowNum() - this.mSheet.getFirstRowNum();
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
                for (ExcelImportColumnItem item : columnFieldItemMap.values()) {
                    if (item.getColumnIndex() < 0) continue;
                    Cell cell = row.getCell(item.getColumnIndex());
                    Object value = item.setColumnValueFromCell(t, cell, formulaEvaluator);
                    if (value != null) allValueIsNull = false;
                }
                if (allValueIsNull) {
                    continue;//解析所有字段都是空，这行数据跳过
                }
            } catch (Throwable e) {
                if (e instanceof AnalysisException) {
                    throw new RuntimeException("表格第" + row_index + "行:" + e.getLocalizedMessage());
                } else {
                    e.printStackTrace();
                    throw new RuntimeException("表格第" + row_index + "行:解析时发生未知错误，请联系管理员", e);
                }
            }
            Set<ConstraintViolation<@Valid T>> validateSet = Validation.buildDefaultValidatorFactory()
                    .getValidator()
                    .validate(t);
            if (!CollectionUtils.isEmpty(validateSet)) {
                throw new RuntimeException("表格第" + row_index + "行:" + validateSet.iterator().next().getMessage());
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
                for (ExcelImportColumnItem item : columnFieldItemMap.values()) {
                    if (item.getColumnIndex() < 0) continue;
                    try {
                        Cell cell = row.getCell(item.getColumnIndex());
                        Object value = item.setColumnValueFromCell(t, cell, formulaEvaluator);
                        if (value != null) allValueIsNull = false;
                    } catch (Throwable e) {
                        safety.setReadSuccess(false);
                        if (e instanceof AnalysisException) {
                            safety.appendMsg("表格第" + row_index + "行:" + e.getLocalizedMessage() + ";");
                        } else {
                            safety.appendMsg("表格第" + row_index + "行:解析时发生未知错误，请联系管理员;");
                        }
                    }
                }
                if (allValueIsNull) {
                    continue;//解析所有字段都是空，这行数据跳过
                }
            } catch (Throwable e) {
                t = null;
                safety.setReadSuccess(false);
                safety.appendMsg("表格第" + row_index + "行:解析时发生未知错误，请联系管理员;");
            }
            safety.setData(t);
            if (t != null) {
                Set<ConstraintViolation<@Valid T>> validateSet = Validation.buildDefaultValidatorFactory()
                        .getValidator()
                        .validate(t);
                if (!CollectionUtils.isEmpty(validateSet)) {
                    safety.setReadSuccess(false);
                    if (safety.getMsg().isEmpty()) {
                        safety.appendMsg("表格第" + row_index + "行:");
                    }
                    for (ConstraintViolation<@Valid T> validate : validateSet) {
                        safety.appendMsg(validate.getMessage() + ";");
                    }
                }
            }
            if (consumer != null) {
                try {
                    consumer.accept(safety);
                } catch (Throwable ignore) {
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
            ExcelImportColumnItem item_column_field = new ExcelImportColumnItem(item_field);
            if (item_column_field.isValid() && !item_column_field.isPicture()) {//有效且不按图片处理
                columnFieldItemMap.put(item_column_field.getHeadName(), item_column_field);
            }
        }
        int first_cell = row.getFirstCellNum(), last_cell = row.getLastCellNum();
        for (int cell_index = first_cell; cell_index < last_cell; ++cell_index) {//遍历表头行
            Cell cell = row.getCell(cell_index);
            ExcelImportColumnItem item_column_field = columnFieldItemMap.get(readHeadCell(cell));
            if (item_column_field != null) {
                item_column_field.setColumnIndex(cell_index);
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
}
