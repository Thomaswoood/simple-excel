package com.thomas.alib.excel.importer;

import com.thomas.alib.excel.exception.base.ExcelColumnRender;
import com.thomas.alib.excel.exception.AnalysisException;
import com.thomas.alib.excel.utils.StringUtils;
import org.apache.poi.ss.formula.eval.NotImplementedException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Date;

/**
 * Excel导入对应每个列(Column)的解析工具
 */
class ExcelImportColumnItem extends ExcelColumnRender {

    ExcelImportColumnItem(Field field) {
        super(field);
    }

    /**
     * 向数据源中设置本列的值
     *
     * @param source           数据源
     * @param cell             单元格信息
     * @param formulaEvaluator 公式转换器
     * @throws AnalysisException 解析异常
     */
    Object setColumnValueFromCell(Object source, Cell cell, FormulaEvaluator formulaEvaluator) throws AnalysisException {
        try {
            Object value = null;
            if (cell == null) {
                value = converter.inverseConvert(null);
            } else {
                CellType cellType = cell.getCellType();
                switch (cellType) {
                    case NUMERIC:
                        if (columnClass == Date.class || columnClass == LocalDate.class || columnClass == LocalDateTime.class) {
                            value = converter.inverseConvert(cell.getDateCellValue());
                        } else {
                            value = converter.inverseConvert(cell.getNumericCellValue());
                            if (columnClass == String.class) {
                                value = new BigDecimal(String.valueOf(value)).stripTrailingZeros().toPlainString();
                            } else if (columnClass == int.class || columnClass == Integer.class) {
                                value = Double.valueOf(String.valueOf(value)).intValue();
                            } else if (columnClass == long.class || columnClass == Long.class) {
                                value = Double.valueOf(String.valueOf(value)).longValue();
                            } else if (columnClass == short.class || columnClass == Short.class) {
                                value = Double.valueOf(String.valueOf(value)).shortValue();
                            } else if (columnClass == double.class || columnClass == Double.class) {
                                value = Double.valueOf(String.valueOf(value));
                            } else if (columnClass == float.class || columnClass == Float.class) {
                                value = Double.valueOf(String.valueOf(value)).floatValue();
                            } else if (columnClass == BigDecimal.class) {
                                value = new BigDecimal(String.valueOf(value));
                            }
                        }
                        break;
                    case STRING:
                        String stringValue = cell.getStringCellValue();
                        if (StringUtils.isNotEmpty(prefix) && stringValue.startsWith(prefix))
                            stringValue = stringValue.substring(prefix.length());
                        if (StringUtils.isNotEmpty(suffix) && stringValue.endsWith(suffix))
                            stringValue = stringValue.substring(0, stringValue.length() - suffix.length());
                        value = converter.inverseConvert(stringValue);
                        if (columnClass == int.class || columnClass == Integer.class) {
                            if (value != null) value = Integer.parseInt(String.valueOf(value));
                        } else if (columnClass == long.class || columnClass == Long.class) {
                            if (value != null) value = Long.parseLong(String.valueOf(value));
                        } else if (columnClass == short.class || columnClass == Short.class) {
                            if (value != null) value = Short.parseShort(String.valueOf(value));
                        } else if (columnClass == double.class || columnClass == Double.class) {
                            if (value != null) value = Double.parseDouble(String.valueOf(value));
                        } else if (columnClass == float.class || columnClass == Float.class) {
                            if (value != null) value = Float.parseFloat(String.valueOf(value));
                        } else if (columnClass == BigDecimal.class) {
                            if (value != null) value = new BigDecimal(String.valueOf(value));
                        }
                        break;
                    case FORMULA://公式，2023-02-04补充支持
                        CellValue formulaCellValue = formulaEvaluator.evaluate(cell);
                        switch (formulaCellValue.getCellType()) {
                            case NUMERIC:
                                value = converter.inverseConvert(formulaCellValue.getNumberValue());
                                if (columnClass == String.class) {
                                    value = new BigDecimal(String.valueOf(value)).stripTrailingZeros().toPlainString();
                                } else if (columnClass == int.class || columnClass == Integer.class) {
                                    value = Double.valueOf(String.valueOf(value)).intValue();
                                } else if (columnClass == long.class || columnClass == Long.class) {
                                    value = Double.valueOf(String.valueOf(value)).longValue();
                                } else if (columnClass == short.class || columnClass == Short.class) {
                                    value = Double.valueOf(String.valueOf(value)).shortValue();
                                } else if (columnClass == double.class || columnClass == Double.class) {
                                    value = Double.valueOf(String.valueOf(value));
                                } else if (columnClass == float.class || columnClass == Float.class) {
                                    value = Double.valueOf(String.valueOf(value)).floatValue();
                                } else if (columnClass == BigDecimal.class) {
                                    value = new BigDecimal(String.valueOf(value));
                                }
                                break;
                            case STRING:
                                String formulaCellValueStringValue = formulaCellValue.getStringValue();
                                if (StringUtils.isNotEmpty(prefix) && formulaCellValueStringValue.startsWith(prefix))
                                    formulaCellValueStringValue = formulaCellValueStringValue.substring(prefix.length());
                                if (StringUtils.isNotEmpty(suffix) && formulaCellValueStringValue.endsWith(suffix))
                                    formulaCellValueStringValue = formulaCellValueStringValue.substring(0, formulaCellValueStringValue.length() - suffix.length());
                                value = converter.inverseConvert(formulaCellValueStringValue);
                                if (columnClass == int.class || columnClass == Integer.class) {
                                    if (value != null) value = Integer.parseInt(String.valueOf(value));
                                } else if (columnClass == long.class || columnClass == Long.class) {
                                    if (value != null) value = Long.parseLong(String.valueOf(value));
                                } else if (columnClass == short.class || columnClass == Short.class) {
                                    if (value != null) value = Short.parseShort(String.valueOf(value));
                                } else if (columnClass == double.class || columnClass == Double.class) {
                                    if (value != null) value = Double.parseDouble(String.valueOf(value));
                                } else if (columnClass == float.class || columnClass == Float.class) {
                                    if (value != null) value = Float.parseFloat(String.valueOf(value));
                                } else if (columnClass == BigDecimal.class) {
                                    if (value != null) value = new BigDecimal(String.valueOf(value));
                                }
                                break;
                            case BOOLEAN:
                                if (columnClass == String.class) {
                                    value = String.valueOf(formulaCellValue.getBooleanValue());
                                } else if (columnClass == Number.class) {
                                    value = formulaCellValue.getBooleanValue() ? 1 : 0;
                                } else {
                                    value = formulaCellValue.getBooleanValue();
                                }
                                break;
                            case ERROR:
                            case BLANK:
                                value = converter.inverseConvert(null);
                                break;
                        }
                        break;
                    case BOOLEAN:
                        if (columnClass == String.class) {
                            value = String.valueOf(cell.getBooleanCellValue());
                        } else if (columnClass == Number.class) {
                            value = cell.getBooleanCellValue() ? 1 : 0;
                        } else {
                            value = cell.getBooleanCellValue();
                        }
                        break;
                    case ERROR:
                    case BLANK:
                        value = converter.inverseConvert(null);
                        break;
                }
            }
            if (source != null && value != null) columnField.set(source, value);
            return value;
        } catch (Throwable e) {
            if (e instanceof AnalysisException) {
                throw new AnalysisException(headName + "-" + e.getLocalizedMessage());
            } else if (e instanceof NotImplementedException) {
                throw new AnalysisException(headName + "-该列公式暂不支持，无法解析");
            } else {
                //成员值无效
                throw new AnalysisException(headName + "-无法解析");
            }
        }
    }
}
