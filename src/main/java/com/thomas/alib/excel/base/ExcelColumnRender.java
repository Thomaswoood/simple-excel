package com.thomas.alib.excel.base;

import com.thomas.alib.excel.annotation.ExcelColumn;
import com.thomas.alib.excel.converter.*;
import com.thomas.alib.excel.utils.StringUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.lang.reflect.Field;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Date;

public class ExcelColumnRender implements AutoCloseable {
    private static final Logger logger = LoggerFactory.getLogger(ExcelColumnRender.class);
    /**
     * 列属性信息
     */
    protected Field columnField;
    /**
     * 列属性类型
     */
    protected Class<?> columnClass;
    /**
     * 列属性表格注解对象
     */
    protected ExcelColumn excelColumn;
    /**
     * 是否有效
     */
    protected boolean isValid;
    /**
     * 表头名
     */
    protected String headName;
    /**
     * 排序字段
     */
    protected int orderNum;
    /**
     * 默认前缀
     */
    protected String prefix;
    /**
     * 默认后缀
     */
    protected String suffix;
    /**
     * 列宽度
     */
    protected int columnWidth;
    /**
     * 属性值得转换器
     */
    protected Converter converter;
    /**
     * 是否按图片处理
     */
    protected boolean isPicture;
    /**
     * 列的index
     */
    private int columnIndex;

    public ExcelColumnRender(Field field) {
        this.columnField = field;
        init();
    }

    protected void init() {
        columnField.setAccessible(true);// 抑制Java的访问控制检查
        columnClass = columnField.getType();
        excelColumn = columnField.getAnnotation(ExcelColumn.class);
        if (excelColumn != null) {//包含column注解
            isValid = true;
            String item_header_name = excelColumn.headerName();//列表头名
            if (StringUtils.isEmpty(item_header_name))//为空默认置为字段名
                headName = columnField.getName();
            else//否则显示设置的表头名
                headName = item_header_name;
            orderNum = excelColumn.orderNum();//列排序字段
            columnIndex = excelColumn.orderNum();//列在excel中实际顺序字段(仅在import模式中使用，默认取用注解设置的，在读取表头时也会再根据表格实际情况再设置一次)
            columnWidth = excelColumn.columnWidth();//列宽度
            prefix = excelColumn.prefix();//默认前缀
            suffix = excelColumn.suffix();//默认后缀
            if (excelColumn.converter() == DefaultConverter.class) {//没有配置转换器，根据类型或者beforeConverter判断自动处理生成一个
                if (columnClass == LocalDateTime.class) {
                    converter = new LocalDateTimeConverter();
                } else if (columnClass == LocalDate.class) {
                    converter = new LocalDateConverter();
                } else if (columnClass == Date.class) {
                    converter = new DateConverter();
                } else if (excelColumn.beforeConvert().length > 0//如果设置了before和after则自动使用SimpleConvert
                        && excelColumn.afterConvert().length > 0
                        && excelColumn.beforeConvert().length == excelColumn.afterConvert().length) {
                    converter = new SimpleConvert(excelColumn.beforeConvert(), excelColumn.afterConvert());
                } else {
                    converter = new DefaultConverter();
                }
            } else {
                try {
                    if (excelColumn.converter().isEnum()) {//枚举类型不可创建，需要取出一个
                        converter = excelColumn.converter().getEnumConstants()[0];
                    } else {//其他类型则自动创建一个
                        converter = excelColumn.converter().newInstance();
                    }
                } catch (Throwable e) {//发生未知错误，使用默认转化器
                    logger.error("表格\"" + headName + "\"列-在创建您配置的converter时发生错误，将使用默认方式继续处理，具体错误为:", e);
                    converter = new DefaultConverter();
                }
            }
            isPicture = excelColumn.isPicture();//是否按图片处理
        } else isValid = false;
    }

    /**
     * 是否有效
     */
    public boolean isValid() {
        return isValid;
    }

    /**
     * 表头名
     */
    public String getHeadName() {
        return headName;
    }

    /**
     * 列宽度
     */
    public int getColumnWidth() {
        return columnWidth;
    }

    /**
     * 是否按图片处理
     */
    public boolean isPicture() {
        return isPicture;
    }

    /**
     * 获取列的index
     *
     * @return 列的index
     */
    public int getColumnIndex() {
        return columnIndex;
    }

    /**
     * 外部设置列的index
     *
     * @param columnIndex 列的index
     */
    public void setColumnIndex(int columnIndex) {
        this.columnIndex = columnIndex;
    }

    @Override
    public void close() {
        columnField = null;
        columnClass = null;
        excelColumn = null;
        isValid = false;
        headName = null;
        orderNum = 0;
        prefix = null;
        suffix = null;
        columnWidth = 0;
        converter = null;
        isPicture = false;
        columnIndex = 0;
    }
}
