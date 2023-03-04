package com.thomas.alib.excel.base;

import com.thomas.alib.excel.annotation.ExcelColumn;
import com.thomas.alib.excel.converter.*;
import com.thomas.alib.excel.loader.PictureLoader;
import com.thomas.alib.excel.loader.PictureLoaderDefault;
import com.thomas.alib.excel.utils.StringUtils;

import java.lang.reflect.Field;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Date;

public class ExcelColumnRender {
    /**
     * 列属性信息
     */
    protected Field columnField;
    /**
     * 列属性类型
     */
    protected Class<?> columnClass;
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
     * 图片加载器
     */
    protected PictureLoader<?> pictureLoader;
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
        ExcelColumn excel_column = columnField.getAnnotation(ExcelColumn.class);
        if (excel_column != null) {//包含column注解
            isValid = true;
            String item_header_name = excel_column.headerName();//列表头名
            if (StringUtils.isEmpty(item_header_name))//为空默认置为字段名
                headName = columnField.getName();
            else//否则显示设置的表头名
                headName = item_header_name;
            orderNum = excel_column.orderNum();//列排序字段
            columnIndex = excel_column.orderNum();//列在excel中实际顺序字段(仅在import模式中使用，默认取用注解设置的，在读取表头时也会再根据表格实际情况再设置一次)
            columnWidth = excel_column.columnWidth();//列宽度
            prefix = excel_column.prefix();//默认前缀
            suffix = excel_column.suffix();//默认后缀
            if (excel_column.converter() == DefaultConverter.class) {//没有配置转换器，根据类型或者beforeConverter判断自动处理生成一个
                if (columnClass == LocalDateTime.class) {
                    converter = new LocalDateTimeConverter();
                } else if (columnClass == LocalDate.class) {
                    converter = new LocalDateConverter();
                } else if (columnClass == Date.class) {
                    converter = new DateConverter();
                } else if (excel_column.beforeConvert().length > 0//如果设置了before和after则自动使用SimpleConvert
                        && excel_column.afterConvert().length > 0
                        && excel_column.beforeConvert().length == excel_column.afterConvert().length) {
                    converter = new SimpleConvert(excel_column.beforeConvert(), excel_column.afterConvert());
                } else {
                    converter = new DefaultConverter();
                }
            } else {
                try {
                    if (excel_column.converter().isEnum()) {//枚举类型不可创建，需要取出一个
                        converter = excel_column.converter().getEnumConstants()[0];
                    } else {//其他类型则自动创建一个
                        converter = excel_column.converter().newInstance();
                    }
                } catch (Throwable e) {//发生未知错误，使用默认转化器
                    converter = new DefaultConverter();
                }
            }
            isPicture = excel_column.isPicture();//是否按图片处理
            if (isPicture) {
                if (excel_column.pictureLoader() == PictureLoaderDefault.class) {
                    //使用默认图片加载器
                    pictureLoader = new PictureLoaderDefault();
                } else {
                    //使用自定义图片加载器
                    try {
                        if (excel_column.pictureLoader().isEnum()) {//枚举类型不可创建，需要取出一个
                            pictureLoader = excel_column.pictureLoader().getEnumConstants()[0];
                        } else {//其他类型则自动创建一个
                            pictureLoader = excel_column.pictureLoader().newInstance();
                        }
                    } catch (Throwable e) {//发生未知错误，使用默认转化器
                        pictureLoader = new PictureLoaderDefault();
                    }
                }
            }
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
     * 获取图片加载器
     *
     * @return 图片加载器
     */
    public PictureLoader<?> getPictureLoader() {
        return pictureLoader;
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
}
