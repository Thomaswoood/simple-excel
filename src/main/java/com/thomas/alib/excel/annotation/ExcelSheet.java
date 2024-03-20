package com.thomas.alib.excel.annotation;

import com.thomas.alib.excel.enums.SortType;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 表示该类为一个Sheet的注解，通过此注解配置该sheet内的一些基础属性
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.TYPE})
public @interface ExcelSheet {

    /**
     * sheet名称
     */
    String sheetName() default "";

    /**
     * 是否显示序号
     */
    boolean showIndex() default true;

    /**
     * 列排序方式
     */
    SortType columnSortType() default SortType.S2B;

    /**
     * 表头行高度，默认-1不设置
     */
    short headRowHeight() default -1;

    /**
     * 数据行高度，默认-1不设置
     */
    short dataRowHeight() default -1;

    /**
     * 全局通用基础样式，使用优先级最低，仅导出时有作用
     */
    ExcelStyle baseStyle() default @ExcelStyle;

    /**
     * 表头样式，在{@link #baseStyle}基础之上设置，当与其定义了相同的样式属性时，以本注解中的为主，仅导出时有作用
     */
    ExcelStyle headStyle() default @ExcelStyle;

    /**
     * 数据样式，在{@link #baseStyle}基础之上设置，当与其定义了相同的样式属性时，以本注解中的为主，仅导出时有作用
     */
    ExcelStyle dataStyle() default @ExcelStyle;
}
