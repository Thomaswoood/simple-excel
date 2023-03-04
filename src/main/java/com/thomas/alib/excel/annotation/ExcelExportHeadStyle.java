package com.thomas.alib.excel.annotation;

import com.thomas.alib.excel.enums.BorderType;
import com.thomas.alib.excel.enums.HorAlignment;
import com.thomas.alib.excel.enums.VerAlignment;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 导出时表头行的样式注解
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.TYPE})
public @interface ExcelExportHeadStyle {

    /**
     * 字体名称，默认为空不设置
     */
    String fontName() default "";

    /**
     * 文字大小，默认0不设置
     */
    short textSize() default 0;

    /**
     * 文字是否加粗，默认不加粗
     */
    boolean isBold() default false;

    /**
     * 文字是否斜体，默认不斜体
     */
    boolean isItalic() default false;

    /**
     * 文字是否划线，默认不划线
     */
    boolean isStrikeout() default false;

    /**
     * 文字颜色，默认-1不设置
     */
    short textColor() default -1;

    /**
     * 水平对齐方式，默认自动对齐
     */
    HorAlignment horAlignment() default HorAlignment.GENERAL;

    /**
     * 垂直对齐方式，默认垂直居中
     */
    VerAlignment verAlignment() default VerAlignment.CENTER;

    /**
     * 文字是否自动换行，默认否
     */
    boolean isWrapText() default false;

    /**
     * 4边边框样式，默认无边框
     */
    BorderType fourSideBorder() default BorderType.NONE;
}
