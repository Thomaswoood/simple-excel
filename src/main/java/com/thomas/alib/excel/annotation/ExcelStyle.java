package com.thomas.alib.excel.annotation;

import com.thomas.alib.excel.constants.Constants;
import com.thomas.alib.excel.enums.SEBoolean;
import com.thomas.alib.excel.enums.SEBorderStyle;
import com.thomas.alib.excel.enums.SEHorAlignment;
import com.thomas.alib.excel.enums.SEVerAlignment;

import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;

/**
 * 样式注解
 */
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelStyle {

    /**
     * 字体名称，默认为空不设置
     */
    String fontName() default Constants.NOT_SET_S;

    /**
     * 文字大小，默认0不设置
     */
    double textSize() default Constants.NOT_SET_D;

    /**
     * 文字是否加粗，默认不设置
     */
    SEBoolean isBold() default SEBoolean.NOT_SET;

    /**
     * 文字是否斜体，默认不设置
     */
    SEBoolean isItalic() default SEBoolean.NOT_SET;

    /**
     * 文字是否划线，默认不设置
     */
    SEBoolean isStrikeout() default SEBoolean.NOT_SET;

    /**
     * 文字颜色，默认为空不设置，设置时使用#000000格式
     */
    String textColor() default Constants.NOT_SET_S;

    /**
     * 背景颜色，默认为空不设置，设置时使用#000000格式
     */
    String bgColor() default Constants.NOT_SET_S;

    /**
     * 水平对齐方式，默认不设置
     */
    SEHorAlignment horAlignment() default SEHorAlignment.NOT_SET;

    /**
     * 垂直对齐方式，默认不设置
     */
    SEVerAlignment verAlignment() default SEVerAlignment.NOT_SET;

    /**
     * 文字是否自动换行，默认不设置
     */
    SEBoolean isWrapText() default SEBoolean.NOT_SET;

    /**
     * 4边边框样式，默认不设置
     */
    SEBorderStyle borderStyle() default SEBorderStyle.NOT_SET;

    /**
     * 4边边框颜色，默认为空不设置，设置时使用#000000格式
     * 需要设置{@link #borderStyle()}后才会生效
     */
    String borderColor() default Constants.NOT_SET_S;
}
