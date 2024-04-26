package com.thomas.alib.excel.exporter;

import com.thomas.alib.excel.annotation.ExcelStyle;
import com.thomas.alib.excel.constants.Constants;
import com.thomas.alib.excel.enums.SEBoolean;
import com.thomas.alib.excel.enums.SEBorderStyle;
import com.thomas.alib.excel.enums.SEHorAlignment;
import com.thomas.alib.excel.enums.SEVerAlignment;
import com.thomas.alib.excel.utils.StringUtils;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;

import java.awt.*;

/**
 * Excel导出时表格样式处理者
 */
class ExcelExportStyleProcessor {

    /**
     * 字体名称，默认为空不设置
     */
    private String fontName;

    /**
     * 文字大小，默认0不设置
     */
    private double textSize;

    /**
     * 文字是否加粗，默认不设置
     */
    private SEBoolean isBold;

    /**
     * 文字是否斜体，默认不设置
     */
    private SEBoolean isItalic;

    /**
     * 文字是否划线，默认不设置
     */
    private SEBoolean isStrikeout;

    /**
     * 文字颜色，默认为空不设置
     */
    private String textColor;

    /**
     * 背景颜色，默认为空不设置
     */
    private String bgColor;

    /**
     * 水平对齐方式，默认不设置
     */
    private SEHorAlignment horAlignment;

    /**
     * 垂直对齐方式，默认不设置
     */
    private SEVerAlignment verAlignment;

    /**
     * 文字是否自动换行，默认不设置
     */
    private SEBoolean isWrapText;

    /**
     * 4边边框样式，默认不设置
     */
    private SEBorderStyle borderStyle;

    /**
     * 4边边框颜色，默认为空不设置
     */
    private String borderColor;

    /**
     * 构造方法，将注解中设置的样式读取到本样式处理者中
     */
    ExcelExportStyleProcessor(ExcelStyle style) {
        if (style == null) {
            initDefault();
        } else {
            //取出注解设置的值
            fontName = style.fontName();//字体名称
            textSize = style.textSize();//文字大小
            isBold = style.isBold();//文字是否加粗
            isItalic = style.isItalic();//文字是否斜体
            isStrikeout = style.isStrikeout();//文字是否划线
            textColor = style.textColor();//文字颜色
            bgColor = style.bgColor();//背景颜色
            horAlignment = style.horAlignment();//水平对齐方式
            verAlignment = style.verAlignment();//垂直对齐方式
            isWrapText = style.isWrapText();//文字是否自动换行
            borderStyle = style.borderStyle();//4边边框样式
            borderColor = style.borderColor();//4边边框颜色
        }
    }

    /**
     * 构造方法，根据其他样式处理者，创建一个一样的新对象
     */
    ExcelExportStyleProcessor(ExcelExportStyleProcessor other) {
        if (other == null) {
            initDefault();
        } else {
            //取出注解设置的值
            this.fontName = other.fontName;//字体名称
            this.textSize = other.textSize;//文字大小
            this.isBold = other.isBold;//文字是否加粗
            this.isItalic = other.isItalic;//文字是否斜体
            this.isStrikeout = other.isStrikeout;//文字是否划线
            this.textColor = other.textColor;//文字颜色
            this.bgColor = other.bgColor;//背景颜色
            this.horAlignment = other.horAlignment;//水平对齐方式
            this.verAlignment = other.verAlignment;//垂直对齐方式
            this.isWrapText = other.isWrapText;//文字是否自动换行
            this.borderStyle = other.borderStyle;//4边边框样式
            this.borderColor = other.borderColor;//4边边框颜色
        }
    }

    /**
     * 按照默认值初始化
     */
    private void initDefault() {
        //设置默认值
        fontName = Constants.NOT_SET_S;//字体名称，默认为空不设置
        textSize = Constants.NOT_SET_D;//文字大小，默认0不设置
        isBold = SEBoolean.NOT_SET;//文字是否加粗，默认不设置
        isItalic = SEBoolean.NOT_SET;//文字是否斜体，默认不设置
        isStrikeout = SEBoolean.NOT_SET;//文字是否划线，默认不设置
        textColor = Constants.NOT_SET_S;//文字颜色，默认为空不设置
        bgColor = Constants.NOT_SET_S;//背景颜色，默认为空不设置
        horAlignment = SEHorAlignment.NOT_SET;//水平对齐方式，默认不设置
        verAlignment = SEVerAlignment.NOT_SET;//垂直对齐方式，默认不设置
        isWrapText = SEBoolean.NOT_SET;//文字是否自动换行，默认不设置
        borderStyle = SEBorderStyle.NOT_SET;//4边边框样式（仅支持基础样式），默认不设置
        borderColor = Constants.NOT_SET_S;//4边边框颜色，默认为空不设置
    }

    /**
     * 从注解中取出已设置的值覆盖到本样式处理者中
     *
     * @param style 样式注解对象
     * @return 链式调用，返回自身
     */
    ExcelExportStyleProcessor coverBySourceExceptNotSet(ExcelStyle style) {
        if (style != null) {
            //取出注解设置的值
            if (hadSet(style.fontName())) fontName = style.fontName();//字体名称
            if (hadSet(style.textSize())) textSize = style.textSize();//文字大小
            if (hadSet(style.isBold())) isBold = style.isBold();//文字是否加粗
            if (hadSet(style.isItalic())) isItalic = style.isItalic();//文字是否斜体
            if (hadSet(style.isStrikeout())) isStrikeout = style.isStrikeout();//文字是否划线
            if (hadSet(style.textColor())) textColor = style.textColor();//文字颜色
            if (hadSet(style.bgColor())) bgColor = style.bgColor();//背景颜色
            if (hadSet(style.horAlignment())) horAlignment = style.horAlignment();//水平对齐方式
            if (hadSet(style.verAlignment())) verAlignment = style.verAlignment();//垂直对齐方式
            if (hadSet(style.isWrapText())) isWrapText = style.isWrapText();//文字是否自动换行
            if (hadSet(style.borderStyle())) borderStyle = style.borderStyle();//4边边框样式
            if (hadSet(style.borderColor())) borderColor = style.borderColor();//4边边框颜色
        }
        return this;
    }

    /**
     * 从注解中取出已设置的值，覆盖到一个基于本样式处理者属性的一个新对象中
     *
     * @param style 样式注解对象
     * @return 基于原样式处理者属性，并已根据样式注解覆盖后，一个新对象
     */
    ExcelExportStyleProcessor coverBySourceExceptNotSetInNew(ExcelStyle style) {
        ExcelExportStyleProcessor result = new ExcelExportStyleProcessor(this);
        return result.coverBySourceExceptNotSet(style);
    }

    /**
     * 根据workbook创建Excel样式对象
     *
     * @param sxssfWorkbook excel导出构建对象
     * @return Excel样式对象
     */
    XSSFCellStyle createXSSFCellStyle(SXSSFWorkbook sxssfWorkbook) {
        //根据配置数据初始化样式
        XSSFCellStyle xssf_cell_style = (XSSFCellStyle) sxssfWorkbook.createCellStyle();
        //设置字体样式
        XSSFFont font = null;
        //设置字体名称
        if (hadSet(fontName)) {
            font = (XSSFFont) sxssfWorkbook.createFont();
            font.setFontName(fontName);
        }
        //设置文字大小
        if (hadSet(textSize)) {
            if (font == null) font = (XSSFFont) sxssfWorkbook.createFont();
            font.setFontHeight(textSize);
        }
        //设置是否加粗
        if (hadSet(isBold)) {
            if (font == null) font = (XSSFFont) sxssfWorkbook.createFont();
            font.setBold(isBold.getValue());
        }
        //设置是否斜体
        if (hadSet(isItalic)) {
            if (font == null) font = (XSSFFont) sxssfWorkbook.createFont();
            font.setItalic(isItalic.getValue());
        }
        //设置是否划线
        if (hadSet(isStrikeout)) {
            if (font == null) font = (XSSFFont) sxssfWorkbook.createFont();
            font.setStrikeout(isStrikeout.getValue());
        }
        //设置文字颜色
        if (hadSet(textColor)) {
            if (font == null) font = (XSSFFont) sxssfWorkbook.createFont();
            font.setColor(new XSSFColor(Color.decode(textColor), new DefaultIndexedColorMap()));
        }
        //将字体设置进样式对象
        if (font != null) xssf_cell_style.setFont(font);
        //设置背景色
        if (hadSet(bgColor)) {
            //xssf_cell_style.setFillBackgroundColor(new XSSFColor(Color.decode(bgColor), new DefaultIndexedColorMap()));
            xssf_cell_style.setFillForegroundColor(new XSSFColor(Color.decode(bgColor), new DefaultIndexedColorMap()));
            xssf_cell_style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }
        //设置水平对齐方式
        if (hadSet(horAlignment)) {
            switch (horAlignment) {
                case GENERAL://自动对齐
                    xssf_cell_style.setAlignment(HorizontalAlignment.GENERAL);
                    break;
                case LEFT://左对齐
                    xssf_cell_style.setAlignment(HorizontalAlignment.LEFT);
                    break;
                case CENTER://居中
                    xssf_cell_style.setAlignment(HorizontalAlignment.CENTER);
                    break;
                case RIGHT://右对齐
                    xssf_cell_style.setAlignment(HorizontalAlignment.RIGHT);
                    break;
                case FILL://分散对齐
                    xssf_cell_style.setAlignment(HorizontalAlignment.FILL);
                    break;
                case JUSTIFY://两端对齐
                    xssf_cell_style.setAlignment(HorizontalAlignment.JUSTIFY);
                    break;
            }
        }
        //设置垂直对齐方式
        if (hadSet(verAlignment)) {
            switch (verAlignment) {
                case TOP://顶端对齐
                    xssf_cell_style.setVerticalAlignment(VerticalAlignment.TOP);
                    break;
                case CENTER://垂直居中
                    xssf_cell_style.setVerticalAlignment(VerticalAlignment.CENTER);
                    break;
                case BOTTOM://底端对齐
                    xssf_cell_style.setVerticalAlignment(VerticalAlignment.BOTTOM);
                    break;
                case JUSTIFY://两端对齐
                    xssf_cell_style.setVerticalAlignment(VerticalAlignment.JUSTIFY);
                    break;
            }
        }
        //设置是否自动换行
        if (hadSet(isWrapText)) {
            xssf_cell_style.setWrapText(isWrapText.getValue());
        }
        //设置边框
        if (hadSet(borderStyle)) {
            BorderStyle border_style = BorderStyle.NONE;
            switch (borderStyle) {
                case THIN://细线
                    border_style = BorderStyle.THIN;
                    break;
                case MEDIUM://粗一点的线
                    border_style = BorderStyle.MEDIUM;
                    break;
                case DASHED://细虚线
                    border_style = BorderStyle.DASHED;
                    break;
                case THICK://更粗的线
                    border_style = BorderStyle.THICK;
                    break;
                case DOUBLE://双划线
                    border_style = BorderStyle.DOUBLE;
                    break;
                case MEDIUM_DASHED://粗一点的虚线
                    border_style = BorderStyle.MEDIUM_DASHED;
                    break;
                case DASH_DOT://一长一短的虚线
                    border_style = BorderStyle.DASH_DOT;
                    break;
                case MEDIUM_DASH_DOT://粗一点的一长一短的虚线
                    border_style = BorderStyle.MEDIUM_DASH_DOT;
                    break;
                case DASH_DOT_DOT://一长两短的虚线
                    border_style = BorderStyle.DASH_DOT_DOT;
                    break;
                case MEDIUM_DASH_DOT_DOT://粗一点的一长两短的虚线
                    border_style = BorderStyle.MEDIUM_DASH_DOT_DOT;
                    break;
                case SLANTED_DASH_DOT://斜画的一长一短的虚线
                    border_style = BorderStyle.SLANTED_DASH_DOT;
                    break;
            }
            if (border_style != BorderStyle.NONE) {
                xssf_cell_style.setBorderLeft(border_style);
                xssf_cell_style.setBorderTop(border_style);
                xssf_cell_style.setBorderRight(border_style);
                xssf_cell_style.setBorderBottom(border_style);
                //设置边框颜色
                if (hadSet(borderColor)) {
                    XSSFColor xssfColor = new XSSFColor(Color.decode(borderColor), new DefaultIndexedColorMap());
                    xssf_cell_style.setLeftBorderColor(xssfColor);
                    xssf_cell_style.setTopBorderColor(xssfColor);
                    xssf_cell_style.setRightBorderColor(xssfColor);
                    xssf_cell_style.setBottomBorderColor(xssfColor);
                }
            }
        }
        return xssf_cell_style;
    }

    /**
     * 判断是否是已设置属性值
     *
     * @param value 带判断的值
     * @return 是否是已设置属性值
     */
    static boolean hadSet(String value) {
        return !StringUtils.equals(value, Constants.NOT_SET_S);
    }

    /**
     * 判断是否是已设置属性值
     *
     * @param value 带判断的值
     * @return 是否是已设置属性值
     */
    static boolean hadSet(double value) {
        return value != Constants.NOT_SET_D;
    }

    /**
     * 判断是否是已设置属性值
     *
     * @param value 带判断的值
     * @return 是否是已设置属性值
     */
    static boolean hadSet(SEBoolean value) {
        return value != SEBoolean.NOT_SET;
    }

    /**
     * 判断是否是已设置属性值
     *
     * @param value 带判断的值
     * @return 是否是已设置属性值
     */
    static boolean hadSet(SEBorderStyle value) {
        return value != SEBorderStyle.NOT_SET;
    }

    /**
     * 判断是否是已设置属性值
     *
     * @param value 带判断的值
     * @return 是否是已设置属性值
     */
    static boolean hadSet(SEHorAlignment value) {
        return value != SEHorAlignment.NOT_SET;
    }

    /**
     * 判断是否是已设置属性值
     *
     * @param value 带判断的值
     * @return 是否是已设置属性值
     */
    static boolean hadSet(SEVerAlignment value) {
        return value != SEVerAlignment.NOT_SET;
    }

    /**
     * 判断是否是已设置样式属性值
     *
     * @param style 带判断的样式注解对象
     * @return 是否是已设置样式属性值
     */
    static boolean hadSet(ExcelStyle style) {
        if (style == null) return false;
        return hadSet(style.fontName())
                || hadSet(style.textSize())
                || hadSet(style.isBold())
                || hadSet(style.isItalic())
                || hadSet(style.isStrikeout())
                || hadSet(style.textColor())
                || hadSet(style.bgColor())
                || hadSet(style.horAlignment())
                || hadSet(style.verAlignment())
                || hadSet(style.isWrapText())
                || hadSet(style.borderStyle())
                || hadSet(style.borderColor());
    }

    /**
     * 读取样式注解信息
     *
     * @param style 样式注解对象
     */
    static ExcelExportStyleProcessor read(ExcelStyle style) {
        return new ExcelExportStyleProcessor(style);
    }
}
