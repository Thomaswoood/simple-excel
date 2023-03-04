package com.thomas.alib.excel.exporter;

import com.thomas.alib.excel.annotation.ExcelExportDataStyle;
import com.thomas.alib.excel.annotation.ExcelExportHeadStyle;
import com.thomas.alib.excel.enums.BorderType;
import com.thomas.alib.excel.enums.HorAlignment;
import com.thomas.alib.excel.enums.VerAlignment;
import com.thomas.alib.excel.utils.StringUtils;
import org.apache.poi.ss.usermodel.*;

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
    private short textSize;

    /**
     * 文字是否加粗，默认不加粗
     */
    private boolean isBold;

    /**
     * 文字是否斜体，默认不斜体
     */
    private boolean isItalic;

    /**
     * 文字是否划线，默认不划线
     */
    private boolean isStrikeout;

    /**
     * 文字颜色，默认-1不设置
     */
    private short textColor;

    /**
     * 水平对齐方式，默认自动对齐
     */
    private HorAlignment horAlignment;

    /**
     * 垂直对齐方式，默认垂直居中
     */
    private VerAlignment verAlignment;

    /**
     * 文字是否自动换行，默认否
     */
    private boolean isWrapText;

    /**
     * 4边边框样式，默认无边框
     */
    private BorderType fourSideBorder;
    /**
     * 根据配置数据创建出的复用样式
     */
    private CellStyle baseStyle;

    /**
     * 默认构造方法，设置默认值
     */
    ExcelExportStyleProcessor() {
        //设置默认值
        fontName = null;//表头行字体名称，默认为空不设置
        textSize = 0;//表头行文字大小，默认0不设置
        isBold = false;//表头行文字是否加粗，默认不加粗
        isItalic = false;//表头行文字是否斜体，默认不斜体
        isStrikeout = false;//表头行文字是否划线，默认不划线
        textColor = -1;//表头行文字颜色，默认-1不设置
        horAlignment = HorAlignment.GENERAL;//表头行水平对齐方式，默认自动对齐
        verAlignment = VerAlignment.CENTER;//表头行垂直对齐方式，默认垂直居中
        isWrapText = false;//表头行文字是否自动换行，默认否
        fourSideBorder = BorderType.NONE;//表头行4边边框样式（仅支持基础样式），默认0无边框
    }

    /**
     * 根据workbook初始化样式
     *
     * @param workbook excel导出构建对象
     */
    void initWith(Workbook workbook) {
        //根据配置数据初始化样式
        baseStyle = workbook.createCellStyle();
        //设置水平对齐方式
        switch (horAlignment) {
            case GENERAL://自动对齐
                baseStyle.setAlignment(HorizontalAlignment.GENERAL);
                break;
            case LEFT://左对齐
                baseStyle.setAlignment(HorizontalAlignment.LEFT);
                break;
            case CENTER://居中
                baseStyle.setAlignment(HorizontalAlignment.CENTER);
                break;
            case RIGHT://右对齐
                baseStyle.setAlignment(HorizontalAlignment.RIGHT);
                break;
            case FILL://分散对齐
                baseStyle.setAlignment(HorizontalAlignment.FILL);
                break;
            case JUSTIFY://两端对齐
                baseStyle.setAlignment(HorizontalAlignment.JUSTIFY);
                break;
        }
        //设置垂直对齐方式
        switch (verAlignment) {
            case TOP://顶端对齐
                baseStyle.setVerticalAlignment(VerticalAlignment.TOP);
                break;
            case CENTER://垂直居中
                baseStyle.setVerticalAlignment(VerticalAlignment.CENTER);
                break;
            case BOTTOM://底端对齐
                baseStyle.setVerticalAlignment(VerticalAlignment.BOTTOM);
                break;
            case JUSTIFY://两端对齐
                baseStyle.setVerticalAlignment(VerticalAlignment.JUSTIFY);
                break;
        }
        //设置是否自动换行
        baseStyle.setWrapText(isWrapText);
        //设置边框
        if (fourSideBorder != BorderType.NONE) {
            BorderStyle border_flag = BorderStyle.NONE;
            switch (fourSideBorder) {
                case THIN://细线
                    border_flag = BorderStyle.THIN;
                    break;
                case MEDIUM://粗一点的线
                    border_flag = BorderStyle.MEDIUM;
                    break;
                case DASHED://细虚线
                    border_flag = BorderStyle.DASHED;
                    break;
                case THICK://更粗的线
                    border_flag = BorderStyle.THICK;
                    break;
                case DOUBLE://双划线
                    border_flag = BorderStyle.DOUBLE;
                    break;
                case MEDIUM_DASHED://粗一点的虚线
                    border_flag = BorderStyle.MEDIUM_DASHED;
                    break;
                case DASH_DOT://一长一短的虚线
                    border_flag = BorderStyle.DASH_DOT;
                    break;
                case MEDIUM_DASH_DOT://粗一点的一长一短的虚线
                    border_flag = BorderStyle.MEDIUM_DASH_DOT;
                    break;
                case DASH_DOT_DOT://一长两短的虚线
                    border_flag = BorderStyle.DASH_DOT_DOT;
                    break;
                case MEDIUM_DASH_DOT_DOT://粗一点的一长两短的虚线
                    border_flag = BorderStyle.MEDIUM_DASH_DOT_DOT;
                    break;
                case SLANTED_DASH_DOT://斜画的一长一短的虚线
                    border_flag = BorderStyle.SLANTED_DASH_DOT;
                    break;
            }
            if (border_flag != BorderStyle.NONE) {
                baseStyle.setBorderLeft(border_flag);
                baseStyle.setBorderTop(border_flag);
                baseStyle.setBorderRight(border_flag);
                baseStyle.setBorderBottom(border_flag);
            }
        }
        //设置字体样式
        Font font = workbook.createFont();
        //设置字体名称
        if (StringUtils.isNotBlank(fontName)) {
            font.setFontName(fontName);
        }
        //设置文字大小
        if (textSize > 0) {
            font.setFontHeightInPoints(textSize);
        }
        //设置是否加粗
        font.setBold(isBold);
        //设置是否斜体
        font.setItalic(isItalic);
        //设置是否划线
        font.setStrikeout(isStrikeout);
        //设置文字颜色
        if (textColor >= 0) {
            font.setColor(textColor);
        }
        //将字体设置进样式对象
        baseStyle.setFont(font);
    }

    /**
     * 读取表头样式注解信息
     *
     * @param sheet_data_clazz 数据源泛型class
     */
    static ExcelExportStyleProcessor withHead(Workbook workbook, Class<?> sheet_data_clazz) {
        //处理样式类返回对象
        ExcelExportStyleProcessor processor = new ExcelExportStyleProcessor();
        //取得表格sheet相关注解
        ExcelExportHeadStyle style = sheet_data_clazz.getAnnotation(ExcelExportHeadStyle.class);
        if (style != null) {
            //已设置sheet注解，取出设置的值
            processor.fontName = style.fontName();//表头行字体名称
            processor.textSize = style.textSize();//表头行文字大小
            processor.isBold = style.isBold();//表头行文字是否加粗
            processor.isItalic = style.isItalic();//表头行文字是否斜体
            processor.isStrikeout = style.isStrikeout();//表头行文字是否划线
            processor.textColor = style.textColor();//表头行文字颜色
            processor.horAlignment = style.horAlignment();//表头行水平对齐方式
            processor.verAlignment = style.verAlignment();//表头行垂直对齐方式
            processor.isWrapText = style.isWrapText();//表头行文字是否自动换行
            processor.fourSideBorder = style.fourSideBorder();//表头行4边边框样式
        }
        processor.initWith(workbook);
        return processor;
    }

    /**
     * 读取表头样式注解信息
     *
     * @param sheet_data_clazz 数据源泛型class
     */
    static ExcelExportStyleProcessor withData(Workbook workbook, Class<?> sheet_data_clazz) {
        //处理样式类返回对象
        ExcelExportStyleProcessor processor = new ExcelExportStyleProcessor();
        //取得表格sheet相关注解
        ExcelExportDataStyle style = sheet_data_clazz.getAnnotation(ExcelExportDataStyle.class);
        if (style != null) {
            //已设置sheet注解，取出设置的值
            processor.fontName = style.fontName();//数据行字体名称
            processor.textSize = style.textSize();//数据行文字大小
            processor.isBold = style.isBold();//数据行文字是否加粗
            processor.isItalic = style.isItalic();//数据行文字是否斜体
            processor.isStrikeout = style.isStrikeout();//数据行文字是否划线
            processor.textColor = style.textColor();//数据行文字颜色
            processor.horAlignment = style.horAlignment();//数据行水平对齐方式
            processor.verAlignment = style.verAlignment();//数据行垂直对齐方式
            processor.isWrapText = style.isWrapText();//数据行文字是否自动换行
            processor.fourSideBorder = style.fourSideBorder();//数据行4边边框样式
        }
        processor.initWith(workbook);
        return processor;
    }

    /**
     * 给单元格设置样式
     *
     * @param cell 待设置的单元格对象
     */
    void setStyle(Cell cell) {
        cell.setCellStyle(baseStyle);
    }
}
