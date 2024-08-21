package com.thomas.alib.excel.exporter;

import com.thomas.alib.excel.base.ExcelColumnRender;
import com.thomas.alib.excel.loader.PictureLoader;
import com.thomas.alib.excel.loader.PictureLoaderDefault;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.lang.reflect.Field;

/**
 * Excel导出对应每个列(Column)的解析工具
 */
class ExcelExportColumnItem extends ExcelColumnRender implements Comparable<ExcelExportColumnItem> {
    private static final Logger logger = LoggerFactory.getLogger(ExcelExportColumnItem.class);
    /**
     * 图片加载器
     */
    protected PictureLoader<?> pictureLoader;
    /**
     * 列定制表头行样式，无特殊定制时为空
     */
    private XSSFCellStyle columnHeadStyle;
    /**
     * 列定制数据行样式，无特殊定制时为空
     */
    private XSSFCellStyle columnDataStyle;

    ExcelExportColumnItem(Field field, SXSSFWorkbook sxssf_workbook, ExcelExportStyleProcessor head_style_processor, ExcelExportStyleProcessor data_style_processor) {
        super(field);
        if (isValid) {
            if (ExcelExportStyleProcessor.hadSet(excelColumn.columnStyle())) {
                //数据样式处理
                columnDataStyle = data_style_processor.coverBySourceExceptNotSetInNew(excelColumn.columnStyle()).createXSSFCellStyle(sxssf_workbook);
                //判断列样式是否影响表头
                if (excelColumn.columnStyleInHead()) {
                    //表头样式处理
                    columnHeadStyle = head_style_processor.coverBySourceExceptNotSetInNew(excelColumn.columnStyle()).createXSSFCellStyle(sxssf_workbook);
                }
            }
            //判断是否按图片处理，按图片处理导出时，需要初始化图片加载器
            if (isPicture) {
                if (excelColumn.pictureLoader() == PictureLoaderDefault.class) {
                    //使用默认图片加载器
                    pictureLoader = new PictureLoaderDefault();
                } else {
                    //使用自定义图片加载器
                    try {
                        if (excelColumn.pictureLoader().isEnum()) {//枚举类型不可创建，需要取出一个
                            pictureLoader = excelColumn.pictureLoader().getEnumConstants()[0];
                        } else {//其他类型则自动创建一个
                            pictureLoader = excelColumn.pictureLoader().newInstance();
                        }
                    } catch (Throwable e) {//发生未知错误，使用默认图片加载器
                        logger.error("表格\"" + headName + "\"列-在创建您配置的pictureLoader时发生错误，将使用默认方式继续处理，具体错误为:", e);
                        pictureLoader = new PictureLoaderDefault();
                    }
                }
            }
        }
    }

    /**
     * 从数据源中取出图片作为byte数组
     *
     * @param source    数据源
     * @param row_index 数据行序号，打印日志使用
     * @return 图片byte数组
     */
    byte[] getColumnPictureBytesFromSource(Object source, int row_index) {
        try {
            return pictureLoader.insideLoad(getColumnValueFromSource(source, row_index));
        } catch (Throwable e) {
            logger.error("表格\"" + headName + "\"列-第" + row_index + "行-加载图片时发生错误:", e);
            return null;
        }
    }

    /**
     * 从数据源中取出本列的值，方法中自动处理convert转化
     *
     * @param source    数据源
     * @param row_index 数据行序号，打印日志使用
     * @return 本列的值
     */
    String getColumnValueFromSource(Object source, int row_index) {
        try {
            //将数据源的成员值转化后返回
            String convertCenter = converter.convert(columnField.get(source));
            if (convertCenter == null) convertCenter = "";
            return prefix + convertCenter + suffix;
        } catch (Throwable e) {
            //成员值无效
            logger.error("表格\"" + headName + "\"列-第" + row_index + "行-读取属性值时发生错误:", e);
            return "读取失败";
        }
    }

    public XSSFCellStyle getColumnHeadStyle() {
        return columnHeadStyle;
    }

    public XSSFCellStyle getColumnDataStyle() {
        return columnDataStyle;
    }

    /**
     * 比较
     */
    @Override
    public int compareTo(ExcelExportColumnItem o) {
        return this.orderNum - o.orderNum;
    }

    @Override
    public void close() {
        super.close();
        pictureLoader = null;
        columnHeadStyle = null;
        columnDataStyle = null;
    }
}
