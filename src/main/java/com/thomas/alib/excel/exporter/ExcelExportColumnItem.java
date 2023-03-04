package com.thomas.alib.excel.exporter;

import com.thomas.alib.excel.base.ExcelColumnRender;
import com.thomas.alib.excel.converter.DefaultConverter;

import java.lang.reflect.Field;

/**
 * Excel导出对应每个列(Column)的解析工具
 */
class ExcelExportColumnItem extends ExcelColumnRender implements Comparable<ExcelExportColumnItem> {

    ExcelExportColumnItem(Field field) {
        super(field);
    }

    /**
     * 从数据源中取出图片作为byte数组
     *
     * @param source 数据源
     * @return 图片byte数组
     */
    byte[] getColumnPictureBytesFromSource(Object source) {
        try {
            if (converter == null || converter instanceof DefaultConverter) {
                //没配置转化器或是默认转化器，认为外部没有主动配置转化器，不需要转化，直接取出值传递给
                return pictureLoader.insideLoad(columnField.get(source));
            } else {
                return pictureLoader.insideLoad(getColumnValueFromSource(source));
            }
        } catch (Throwable e) {
            return null;
        }
    }

    /**
     * 从数据源中取出本列的值，方法中自动处理convert转化
     *
     * @param source 数据源
     * @return 本列的值
     */
    String getColumnValueFromSource(Object source) {
        try {
            //将数据源的成员值转化后返回
            String convertCenter = converter.convert(columnField.get(source));
            if (convertCenter == null) convertCenter = "";
            return prefix + convertCenter + suffix;
        } catch (Throwable e) {
            //成员值无效
            return "读取失败";
        }
    }

    /**
     * 比较
     */
    @Override
    public int compareTo(ExcelExportColumnItem o) {
        return this.orderNum - o.orderNum;
    }
}
