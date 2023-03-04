package com.thomas.alib.excel.converter;

import com.thomas.alib.excel.exception.AnalysisException;

/**
 * 字段转换器
 * 用于:
 * 表格数据模型中存储的是code值，而实际显示需要显示name的;
 * 表格数据模型中存储的是date，而实际需要格式化展示的；
 * 等等情况
 */
public interface Converter {

    /**
     * 字段转换器
     * 用于:
     * 表格数据模型中存储的是code值，而实际显示需要显示name的;
     * 表格数据模型中存储的是date，而实际需要格式化展示的；
     * 等等情况
     *
     * @param o 转换前的值
     * @return 转换后的值
     */
    String convert(Object o);

    /**
     * 字段逆转换器
     * 用于：
     * 表格导入时使用，表格内存储的都是个时候的易显易读的数据，
     * 而实际数据模型中是code或者是data的值，
     * 等等情况
     * 此方法基本上是convert的逆逻辑，但是也不全是，在jfinal-ext依赖读取表格时，
     * 比如：
     * 可以将date型数据直接读取为{@link java.util.Date}对象；
     * 可以将数字型（包括整型）数据直接读取成double型数据；
     * 目前考虑到的情况仅这两种，后续有需要会再完善
     *
     * @param o 转换后的值
     * @return 转换前的值
     * @throws AnalysisException excel解析异常
     */
    Object inverseConvert(Object o) throws AnalysisException;
}
