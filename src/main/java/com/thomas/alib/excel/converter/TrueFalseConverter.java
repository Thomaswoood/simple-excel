package com.thomas.alib.excel.converter;

import com.thomas.alib.excel.exception.AnalysisException;
import com.thomas.alib.excel.utils.StringUtils;

import java.util.Date;

/**
 * 是否转换器 0否1是
 */
public enum TrueFalseConverter implements  Converter {
    FALSE(0, "否"),
    TRUE(1, "是");

    private final int code;
    private final String name;

    TrueFalseConverter(int code, String name) {
        this.code = code;
        this.name = name;
    }

    public int getCode() {
        return code;
    }

    public String getName() {
        return name;
    }

    public static TrueFalseConverter getFromCode(Integer value) {
        if (value == null) return FALSE;
        for (TrueFalseConverter item : TrueFalseConverter.values()) {
            if (item.code == value) {
                return item;
            }
        }
        return FALSE;
    }

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
    @Override
    public String convert(Object o) {
        if (o == null) return "";
        for (TrueFalseConverter item : values()) {
            if (StringUtils.equals(item.code + "", String.valueOf(o))) {
                return item.name;
            }
        }
        return o.toString();
    }

    /**
     * 字段逆转换器
     * 用于：
     * 表格导入时使用，表格内存储的都是个时候的易显易读的数据，
     * 而实际数据模型中是code或者是data的值，
     * 等等情况
     * 此方法基本上是convert的逆逻辑，但是也不全是，在jfinal-ext依赖读取表格时，
     * 比如：
     * 可以将date型数据直接读取为{@link Date}对象；
     * 可以将数字型（包括整型）数据直接读取成double型数据；
     * 目前考虑到的情况仅这两种，后续有需要会再完善
     *
     * @param o 转换后的值
     * @return 转换前的值
     * @throws AnalysisException excel解析异常
     */
    @Override
    public Object inverseConvert(Object o) throws AnalysisException {
        if (o == null) return null;
        for (TrueFalseConverter item : values()) {
            if (item.name.equals(String.valueOf(o))) {
                return item.code;
            }
        }
        return null;
    }
}