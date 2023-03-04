package com.thomas.alib.excel.converter;

import com.thomas.alib.excel.annotation.ExcelColumn;

/**
 * 默认的简易转化器，配合
 * {@link ExcelColumn#beforeConvert()}以及{@link ExcelColumn#afterConvert()}使用
 * 将属性按照配置的的转换方式输出
 */
public class SimpleConvert implements Converter {
    private final String[] before;
    private final String[] after;

    public SimpleConvert(String[] before, String[] after) {
        this.before = before;
        this.after = after;
    }

    private String simpleConvert(String[] original, String[] converted, Object o) {
        try {
            String value = String.valueOf(o);
            for (int i = 0; i < original.length; i++) {
                if (value.equals(original[i])) return converted[i];
            }
        } catch (Throwable ignore) {
        }
        return null;
    }

    @Override
    public String convert(Object o) {
        String value = simpleConvert(before, after, o);
        return value == null ? (o == null ? "" : o.toString()) : value;
    }

    @Override
    public Object inverseConvert(Object o) {
        return simpleConvert(after, before, o);
    }
}
