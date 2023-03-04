package com.thomas.alib.excel.converter;

/**
 * 默认转换器，将任意属性对象以字符串形式输出
 */
public class DefaultConverter implements Converter {
    @Override
    public String convert(Object o) {
        return o == null ? "" : o.toString();
    }

    @Override
    public Object inverseConvert(Object o) {
        return o;
    }
}
