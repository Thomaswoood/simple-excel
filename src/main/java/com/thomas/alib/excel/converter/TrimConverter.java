package com.thomas.alib.excel.converter;

/**
 * trim转换器，只能给String类型属性使用，会将字符换进行trim转换
 */
public class TrimConverter implements Converter {
    @Override
    public String convert(Object o) {
        return o == null ? "" : o.toString().trim();
    }

    @Override
    public Object inverseConvert(Object o) {
        return o == null ? "" : o.toString().trim();
    }
}
