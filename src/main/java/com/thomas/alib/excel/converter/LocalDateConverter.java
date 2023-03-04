package com.thomas.alib.excel.converter;

import com.thomas.alib.excel.exception.AnalysisException;
import com.thomas.alib.excel.utils.StringUtils;

import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.time.temporal.TemporalAccessor;
import java.util.Date;

/**
 * 默认的时间类型对象的转换器
 */
public class LocalDateConverter implements Converter {
    private static DateTimeFormatter DATETIME_FORMATTER_YMD = DateTimeFormatter.ofPattern("yyyy-MM-dd");
    private static DateTimeFormatter DATETIME_FORMATTER_YMD_I = DateTimeFormatter.ofPattern("yyyy-M-d");
    private static DateTimeFormatter DATETIME_FORMATTER_YMD_J = DateTimeFormatter.ofPattern("yyyy/M/d");
    private static DateTimeFormatter DATETIME_FORMATTER_YMD_K = DateTimeFormatter.ofPattern("yyyy年M月d日");
    private static DateTimeFormatter DATETIME_FORMATTER_YMD_L = DateTimeFormatter.ofPattern("yyyy.M.d");

    @Override
    public String convert(Object o) {
        if (o == null) return "";
        return DATETIME_FORMATTER_YMD.format((TemporalAccessor) o);
    }

    /**
     * 目前仅支持如下格式日期的逆转换，如有需要再补充
     */
    @Override
    public Object inverseConvert(Object o) throws AnalysisException {
        if (o == null) return null;
        if (o instanceof Date) {
            return ((Date) o).toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
        }
        String source = String.valueOf(o);
        if (StringUtils.isBlank(source)) return null;
        try {
            return LocalDate.parse(source, DATETIME_FORMATTER_YMD_I);
        } catch (Throwable ignore) {
        }
        try {
            return LocalDate.parse(source, DATETIME_FORMATTER_YMD_J);
        } catch (Throwable ignore) {
        }
        try {
            return LocalDate.parse(source, DATETIME_FORMATTER_YMD_K);
        } catch (Throwable ignore) {
        }
        try {
            return LocalDate.parse(source, DATETIME_FORMATTER_YMD_L);
        } catch (Throwable ignore) {
        }
        throw new AnalysisException("日期格式错误");
    }
}
