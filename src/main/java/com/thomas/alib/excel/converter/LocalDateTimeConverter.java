package com.thomas.alib.excel.converter;

import com.thomas.alib.excel.exception.AnalysisException;
import com.thomas.alib.excel.utils.StringUtils;

import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.time.temporal.TemporalAccessor;
import java.util.Date;

/**
 * 默认的时间类型对象的转换器
 */
public class LocalDateTimeConverter implements Converter {
    private static DateTimeFormatter DATETIME_FORMATTER_YMDHMS = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
    private static DateTimeFormatter DATETIME_FORMATTER_YMDHMS_A = DateTimeFormatter.ofPattern("yyyy-M-d H:m:s");
    private static DateTimeFormatter DATETIME_FORMATTER_YMDHM_B = DateTimeFormatter.ofPattern("yyyy-M-d H:m");
    private static DateTimeFormatter DATETIME_FORMATTER_YMDHMS_C = DateTimeFormatter.ofPattern("yyyy/M/d H:m:s");
    private static DateTimeFormatter DATETIME_FORMATTER_YMDHM_D = DateTimeFormatter.ofPattern("yyyy/M/d H:m");
    private static DateTimeFormatter DATETIME_FORMATTER_YMDHMS_E = DateTimeFormatter.ofPattern("yyyy年M月d日 H:m:s");
    private static DateTimeFormatter DATETIME_FORMATTER_YMDHM_F = DateTimeFormatter.ofPattern("yyyy年M月d日 H:m");
    private static DateTimeFormatter DATETIME_FORMATTER_YMDHMS_G = DateTimeFormatter.ofPattern("yyyy.M.d H:m:s");
    private static DateTimeFormatter DATETIME_FORMATTER_YMDHM_H = DateTimeFormatter.ofPattern("yyyy.M.d H:m");

    @Override
    public String convert(Object o) {
        if (o == null) return "";
        return DATETIME_FORMATTER_YMDHMS.format((TemporalAccessor) o);
    }

    /**
     * 目前仅支持如下格式日期的逆转换，如有需要再补充
     */
    @Override
    public Object inverseConvert(Object o) throws AnalysisException {
        if (o == null) return null;
        if (o instanceof Date) {
            return ((Date) o).toInstant().atZone(ZoneId.systemDefault()).toLocalDateTime();
        }
        String source = String.valueOf(o);
        if (StringUtils.isBlank(source)) return null;
        if (source.contains(":")) {
            try {
                return LocalDateTime.parse(source, DATETIME_FORMATTER_YMDHMS_A);
            } catch (Throwable ignore) {
            }
            try {
                return LocalDateTime.parse(source, DATETIME_FORMATTER_YMDHM_B);
            } catch (Throwable ignore) {
            }
            try {
                return LocalDateTime.parse(source, DATETIME_FORMATTER_YMDHMS_C);
            } catch (Throwable ignore) {
            }
            try {
                return LocalDateTime.parse(source, DATETIME_FORMATTER_YMDHM_D);
            } catch (Throwable ignore) {
            }
            try {
                return LocalDateTime.parse(source, DATETIME_FORMATTER_YMDHMS_E);
            } catch (Throwable ignore) {
            }
            try {
                return LocalDateTime.parse(source, DATETIME_FORMATTER_YMDHM_F);
            } catch (Throwable ignore) {
            }
            try {
                return LocalDateTime.parse(source, DATETIME_FORMATTER_YMDHMS_G);
            } catch (Throwable ignore) {
            }
            try {
                return LocalDateTime.parse(source, DATETIME_FORMATTER_YMDHM_H);
            } catch (Throwable ignore) {
            }
        } else {
            try {
                return LocalDateTime.parse(source + " 00:00", DATETIME_FORMATTER_YMDHM_B);
            } catch (Throwable ignore) {
            }
            try {
                return LocalDateTime.parse(source + " 00:00", DATETIME_FORMATTER_YMDHM_D);
            } catch (Throwable ignore) {
            }
            try {
                return LocalDateTime.parse(source + " 00:00", DATETIME_FORMATTER_YMDHM_F);
            } catch (Throwable ignore) {
            }
            try {
                return LocalDateTime.parse(source + " 00:00", DATETIME_FORMATTER_YMDHM_H);
            } catch (Throwable ignore) {
            }
        }
        throw new AnalysisException("时间格式错误");
    }
}
