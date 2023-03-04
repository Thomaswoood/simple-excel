package com.thomas.alib.excel.converter;

import com.thomas.alib.excel.exception.AnalysisException;
import com.thomas.alib.excel.utils.StringUtils;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * 默认的时间类型对象的转换器
 */
public class DateConverter implements Converter {
    private static SimpleDateFormat SIMPLE_DATE_FORMAT_YMDHMS = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
    private static SimpleDateFormat SIMPLE_DATE_FORMAT_YMDHMS_A = new SimpleDateFormat("yyyy-M-d H:m:s");
    private static SimpleDateFormat SIMPLE_DATE_FORMAT_YMDHM_B = new SimpleDateFormat("yyyy-M-d H:m");
    private static SimpleDateFormat SIMPLE_DATE_FORMAT_YMDHMS_C = new SimpleDateFormat("yyyy/M/d H:m:s");
    private static SimpleDateFormat SIMPLE_DATE_FORMAT_YMDHM_D = new SimpleDateFormat("yyyy/M/d H:m");
    private static SimpleDateFormat SIMPLE_DATE_FORMAT_YMDHMS_E = new SimpleDateFormat("yyyy年M月d日 H:m:s");
    private static SimpleDateFormat SIMPLE_DATE_FORMAT_YMDHM_F = new SimpleDateFormat("yyyy年M月d日 H:m");
    private static SimpleDateFormat SIMPLE_DATE_FORMAT_YMDHMS_G = new SimpleDateFormat("yyyy.M.d H:m:s");
    private static SimpleDateFormat SIMPLE_DATE_FORMAT_YMDHM_H = new SimpleDateFormat("yyyy.M.d H:m");
    private static SimpleDateFormat SIMPLE_DATE_FORMAT_YMD_I = new SimpleDateFormat("yyyy-M-d");
    private static SimpleDateFormat SIMPLE_DATE_FORMAT_YMD_J = new SimpleDateFormat("yyyy/M/d");
    private static SimpleDateFormat SIMPLE_DATE_FORMAT_YMD_K = new SimpleDateFormat("yyyy年M月d日");
    private static SimpleDateFormat SIMPLE_DATE_FORMAT_YMD_L = new SimpleDateFormat("yyyy.M.d");

    @Override
    public String convert(Object o) {
        if (o == null) return "";
        return SIMPLE_DATE_FORMAT_YMDHMS.format(o);
    }

    /**
     * 目前仅支持如下格式日期的逆转换，如有需要再补充
     */
    @Override
    public Object inverseConvert(Object o) throws AnalysisException {
        if (o == null) return null;
        if (o instanceof Date) return o;
        String source = String.valueOf(o);
        if (StringUtils.isBlank(source)) return null;
        if (source.contains(":")) {
            try {
                return SIMPLE_DATE_FORMAT_YMDHMS_A.parse(source);
            } catch (ParseException ignore) {
            }
            try {
                return SIMPLE_DATE_FORMAT_YMDHM_B.parse(source);
            } catch (ParseException ignore) {
            }
            try {
                return SIMPLE_DATE_FORMAT_YMDHMS_C.parse(source);
            } catch (ParseException ignore) {
            }
            try {
                return SIMPLE_DATE_FORMAT_YMDHM_D.parse(source);
            } catch (ParseException ignore) {
            }
            try {
                return SIMPLE_DATE_FORMAT_YMDHMS_E.parse(source);
            } catch (ParseException ignore) {
            }
            try {
                return SIMPLE_DATE_FORMAT_YMDHM_F.parse(source);
            } catch (ParseException ignore) {
            }
            try {
                return SIMPLE_DATE_FORMAT_YMDHMS_G.parse(source);
            } catch (ParseException ignore) {
            }
            try {
                return SIMPLE_DATE_FORMAT_YMDHM_H.parse(source);
            } catch (ParseException ignore) {
            }
            throw new AnalysisException("时间格式错误");
        } else {
            try {
                return SIMPLE_DATE_FORMAT_YMD_I.parse(source);
            } catch (ParseException ignore) {
            }
            try {
                return SIMPLE_DATE_FORMAT_YMD_J.parse(source);
            } catch (ParseException ignore) {
            }
            try {
                return SIMPLE_DATE_FORMAT_YMD_K.parse(source);
            } catch (ParseException ignore) {
            }
            try {
                return SIMPLE_DATE_FORMAT_YMD_L.parse(source);
            } catch (ParseException ignore) {
            }
            throw new AnalysisException("日期格式错误");
        }
    }
}
