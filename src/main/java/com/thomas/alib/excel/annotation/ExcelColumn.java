package com.thomas.alib.excel.annotation;


import com.thomas.alib.excel.converter.Converter;
import com.thomas.alib.excel.converter.DefaultConverter;
import com.thomas.alib.excel.loader.PictureLoader;
import com.thomas.alib.excel.loader.PictureLoaderDefault;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 表示列中某个属性为表格的一列的注解，可配置该列对应相关属性
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD})
public @interface ExcelColumn {

    /**
     * 列表头名称
     */
    String headerName() default "";

    /**
     * 列排序字段
     */
    int orderNum() default -1;

    /**
     * 列宽度
     */
    int columnWidth() default 4000;

    /**
     * 默认前缀
     */
    String prefix() default "";

    /**
     * 默认后缀
     */
    String suffix() default "";

    /**
     * 当不想实现Converter转换器时使用
     * 需要同时提供{@link #afterConvert()}
     * 本方法中提供转化之前的值，字符串类型
     * {@link #afterConvert()} 中提供转化之后的值，字符串类型
     *
     * @return 转化之前的值，字符串类型
     */
    String[] beforeConvert() default {};

    /**
     * 当不想实现Converter转换器时使用
     * 需要同时提供{@link #beforeConvert()}
     * 本方法中提供转化之后的值，字符串类型
     * {@link #beforeConvert()} 中提供转化之前的值，字符串类型
     *
     * @return 转化之后的值，字符串类型
     */
    String[] afterConvert() default {};

    /**
     * 指定转换器
     */
    Class<? extends Converter> converter() default DefaultConverter.class;

    /**
     * 是否按图片处理。注：按图片处理导出的字段值，需要提供图片加载器，默认是认为图片数据源是网络上完整的url来加载
     */
    boolean isPicture() default false;

    /**
     * 图片加载器，默认认为图片数据源是网络上完整的url来加载
     */
    Class<? extends PictureLoader<?>> pictureLoader() default PictureLoaderDefault.class;
}
