package com.thomas.alib.excel.utils;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * 反射工具
 */
public class ReflectUtil {

    /**
     * 获取指定类的全部成员数组，包括父类
     *
     * @param clazz 指定的类
     * @return 全部成员数组，包括父类
     */
    public static List<Field> getAccessibleFieldIncludeSuper(Class<?> clazz) {
        List<Field> result = new ArrayList<>();
        try {
            result.addAll(Arrays.asList(clazz.getDeclaredFields()));
            Class<?> super_clazz = clazz.getSuperclass();
            while (super_clazz != null) {
                result.addAll(Arrays.asList(super_clazz.getDeclaredFields()));
                super_clazz = super_clazz.getSuperclass();
            }
        } catch (Throwable ignored) {
        }
        return result;
    }
}
