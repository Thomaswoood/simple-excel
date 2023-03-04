package com.thomas.alib.excel.utils;

import com.sun.istack.internal.Nullable;

import java.util.Collection;
import java.util.Map;

/**
 * 集合工具
 */
public class CollectionUtils {

    /**
     * 集合是否为空
     *
     * @param collection 集合对象
     * @return 是否为空
     */
    public static boolean isEmpty(@Nullable Collection<?> collection) {
        return collection == null || collection.isEmpty();
    }

    /**
     * Map是否为空
     *
     * @param map map对象
     * @return 是否为空
     */
    public static boolean isEmpty(@Nullable Map<?, ?> map) {
        return map == null || map.isEmpty();
    }

    /**
     * 集合是否不为空
     *
     * @param collection 集合对象
     * @return 是否不为空
     */
    public static boolean isNotEmpty(@Nullable Collection<?> collection) {
        return !isEmpty(collection);
    }

    /**
     * Map是否不为空
     *
     * @param map map对象
     * @return 是否不为空
     */
    public static boolean isNotEmpty(@Nullable Map<?, ?> map) {
        return !isEmpty(map);
    }
}
