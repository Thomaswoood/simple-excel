package com.thomas.alib.excel.loader;

import java.io.IOException;

/**
 * 图片加载器，将图片加载为byte数组
 *
 * @param <T> 图片数据源泛型
 */
public interface PictureLoader<T> {

    /**
     * 工具内部调用加载用方法，内部处理类型强转，可能抛出异常
     *
     * @param o 图片数据源
     * @return 图片byte数组
     * @throws Throwable 图片读取错误
     */
    default byte[] insideLoad(Object o) throws Throwable {
        return loadPicture((T) o);
    }

    /**
     * 加载图片为byte数组
     *
     * @param e 图片数据源
     * @return 图片byte数组
     * @throws IOException 图片读取IO错误
     */
    <E extends T> byte[] loadPicture(E e) throws IOException;
}
