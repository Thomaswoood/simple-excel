package com.thomas.alib.excel.interfaces;

/**
 * 将此函数应用于给定参数，并且可将函数内异常消化掉
 *
 * @param <T> 函数源类型的泛型
 * @param <R> 函数执行结果的泛型
 */
public interface EFunction<T, R> {

    /**
     * 将此函数应用于给定参数，并且可将函数内异常消化掉
     *
     * @param t 函数变量源
     * @return 函数执行结果
     */
    R apply(T t) throws Exception;
}
