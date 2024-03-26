package com.thomas.alib.excel.importer.validation;

import java.util.Set;

/**
 * 导入验证器
 */
public interface ImportValidator extends AutoCloseable {
    /**
     * 验证入参对象中包含的所有约束
     *
     * @param t   待验证入参对象，如果传空对象将被忽略
     * @param <T> 参数泛型
     * @return 验证不通过的结果集，如果全部验证通过，返回空
     */
    <T> Set<String> validate(T t);
}
