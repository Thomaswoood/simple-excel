package com.thomas.alib.excel.importer.validation;

import com.thomas.alib.excel.utils.CollectionUtils;
import jakarta.validation.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.Set;
import java.util.stream.Collectors;

/**
 * 基于jakarta.validation的验证器
 */
public class JakartaValidator implements ImportValidator {
    private static Logger logger = LoggerFactory.getLogger(JakartaValidator.class);
    /**
     * 验证器工厂
     */
    ValidatorFactory validatorFactory = null;
    /**
     * 验证器对象
     */
    Validator validator = null;

    /**
     * 构造方法，创建验证器工厂，并初始化验证器
     */
    public JakartaValidator() {
        validatorFactory = Validation.buildDefaultValidatorFactory();
        validator = validatorFactory.getValidator();
    }

    /**
     * 验证入参对象中包含的所有约束
     *
     * @param t   待验证入参对象，如果传空对象将被忽略
     * @param <T> 参数泛型
     * @return 验证不通过的结果集，如果全部验证通过，返回空
     */
    @Override
    public <T> Set<String> validate(T t) {
        //验证器存在才进行验证判断
        if (t != null && validator != null) {
            Set<ConstraintViolation<@Valid T>> validateSet = validator.validate(t);
            if (!CollectionUtils.isEmpty(validateSet)) {
                return validateSet.stream().map(ConstraintViolation::getMessage).collect(Collectors.toSet());
            }
        }
        return null;
    }

    /**
     * 关闭，目前仅用于关闭验证器工厂
     */
    @Override
    public void close() {
        //如果验证器工厂存在，关闭他
        if (validatorFactory != null) {
            try {
                validatorFactory.close();
            } catch (Throwable e) {
                logger.error("验证器工厂关闭时发生错误:", e);
            }
        }
    }
}
