package com.thomas.alib.excel.importer.validation;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * 导入验证器工厂
 */
public class ImportValidatorFactory {
    private static Logger logger = LoggerFactory.getLogger(ImportValidatorFactory.class);

    /**
     * 根据实际依赖情况，创建不同的验证器
     *
     * @return 验证器对象
     */
    public static ImportValidator build() {
        ImportValidator result = null;
        try {
            Class.forName("javax.validation.ValidatorFactory");
            // 如果能加载成功，说明使用 javax.validation 包
            logger.debug("检测到javax.validation包，即将使用javax.validation.ValidatorFactory生成验证器");
            result =  new JavaxValidator();
            logger.debug("基于javax.validation的验证器初始化完成，导入时将基于javax.validation验证");
        } catch (Throwable e) {
            logger.warn("基于javax.validation的验证器初始化失败，可能您没有添加验证器相关依赖，继续检测jakarta.validation包", e);
            // 如果加载失败，则尝试加载 jakarta.validation 包
            try {
                Class.forName("jakarta.validation.ValidatorFactory");
                logger.debug("检测到jakarta.validation包，即将使用jakarta.validation.ValidatorFactory生成验证器");
                result = new JakartaValidator();
                logger.debug("基于jakarta.validation的验证器初始化完成，导入时将基于jakarta.validation验证");
            } catch (Throwable ex) {
                ex.addSuppressed(e);
                // 如果两个都加载失败
                logger.warn("验证器初始化失败，可能您没有添加验证器相关依赖，导入时基于javax.validation或jakarta.validation的验证将不会启用", ex);
            }
        }
        return result;
    }
}
