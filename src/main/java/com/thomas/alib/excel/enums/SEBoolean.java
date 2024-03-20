package com.thomas.alib.excel.enums;

/**
 * Boolean类型枚举
 */
public enum SEBoolean {
    NOT_SET(null),//不设置
    TRUE(true),//true
    FALSE(false);//false


    private final Boolean value;

    SEBoolean(Boolean value) {
        this.value = value;
    }

    public Boolean getValue() {
        return value;
    }

    /**
     * 根据value值获取枚举
     *
     * @param value value值
     * @return 枚举类型
     */
    public static SEBoolean getFromValue(Boolean value) {
        if (value == null) return NOT_SET;
        return value ? TRUE : FALSE;
    }

}
