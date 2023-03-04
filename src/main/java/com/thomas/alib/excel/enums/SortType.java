package com.thomas.alib.excel.enums;

import java.util.Objects;

/**
 * excel导出列排序方式
 */
public enum SortType {
    S2B(1, "由小到大排序"),
    B2S(2, "由大到小排序");

    private int code;
    private String name;

    SortType(int code, String name) {
        this.code = code;
        this.name = name;
    }

    public int getCode() {
        return code;
    }

    public String getName() {
        return name;
    }

    /**
     * 根据code值获取枚举
     *
     * @param code code值
     * @return 枚举类型
     */
    public static SortType getFromCode(Integer code) {
        if (code == null) return S2B;
        for (SortType item : SortType.values()) {
            if (item.code == code) {
                return item;
            }
        }
        return S2B;
    }

    /**
     * 根据name值获取枚举
     *
     * @param name name值
     * @return 枚举类型
     */
    public static SortType getFromName(String name) {
        if (name == null) return S2B;
        for (SortType item : SortType.values()) {
            if (Objects.equals(item.name, name)) {
                return item;
            }
        }
        return S2B;
    }
}
