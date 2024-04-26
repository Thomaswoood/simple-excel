package com.thomas.alib.excel.enums;

import java.util.Comparator;
import java.util.List;
import java.util.function.Supplier;

/**
 * excel导出列排序方式
 */
public enum SortType {
    S2B(Comparator::naturalOrder),//由小到大排序
    B2S(Comparator::reverseOrder),//由大到小排序
    ;

    Supplier<?> comparator;

    SortType(Supplier<?> comparator) {
        this.comparator = comparator;
    }

    public <T extends Comparable<T>> void sort(List<T> list) {
        list.sort((Comparator<T>) comparator.get());
    }
}
