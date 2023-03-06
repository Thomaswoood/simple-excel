package com.thomas.alib.excel.exporter;

import com.thomas.alib.excel.annotation.ExcelSheet;
import com.thomas.alib.excel.enums.SortType;
import com.thomas.alib.excel.utils.CollectionUtils;
import com.thomas.alib.excel.utils.ReflectUtil;
import com.thomas.alib.excel.utils.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

/**
 * Excel导出对应每个工作表(Sheet)的解析工具
 *
 * @param <T> 数据源对象类型泛型
 */
class ExcelExportSheetItem<T> {
    /**
     * excel导出构建对象
     */
    private Workbook workbook;
    /**
     * 数据源列表
     */
    private List<T> sourceList;
    /**
     * sheet数据类型对象
     */
    private Class<T> sheetDataClazz;
    /**
     * sheet名称
     */
    private String sheetName;
    /**
     * 本sheet是否显示序号
     */
    private boolean showIndex;
    /**
     * 列排序方式
     */
    private SortType columnSortType;

    /**
     * 表头行高度，默认-1不设置
     */
    private short headRowHeight;

    /**
     * 数据行高度，默认-1不设置
     */
    private short dataRowHeight;
    /**
     * 表头行样式处理者
     */
    private ExcelExportStyleProcessor headStyleProcessor;
    /**
     * 数据行样式处理者
     */
    private ExcelExportStyleProcessor dataStyleProcessor;

    /**
     * 构造方法
     *
     * @param workbook         excel导出构建对象
     * @param source_list      数据源列表
     * @param sheet_data_clazz 数据源泛型class
     * @param show_index       是否显示序号
     * @param sheet_name       外部设置的sheet名称，如果有值，则此值优先级比注解内的高
     */
    ExcelExportSheetItem(Workbook workbook, List<T> source_list, Class<T> sheet_data_clazz, Boolean show_index, String sheet_name) {
        this.workbook = workbook;
        this.sourceList = source_list;
        //获取数据类型
        this.sheetDataClazz = sheet_data_clazz;
        if (sheetDataClazz == null) throw new RuntimeException("获取导出数据类型失败");
        //取得表格sheet相关注解
        ExcelSheet excel_sheet = sheetDataClazz.getAnnotation(ExcelSheet.class);
        if (StringUtils.isEmpty(sheet_name)) {
            if (excel_sheet == null) {//未设置，设置默认值
                sheetName = sheetDataClazz.getSimpleName();
            } else {//已设置，取出设置的值
                sheetName = excel_sheet.sheetName();
            }
        } else {
            sheetName = sheet_name;
        }
        if (show_index == null) {//如果需要显示序号，则添加序号
            if (excel_sheet == null) {//未设置，设置默认值
                showIndex = true;
            } else {//已设置，取出设置的值
                showIndex = excel_sheet.showIndex();
            }
        } else {
            showIndex = show_index;
        }
        if (excel_sheet == null) {
            //未设置sheet注解，设置默认值
            columnSortType = SortType.S2B;//列排序方式，默认由小到大排序
            headRowHeight = -1;//表头行高度，默认-1不设置
            dataRowHeight = -1;//数据行高度，默认-1不设置
        } else {
            //已设置sheet注解，取出设置的值
            columnSortType = excel_sheet.columnSortType();//列排序方式
            headRowHeight = excel_sheet.headRowHeight();//表头行高度
            dataRowHeight = excel_sheet.dataRowHeight();//数据行高度
        }
        //表头样式处理
        headStyleProcessor = ExcelExportStyleProcessor.withHead(workbook, sheetDataClazz);
        //数据样式处理
        dataStyleProcessor = ExcelExportStyleProcessor.withData(workbook, sheetDataClazz);
    }

    /**
     * 初始化sheet并设置数据到sheet中
     */
    void writeData() {
        //解析全部成员属性
        List<Field> total_field_list = ReflectUtil.getAccessibleFieldIncludeSuper(sheetDataClazz);//全部的成员列表
        List<ExcelExportColumnItem> excel_column_list = new ArrayList<>();//解析成员为excel列信息列表
        for (Field item_field : total_field_list) {//遍历寻找所有有效column成员
            ExcelExportColumnItem item_column_field = new ExcelExportColumnItem(item_field);
            if (item_column_field.isValid()) //有效
                excel_column_list.add(item_column_field);
        }
        switch (columnSortType) {
            case B2S:
                Collections.reverse(excel_column_list);//给列排序
            case S2B:
            default:
                Collections.sort(excel_column_list);//给列排序
        }
        //根据解析信息创建excel数据
        Sheet sheet = workbook.createSheet(sheetName());//创建sheet对象
        //创建并填充表头信息
        Row headRow = sheet.createRow(0);//创建表头
        //设置表头行高
        if (headRowHeight > 0) headRow.setHeight(headRowHeight);
        int r_i = 0;//外部给row计数，因为是否自动显示序号会影响起始位置
        if (showIndex) {//如果需要默认显示序号
            Cell cell = headRow.createCell(r_i);
            cell.setCellValue("序号");
            headStyleProcessor.setStyle(cell);
            sheet.setColumnWidth(r_i, 3000);
            r_i++;
        }
        for (ExcelExportColumnItem column : excel_column_list) {
            Cell cell = headRow.createCell(r_i);
            cell.setCellValue(column.getHeadName());
            headStyleProcessor.setStyle(cell);
            sheet.setColumnWidth(r_i, column.getColumnWidth());
            r_i++;
        }
        if (!CollectionUtils.isEmpty(sourceList)) {
            //创建并填充每行数据信息
            for (int i = 0; i < sourceList.size(); i++) {//创建每一行数据
                T item_source = sourceList.get(i);//逐个取出数据源
                Row dataRow = sheet.createRow(i + 1);//创建数据行
                //设置数据行高
                if (dataRowHeight > 0) dataRow.setHeight(dataRowHeight);
                r_i = 0;//每新开始一行，重置row计数
                if (showIndex) {//如果需要默认显示序号
                    Cell cell = dataRow.createCell(r_i);
                    cell.setCellValue(i + 1);
                    dataStyleProcessor.setStyle(cell);
                    r_i++;
                }
                for (ExcelExportColumnItem column : excel_column_list) {
                    Cell cell = dataRow.createCell(r_i);
                    if (column.isPicture()) {
                        try {
                            byte[] pictureBytes = column.getColumnPictureBytesFromSource(item_source);
                            if (pictureBytes == null) {
                                cell.setCellValue("图片： " + column.getColumnValueFromSource(item_source) + "  加载失败");
                            } else {
                                int addPicture = workbook.addPicture(pictureBytes, workbook.PICTURE_TYPE_PNG);
                                Drawing drawing = sheet.createDrawingPatriarch();
                                ClientAnchor clientAnchor = new XSSFClientAnchor(0, 0, 0, 0, r_i, i + 1, r_i + 1, i + 2);
                                Picture picture = drawing.createPicture(clientAnchor, addPicture);
                            }
                        } catch (Throwable e) {
                            cell.setCellValue("图片： " + column.getColumnValueFromSource(item_source) + "  获取失败");
                        }
                    } else {
                        cell.setCellValue(column.getColumnValueFromSource(item_source));
                    }
                    dataStyleProcessor.setStyle(cell);
                    r_i++;
                }
            }
        }
    }

    /**
     * 获取sheet名
     *
     * @return sheet名
     */
    String sheetName() {
        if (StringUtils.isEmpty(sheetName)) {
            return String.valueOf(System.currentTimeMillis());
        } else {
            return sheetName;
        }
    }
}
