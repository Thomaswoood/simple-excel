package com.thomas.alib.excel.exporter;

import com.thomas.alib.excel.annotation.ExcelSheet;
import com.thomas.alib.excel.enums.SortType;
import com.thomas.alib.excel.utils.CollectionUtils;
import com.thomas.alib.excel.utils.ReflectUtil;
import com.thomas.alib.excel.utils.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

/**
 * Excel导出对应每个工作表(Sheet)的解析工具
 *
 * @param <T> 数据源对象类型泛型
 */
public class ExcelExportSheetItem<T, EE extends ExcelExporterBase<EE>> {
    private static Logger logger = LoggerFactory.getLogger(ExcelExportSheetItem.class);
    /**
     * excel导出构建对象
     */
    private EE excelExporter;
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
     * 表头行样式
     */
    private XSSFCellStyle headStyle;
    /**
     * 数据行样式
     */
    private XSSFCellStyle dataStyle;
    /**
     * 全部的成员列表
     */
    private List<Field> totalFieldList;
    /**
     * 全部的成员对应的excel列信息列表
     */
    private List<ExcelExportColumnItem> excelColumnList;

    /**
     * 构造方法
     *
     * @param excel_exporter   excel导出构建对象
     * @param source_list      数据源列表
     * @param sheet_data_clazz 数据源泛型class
     * @param show_index       是否显示序号
     * @param sheet_name       外部设置的sheet名称，如果有值，则此值优先级比注解内的高
     */
    ExcelExportSheetItem(EE excel_exporter, List<T> source_list, Class<T> sheet_data_clazz, Boolean show_index, String sheet_name) {
        excelExporter = excel_exporter;
        sourceList = source_list;
        //获取数据类型
        sheetDataClazz = sheet_data_clazz;
        if (sheetDataClazz == null) throw new RuntimeException("获取导出数据类型失败");
        //取得表格sheet相关注解
        ExcelSheet excel_sheet = sheetDataClazz.getAnnotation(ExcelSheet.class);
        //判断是否传入了sheet_name，传入的参数值优先级更高，单独判断处理
        if (StringUtils.isEmpty(sheet_name)) {
            if (excel_sheet == null) {//未设置，设置默认值
                sheetName = sheetDataClazz.getSimpleName();
            } else {//已设置，取出设置的值
                sheetName = excel_sheet.sheetName();
            }
        } else {
            sheetName = sheet_name;
        }
        //判断是否传入了show_index，传入的参数值优先级更高，单独判断处理
        if (show_index == null) {//如果需要显示序号，则添加序号
            if (excel_sheet == null) {//未设置，设置默认值
                showIndex = true;
            } else {//已设置，取出设置的值
                showIndex = excel_sheet.showIndex();
            }
        } else {
            showIndex = show_index;
        }
        //表头样式处理
        ExcelExportStyleProcessor head_style_processor;

        //数据样式处理
        ExcelExportStyleProcessor data_style_processor;
        //表格sheet注解判空区分处理
        if (excel_sheet == null) {
            //未设置sheet注解，设置默认值
            columnSortType = SortType.S2B;//列排序方式，默认由小到大排序
            headRowHeight = -1;//表头行高度，默认-1不设置
            dataRowHeight = -1;//数据行高度，默认-1不设置
            head_style_processor = null;
            data_style_processor = null;
        } else {
            //已设置sheet注解，取出设置的值
            columnSortType = excel_sheet.columnSortType();//列排序方式
            headRowHeight = excel_sheet.headRowHeight();//表头行高度
            dataRowHeight = excel_sheet.dataRowHeight();//数据行高度
            //读取表头行样式
            head_style_processor = ExcelExportStyleProcessor.read(excel_sheet.baseStyle()).coverBySourceExceptNotSet(excel_sheet.headStyle());
            headStyle = head_style_processor.getXSSFCellStyle(excelExporter.sxssfWorkbook);
            //读取数据行样式
            data_style_processor = ExcelExportStyleProcessor.read(excel_sheet.baseStyle()).coverBySourceExceptNotSet(excel_sheet.dataStyle());
            dataStyle = data_style_processor.getXSSFCellStyle(excelExporter.sxssfWorkbook);
        }
        //解析全部成员属性
        totalFieldList = ReflectUtil.getAccessibleFieldIncludeSuper(sheetDataClazz);//全部的成员列表
        excelColumnList = new ArrayList<>();//解析成员为excel列信息列表
        for (Field item_field : totalFieldList) {//遍历寻找所有有效column成员
            ExcelExportColumnItem column = new ExcelExportColumnItem(item_field, excelExporter.sxssfWorkbook, head_style_processor, data_style_processor);
            if (column.isValid()) //有效
                excelColumnList.add(column);
        }
        //给列排序
        switch (columnSortType) {
            case B2S:
                Collections.reverse(excelColumnList);
            case S2B:
            default:
                Collections.sort(excelColumnList);
        }
    }

    /**
     * 初始化sheet并设置数据到sheet中
     */
    void writeData() {
        logger.debug(sheetName() + "sheet页签准备处理数据");
        //根据解析信息创建excel数据
        Sheet sheet = excelExporter.sxssfWorkbook.createSheet(sheetName());//创建sheet对象
        //创建并填充表头信息
        Row headRow = sheet.createRow(0);//创建表头
        //设置表头行高
        if (headRowHeight > 0) headRow.setHeight(headRowHeight);
        int r_i = 0;//外部给row计数，因为是否自动显示序号会影响起始位置
        if (showIndex) {//如果需要默认显示序号
            Cell cell = headRow.createCell(r_i);
            cell.setCellValue("序号");
            if (headStyle != null) cell.setCellStyle(headStyle);
            sheet.setColumnWidth(r_i, 3000);
            r_i++;
        }
        for (ExcelExportColumnItem column : excelColumnList) {
            Cell cell = headRow.createCell(r_i);
            cell.setCellValue(column.getHeadName());
            if (column.getHeadStyle() != null) {
                cell.setCellStyle(column.getHeadStyle());
            } else if (headStyle != null) {
                cell.setCellStyle(headStyle);
            }
            sheet.setColumnWidth(r_i, column.getColumnWidth());
            r_i++;
        }
        logger.debug("表头行处理完成");
        if (!CollectionUtils.isEmpty(sourceList)) {
            //创建并填充每行数据信息
            for (int i = 0; i < sourceList.size(); i++) {//创建每一行数据
                int row_index = i + 1;//行序号
                T item_source = sourceList.get(i);//逐个取出数据源
                Row dataRow = sheet.createRow(row_index);//创建数据行
                //设置数据行高
                if (dataRowHeight > 0) dataRow.setHeight(dataRowHeight);
                r_i = 0;//每新开始一行，重置row计数
                if (showIndex) {//如果需要默认显示序号
                    Cell cell = dataRow.createCell(r_i);
                    cell.setCellValue(row_index);
                    if (dataStyle != null) cell.setCellStyle(dataStyle);
                    r_i++;
                }
                for (ExcelExportColumnItem column : excelColumnList) {
                    Cell cell = dataRow.createCell(r_i);
                    if (column.isPicture()) {
                        try {
                            byte[] pictureBytes = column.getColumnPictureBytesFromSource(item_source, row_index);
                            if (pictureBytes == null) {
                                cell.setCellValue("图片： " + column.getColumnValueFromSource(item_source, row_index) + "  加载失败");
                            } else {
                                Drawing drawing = sheet.getDrawingPatriarch();
                                if (drawing == null) drawing = sheet.createDrawingPatriarch();
                                ClientAnchor clientAnchor = drawing.createAnchor(0, 0, 0, 0, r_i, row_index, r_i + 1, i + 2);
                                int addPicture = excelExporter.sxssfWorkbook.addPicture(pictureBytes, SXSSFWorkbook.PICTURE_TYPE_JPEG);
                                Picture picture = drawing.createPicture(clientAnchor, addPicture);
                                picture.getPictureData();
                            }
                        } catch (Throwable e) {
                            logger.error("表格\"" + column.getHeadName() + "\"列-第" + row_index + "行-绘制图片时发生错误:", e);
                            cell.setCellValue("图片： " + column.getColumnValueFromSource(item_source, row_index) + "  绘制失败");
                        }
                    } else {
                        cell.setCellValue(column.getColumnValueFromSource(item_source, row_index));
                    }
                    if (column.getDataStyle() != null) {
                        cell.setCellStyle(column.getDataStyle());
                    } else if (dataStyle != null) {
                        cell.setCellStyle(dataStyle);
                    }
                    r_i++;
                }
                logger.debug("表格第" + row_index + "行处理完成");
            }
        }
        logger.debug("数据行处理完成");
        logger.debug(sheetName() + "sheet页签数据处理完成");
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

    /**
     * 结束当前创建的sheet页签对象的方法调用，返回导出者，继续后续导出操作
     * （目前暂无实际业务场景，为未来扩展预留的方法）
     *
     * @return 导出者
     */
    public EE over() {
        return excelExporter;
    }
}
