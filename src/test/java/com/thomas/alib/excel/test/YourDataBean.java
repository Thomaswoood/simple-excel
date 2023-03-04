package com.thomas.alib.excel.test;

import com.thomas.alib.excel.annotation.ExcelColumn;
import com.thomas.alib.excel.annotation.ExcelExportDataStyle;
import com.thomas.alib.excel.annotation.ExcelExportHeadStyle;
import com.thomas.alib.excel.annotation.ExcelSheet;
import com.thomas.alib.excel.converter.LocalDateConverter;
import com.thomas.alib.excel.enums.BorderType;
import com.thomas.alib.excel.enums.HorAlignment;
import com.thomas.alib.excel.enums.VerAlignment;

import java.time.LocalDate;

@ExcelSheet(sheetName = "YourDataBean示例", showIndex = false)
@ExcelExportHeadStyle(fontName = "宋体", textSize = 12, isBold = true, horAlignment = HorAlignment.CENTER, verAlignment = VerAlignment.CENTER, isWrapText = true, fourSideBorder = BorderType.THIN)
@ExcelExportDataStyle(fontName = "宋体", textSize = 10, horAlignment = HorAlignment.CENTER, verAlignment = VerAlignment.CENTER, isWrapText = true, fourSideBorder = BorderType.THIN)
public class YourDataBean {

    @ExcelColumn(headerName = "姓名", orderNum = 1)
    private String userName;
    @ExcelColumn(headerName = "出生日期", orderNum = 2, columnWidth = 6000, converter = LocalDateConverter.class)
    private LocalDate birthDay;
    @ExcelColumn(headerName = "性别 0男 1女", orderNum = 3, columnWidth = 1300, beforeConvert = {"0", "1"}, afterConvert = {"男", "女"})
    private Integer gender;
    @ExcelColumn(headerName = "头像", orderNum = 4, columnWidth = 3800, isPicture = true)
    private String headPicture;

}
