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
    @ExcelColumn(headerName = "性别", orderNum = 3, columnWidth = 2000, beforeConvert = {"0", "1"}, afterConvert = {"男", "女"})
    private Integer gender;
    @ExcelColumn(headerName = "头像", orderNum = 4, columnWidth = 4500, isPicture = true)
    private String headPicture;

    public YourDataBean() {
    }

    public YourDataBean(String userName, LocalDate birthDay, Integer gender, String headPicture) {
        this.userName = userName;
        this.birthDay = birthDay;
        this.gender = gender;
        this.headPicture = headPicture;
    }

    public String getUserName() {
        return userName;
    }

    public void setUserName(String userName) {
        this.userName = userName;
    }

    public LocalDate getBirthDay() {
        return birthDay;
    }

    public void setBirthDay(LocalDate birthDay) {
        this.birthDay = birthDay;
    }

    public Integer getGender() {
        return gender;
    }

    public void setGender(Integer gender) {
        this.gender = gender;
    }

    public String getHeadPicture() {
        return headPicture;
    }

    public void setHeadPicture(String headPicture) {
        this.headPicture = headPicture;
    }
}
