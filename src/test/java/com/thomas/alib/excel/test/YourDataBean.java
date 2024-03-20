package com.thomas.alib.excel.test;

import com.thomas.alib.excel.annotation.ExcelColumn;
import com.thomas.alib.excel.annotation.ExcelSheet;
import com.thomas.alib.excel.annotation.ExcelStyle;
import com.thomas.alib.excel.converter.LocalDateConverter;
import com.thomas.alib.excel.enums.SEBoolean;
import com.thomas.alib.excel.enums.SEBorderStyle;
import com.thomas.alib.excel.enums.SEHorAlignment;
import com.thomas.alib.excel.enums.SEVerAlignment;

import java.time.LocalDate;

@ExcelSheet(sheetName = "YourDataBean示例", showIndex = false, dataRowHeight = 3000,
        baseStyle = @ExcelStyle(fontName = "宋体", horAlignment = SEHorAlignment.CENTER, verAlignment = SEVerAlignment.CENTER, isWrapText = SEBoolean.TRUE, borderStyle = SEBorderStyle.THIN),
        headStyle = @ExcelStyle(textSize = 12, isBold = SEBoolean.TRUE, bgColor = "#E3E3E3"),
        dataStyle = @ExcelStyle(textSize = 10, textColor = "#8F8F8F"))
public class YourDataBean {

    @ExcelColumn(headerName = "姓名", orderNum = 1, columnStyle = @ExcelStyle(fontName = "黑体", textSize = 25), columnStyleInHead = true)
    private String userName;
    @ExcelColumn(headerName = "出生日期", orderNum = 2, columnWidth = 6000, converter = LocalDateConverter.class)
    private LocalDate birthDay;
    @ExcelColumn(headerName = "性别", orderNum = 3, columnWidth = 2000, beforeConvert = {"0", "1"}, afterConvert = {"男", "女"})
    private Integer gender;
    @ExcelColumn(headerName = "头像", orderNum = 4, columnWidth = 4500, isPicture = true, columnStyle = @ExcelStyle(borderStyle = SEBorderStyle.MEDIUM_DASH_DOT_DOT, borderColor = "#00FF00"))
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
