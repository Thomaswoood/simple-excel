### simple-excel插件工具
基于apache.poi二次开发，使用注解方式调用poi针对excel处理的常用api。支持根据List<?>和配置的注解进行导入导出。
注解支持自动序号、列排序、行高、列宽、前后缀处理、自定义转化器处理、图片加载、字体对齐边框等基本样式处理。
支持通过javax.validation进行导入参数校验。

> 安装-Maven
```
<dependency>
    <groupId>io.github.Thomaswoood</groupId>
    <artifactId>simple-excel</artifactId>
    <version>1.0.3</version>
</dependency>
```

> 类注解示例
```
@ExcelSheet(sheetName = "YourDataBean示例", showIndex = false, dataRowHeight = 3000)
@ExcelExportHeadStyle(fontName = "宋体", textSize = 12, isBold = true, horAlignment = HorAlignment.CENTER, verAlignment = VerAlignment.CENTER, isWrapText = true, fourSideBorder = BorderType.THIN, bgColor = "#E3E3E3")
@ExcelExportDataStyle(fontName = "宋体", textSize = 10, horAlignment = HorAlignment.CENTER, verAlignment = VerAlignment.CENTER, isWrapText = true, fourSideBorder = BorderType.THIN, textColor = "#8F8F8F")
public class YourDataBean {

    @ExcelColumn(headerName = "姓名", orderNum = 1)
    private String userName;
    @ExcelColumn(headerName = "出生日期", orderNum = 2, columnWidth = 6000, converter = LocalDateConverter.class)
    private LocalDate birthDay;
    @ExcelColumn(headerName = "性别", orderNum = 3, columnWidth = 2000, beforeConvert = {"0", "1"}, afterConvert = {"男", "女"})
    private Integer gender;
    @ExcelColumn(headerName = "头像", orderNum = 4, columnWidth = 4500, isPicture = true)
    private String headPicture;

}
```

> 导出示例：

1. 将Excel直接导出至response流中：

```
ExcelExportSimple.with(response)
                .addSheet(exportList)
                .fileName("自定义文件名.xlsx")
                .export();
```

2. 将Excel导出至指定File中：

```
ExcelExportSimple.with(File)
                .addSheet(exportList)
                .export();
```

> 导入示例，导入支持以File或InputStream作为输入源，下方以File为示例：

1. 正常读取模式，中途出错以异常形式抛出（主要为导入数据类型不对或校验不通过异常），并且出错后自定停止后续导入读取：
```
List<YourDataBean> dataList = ExcelImportSimple.readList(file, YourDataBean.class);
```
2. 安全读取模式，中途所有错误和异常将被封装为SafetyResult记录并返回，包括整体汇总和每条的信息，会读取完全部数据后才停止：
```
SafetyResult<List<SafetyResult<YourDataBean>>> readResult = ExcelImportSimple.readListSafety(file, YourDataBean.class);
```

> 导入模式1.0.2新功能，支持以lambda表达式的方式传入InputStream流，并且由工具内部消化可能抛出的异常进行统一处理：
1. 正常读取模式，中途出错以异常形式抛出（主要为导入数据类型不对或校验不通过异常），并且出错后自定停止后续导入读取：
```
List<YourDataBean> dataList = ExcelImportSimple.readList(FileInputStream::new, file, YourDataBean.class);
```
2. 安全读取模式，中途所有错误和异常将被封装为SafetyResult记录并返回，包括整体汇总和每条的信息，会读取完全部数据后才停止：
```
SafetyResult<List<SafetyResult<YourDataBean>>> readResult = ExcelImportSimple.readListSafety(FileInputStream::new, file, YourDataBean.class);
```
3. 你也可以在表单上传文件时直接这样使用：
```
List<YourDataBean> dataList = ExcelImportSimple
        .readList(MultipartFile::getInputStream, multipartFile, YourDataBean.class);
//or
SafetyResult<List<SafetyResult<YourDataBean>>> readResult = ExcelImportSimple
        .readListSafety(MultipartFile::getInputStream, multipartFile, YourDataBean.class);
```
