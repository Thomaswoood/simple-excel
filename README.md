# simple-excel插件工具

## 介绍
基于apache.poi二次开发，使用注解方式调用poi针对excel处理的常用api。支持根据List<?>和配置的注解进行导入导出。
注解支持自动序号、列排序、行高、列宽、前后缀处理、自定义转化器处理、图片加载、字体对齐边框等基本样式处理。
支持通过javax.validation进行导入参数校验。

现在迎来了全新的2.0.0版本：重构了导出样式处理器；支持叠加设置多套样式；支持在@ExcelSheet和@ExcelColumn注解中使用样式的组合注解@ExcelStyle；
支持直接导出到指定路径或Path；集成了日志门面slf4j，您可以自行选择您要使用的日志实现库；细化了异常捕获和日志打印逻辑，方便您定位问题所在。

您可以在@ExcelSheet中设置全局通用的基础样式baseStyle，单独给表头行使用的headStyle，单独给数据行使用的dataStyle，他们都会继承或覆盖baseStyle中设置的样式。
```
@ExcelSheet(baseStyle = @ExcelStyle(fontName = "宋体"), // 这样全局导出样式的字体都会是宋体
        headStyle = @ExcelStyle(textSize = 12), // 在继承baseStyle的同时，设置字号为12
        dataStyle = @ExcelStyle(textSize = 10)) // 在继承baseStyle的同时，设置字号为12
```
您也可以在@ExcelColumn注解中设置该列的样式，他会在dataStyle基础上覆盖生效，也可以控制该列的样式是否在表头行生效，若在表头行生效，则会在headStyle基础上覆盖生效。
```
// 比如您想突出显示某一列，可以单独在列中设置字体和字号，就会在继承dataStyle的同时，覆盖生效该列的样式。
@ExcelColumn(columnStyle = @ExcelStyle(fontName = "黑体", textSize = 25),
        columnStyleInHead = true) // 也可以配置在表头行生效，那么这一列的表头，就会在继承headStyle的同时，覆盖生效该列的样式。
```

## 使用说明
> 安装-Maven
```
<dependency>
    <groupId>io.github.Thomaswoood</groupId>
    <artifactId>simple-excel</artifactId>
    <version>2.0.0</version>
</dependency>
```

> 类注解示例
```
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

3将Excel导出至指定路径中：

```
ExcelExportSimple.with("/data/test.xlxs")
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
