# 说明
## HSSFWorkbook
Excel 97-2003，即 HSSFWorkbook 只需要导入：
poi-3.9-20121203.jar
log4j-1.2.13.jar
slf4j-api-1.6.1.jar

## XSSFWorkbook
Excel > 2003，即 XSSFWorkbook 还需要导入：
poi-ooxml-3.9-20121203.jar 
poi-ooxml-schemas-3.9-20121203.jar
xmlbeans-2.3.0.jar
dom4j-1.6.1.jar

# 备份
## ExportExcel_bk
保留了：handleAnnotationMethods() 方法，若需要对方法进行标注，则将注释去掉就好。
## ExportExcel_bk2
若是单独作为一个导出使用，则这个版本已经为最简版本，考虑到低版本的 office ，所以使用的是 HSSFWorkbook，
即只需要导入 `poi-3.9-20121203.jar/log4j-1.2.13.jar/slf4j-api-1.6.1.jar` 这三个包即可。

## ImportExcel_bk
保留了：handleAnnotationMethods() 方法，若需要对方法进行标注，则将注释去掉就好。
## ImportExcel_bk2
若是作为一个单独的导出使用，这个版本已经为最简版本，考虑到导入的时候，可能是高版本的 office ，所以提供了 xlsx 导入

# 导出
PoiServlet 测试了在 Web 环境下如何导出以及相关API如何使用

# 导入
ImportExcelTest 测试了在 Java 静态工程中的使用方式，ExcelHandler 是 SpringMVC 下的一个导入Demo。