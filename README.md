# ImportAndExportExcel
Java 环境下 Excel 的导入导出

# 说明
## HSSFWorkbook
Excel 97-2003，即 HSSFWorkbook 只需要导入：
`
poi-3.9-20121203.jar/
log4j-1.2.13.jar/
slf4j-api-1.6.1.jar
`

## XSSFWorkbook
Excel > 2003，即 XSSFWorkbook 还需要导入：
`
poi-ooxml-3.9-20121203.jar/
poi-ooxml-schemas-3.9-20121203.jar/
xmlbeans-2.3.0.jar/
dom4j-1.6.1.jar
`

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
## ExportExcelUtil API
只通过三个重载的构造器向外提供了3个接口，最完整的一个构造器方法签名为：
`com.solverpeng.poi.utils.excel.ExportExcelUtil.downLoad4Excel(java.lang.String, java.util.List<java.lang.String>, java.util.List<java.lang.Object[]>, java.lang.String, int, java.lang.String, javax.servlet.http.HttpServletRequest, javax.servlet.http.HttpServletResponse)`

方法体为内容分为：
1. 获取样式
2. 设置标题
3. 设置表头
4. 填充数据
5. 下载

### 说明
1. 填充的数据类型为：`List<Object[]>` 类型。
2. 导出使用原生 `Servlet` 进行的下载，以相应流的方式输出到浏览器。

### 使用
参见 `com.solverpeng.poi.servlet.PoiServlet2.doPost` 方法
```java
protected void doPost(HttpServletRequest req, HttpServletResponse resp) throws ServletException, IOException {
    User user1 = new User("1", "tom", 23, new Date(), "北京");
    User user2 = new User("2", "jerry", 25, new Date(), "上海");
    User user3 = new User("3", "lily", 26, new Date(), "洛杉矶");
    User user4 = new User("4", "lucy", 27, new Date(), "纽约");
    List<User> list = Arrays.asList(user1, user2, user3, user4);

    List<Object[]> objects = new ArrayList<>();
    if (CollectionUtils.isNotEmpty(list)) {
        for (User user : list) {
            Object[] objectArr = new Object[4];
            objectArr[0] = user.getUserName();
            objectArr[1] = user.getAge();
            objectArr[2] = user.getBirth();
            objectArr[3] = user.getAddress();
            objects.add(objectArr);
        }
    }

    String titleName = "用户信息";
    List<String> headers = Arrays.asList("用户姓名", "用户年龄", "用户生日", "居住地址");
    ExportExcelUtil.downLoad4Excel(titleName, headers, objects, "用户信息", req, resp);
}
```

## ExportExcel 和 ExcelField API
相较于 ExportExcelUtil 的灵活，ExportExcel 和 ExcelField 的方式稍微比较繁琐，需要对导出的实体类字段添加 `ExcelField` 注解。
但是对导入的数据支持 `List<Entity>`。
### `ExcelField` 注解类
对其中的两个属性进行说明：
sort字段：按 DESC 进行的排序，即值越大，则对应的表头越靠后。
groups字段：针对同一实体的不同字段导出，在 `ExportExcel` 可以传入。
### 使用
可以参见：`com.solverpeng.poi.servlet.PoiServlet.doPost` 方法的各个测试，以及对应实体 `com.solverpeng.poi.beans.User` 使用的注解。

# 导入
## ImportExcel API
### 说明
实现的原理类似于 ExportExcel，详细内容可以对比学习。
### 使用
ImportExcelTest 测试了在 Java 静态工程中的使用方式，ExcelHandler 是 SpringMVC 下的一个导入Demo。