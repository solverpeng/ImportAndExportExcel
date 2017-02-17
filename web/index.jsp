<%--
  Created by IntelliJ IDEA.
  User: solverpeng
  Date: 2017/2/16 0016
  Time: 15:26
--%>
<%@ page contentType="text/html;charset=UTF-8" language="java" %>
<html>
<head>
    <title>Index Page</title>
</head>
<body>
<center>
    <a href="poi/PoiServlet">export excel</a>
    <br>
    <a href="poi/PoiServlet2">export excel2</a>
    <br><br>
    <form id="importForm" action="importExcel" method="post" enctype="multipart/form-data">
        <input name="file" type="file"/>
        <input type="submit" value="   导    入   "/>
    </form>
</center>
</body>
</html>
