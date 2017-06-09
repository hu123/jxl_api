<%@ page contentType="text/html;charset=UTF-8" language="java" %>
<%@ page isELIgnored="false" %>
<html>
<head>
    <title>Title</title>
</head>
<body>

欢迎教师下载模板文件录入成绩-----><a href="/download">下载样板文件</a><br/>
<hr>

<form action="/upload" enctype="multipart/form-data" method="post" >

    上传文件：<input type="file" name="file1"><br/>
    <input type="submit" value="提交"/>

</form>

</body>
</html>
