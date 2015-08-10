<!--#include file="config.asp"--><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>测试页面</title>
<link href="css.css" rel="stylesheet" type="text/css" />
</head>
<body>
<div id="main">
<div class="box">
<a href="test.asp">测试广播</a>&nbsp;<a href="testpic.asp">测试图片</a><br/>
<%
If request.Form("act") = "广播" Then
addObj param,"content",request.Form("content")
addObj param,"clientip",getIP()
addObj param,"jing",""
addObj param,"wei",""
Call postOne(param)
Response.Write "微博id："&obj.data.id & "<br/>"
Response.Write "时间："&obj.data.time & "<br/>"
Else

Call userInfo(param)
Response.Write "昵称:"&obj.data.nick & "<br/>"
Response.Write "帐号:"&obj.data.name & "<br/>"
Response.Write "广播:"&obj.data.tweetnum & "<br/>"
Response.Write "听众:"&obj.data.fansnum & "<br/>"
Response.Write "收听:"&obj.data.idolnum & "<br/>"
End If
%>
<form method="post" action="">
	<textarea name="content" rows="5" cols="80"></textarea><br />
	<input type="submit" name="act" value="广播">
</form>
</div>
<div class="box">
asp-t.qq.com-sdk by <a href="http://t.qq.com/mcodes">@mcodes</a>
</div>
</body>
</html>