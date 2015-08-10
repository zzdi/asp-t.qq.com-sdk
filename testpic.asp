<!--#include file="config.asp"-->
<!--#include file="upload_vbs.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>测试图片</title>
<link href="css.css" rel="stylesheet" type="text/css" />
</head>
<body>
<div id="main">
<div class="box">
<a href="test.asp">测试广播</a>&nbsp;<a href="testpic.asp">测试图片</a><br/>
<%
If request("act") = "action" Then
formPath="temp/"
Dim upload,file,formName,formPath,iCount,filename,fileExt,i
Set upload=new upload_5xSoft
	For Each formName In upload.file
		Set file=upload.file(formName)
		If file.filesize>0 Then
			fileExt=lcase(right(file.filename,4))
			If fileEXT<>".jpg" And fileEXT<>".gif" And fileEXT<>".png" Then
				response.write "<font size=2>允许图片格式(jpg,gif,png)[<a href=""#"" onclick=history.go(-1)>请重新上传</a>]</font>"
				response.End
			Else
				filename=session.sessionid&fileEXT
				file.SaveAs Server.mappath(formpath&filename)
			End If
		End If
		Set file=Nothing
	Next
	content = upload.Form("content")
	Set upload=Nothing

pic = Server.mappath(formpath&filename)
addObj param,"content",content
addObj param,"clientip",getIP()
addObj param,"jing",""
addObj param,"wei",""
addObj param,"pic",pic
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
<form action="?act=action" method="post" enctype="multipart/form-data" name="sub_upload">
<input name="file" type="file" size="25"><br />
<textarea name="content" rows="5" cols="80"></textarea><br />
<input type="submit" name="submit" value="广播">
</div>
<div class="box">
asp-t.qq.com-sdk by <a href="http://t.qq.com/mcodes">@mcodes</a>
</div>
</body>
</html>