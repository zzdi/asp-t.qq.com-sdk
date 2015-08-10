<!--#include file="config.asp"--><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>asp-t.qq.com-sdk</title>
<meta name="description" content="asp-t.qq.com-sdk" />
<meta name="keywords" content="asp-t.qq.com-sdk" />
<link href="css.css" rel="stylesheet" type="text/css" />
</head>
<body>
<div id="main">
<div class="box">
<%
Call get_oauth_http(get_oauth_url(request_token_url,"GET",param))
aurl = authorize_url&"?"&"&oauth_token="&Session("oauth_token")
%>
<a href="<%=aurl%>">使用微博登陆</a>
</div>
asp-t.qq.com-sdk
<div class="box">
	asp-t.qq.com-sdk by <a href="http://t.qq.com/mcodes">@mcodes</a>
</div>
</div>
</body>
</html>