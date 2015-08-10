# asp-t.qq.com-sdk
## 1、程序版本: V1.1
包含基本get与post以及multipart/form-data方式post数据，能满足asp调用腾讯微博api的基本需求。

## 2、文件说明
配置文件
config.asp 设置App_Key,App_Secret,callback和一些常用asp函数
核心文件
function_js.asp 一些服务端的js处理函数,toJSON()是为了配合签名排序拼接用的,对其他object转换成json并不适用
api.asp 包含是常用api接口,可自行添加
oauth.asp 包含get/post以及multipost。调试的时候可以把注释打开。
multi.asp multipart/form-data方式post数据,发布图片广播时用到
示例文件
test.asp 发布广播示例
testpic.asp 发布带图片的广播示例
附带文件
upload_vbs.asp 无组上传,图片广播示例中用到
temp目录 图片广播示例中上传图片临时储存文件夹
## 3、示例：
* 生成授权地址
``` 
<!--#include file="config.asp"-->
<%
Call get_oauth_http(get_oauth_url(request_token_url,"GET",param))
aurl = authorize_url&"?"&"&oauth_token="&Session("oauth_token")
%>
``` 
* 回调处理
``` 
<!--#include file="config.asp"-->
<%
Call get_oauth_http(get_oauth_url(access_token_url,"GET",param))
%>
``` 
* get获取内容在api.asp中,获取用户@mcodes信息
``` 
<!--#include file="config.asp"-->
<%
addObj param,"name","mcodes"
Call userInfo(param)
Response.Write "昵称:"&obj.data.nick & "<br/>"
Response.Write "帐号:"&obj.data.name & "<br/>"
Response.Write "广播:"&obj.data.tweetnum & "<br/>"
Response.Write "听众:"&obj.data.fansnum & "<br/>"
Response.Write "收听:"&obj.data.idolnum
%>
``` 
* post一条微博信息
``` 
<!--#include file="config.asp"-->
<%
addObj param,"content","测试发送微博。asp-sdk_by_mcodes"
addObj param,"clientip",getIP()
addObj param,"jing",""
addObj param,"wei",""
Call postOne(param)
Response.Write "微博id："&obj.data.id & "<br/>"
Response.Write "时间："&obj.data.time
%>
```
* post一条包含图片的微博信息
``` 
<!--#include file="config.asp"-->
<%
addObj param,"content","测试发送带图片的微博。asp-sdk_by_mcodes"
addObj param,"clientip",getIP()
addObj param,"jing",""
addObj param,"wei",""
addObj param,"pic",Server.mappath("temp/pic.jpg")
Call postOne(param)
Response.Write "微博id："&obj.data.id & "<br/>"
Response.Write "时间："&obj.data.time
%>
```



								2011-5-24
								-- by @mcodes  http://t.qq.com/mcodes
