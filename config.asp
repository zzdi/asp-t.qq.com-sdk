<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'On Error Resume Next
Set param = toObject("{}")
Dim obj
Const App_Key = "****************" '填入你的Key
Const App_Secret = "****************" '填入你的密钥
callback = strUrlEnCode("http://sdk.zzdi.cn/callback.asp") '授权成功后的回调地址
Const format = "json" '返回数据格式
Const request_token_url = "https://open.t.qq.com/cgi-bin/request_token"
Const authorize_url = "https://open.t.qq.com/cgi-bin/authorize"
Const access_token_url = "https://open.t.qq.com/cgi-bin/access_token"

'生成随机串
Function makePassword(byVal maxLen)
	Dim strNewPass
	Dim whatsNext, upper, lower, intCounter
	Randomize
	For intCounter = 1 To maxLen
	whatsNext = Int((1 - 0 + 1) * Rnd + 0)
	If whatsNext = 0 Then
	upper = 101
	lower = 97
	Else
	upper = 57
	lower = 48
	End If
	strNewPass = strNewPass & Chr(Int((upper - lower + 1) * Rnd + lower))
	Next
	makePassword = strNewPass
End function

'UrlEnCode
Function strUrlEnCode(byVal strUrl)
strUrlEnCode = Server.URLEncode(strUrl)
strUrlEnCode = Replace(strUrlEnCode,"%5F","_")
strUrlEnCode = Replace(strUrlEnCode,"%2E",".")
strUrlEnCode = Replace(strUrlEnCode,"%2D","-")
strUrlEnCode = Replace(strUrlEnCode,"+","%20")
End Function

'IP
Function getIP()
 dim strIPAddr
    if request.servervariables("HTTP_X_FORWARDED_FOR") = "" or instr(request.servervariables("HTTP_X_FORWARDED_FOR"), "unknown") > 0 then
        strIPAddr = request.servervariables("REMOTE_ADDR")
    elseif instr(request.servervariables("HTTP_X_FORWARDED_FOR"), ",") > 0 then
        strIPAddr = mid(request.servervariables("HTTP_X_FORWARDED_FOR"), 1, instr(request.servervariables("HTTP_X_FORWARDED_FOR"), ",")-1)
    elseif instr(request.servervariables("HTTP_X_FORWARDED_FOR"), ";") > 0 then
        strIPAddr = mid(request.servervariables("HTTP_X_FORWARDED_FOR"), 1, instr(request.servervariables("HTTP_X_FORWARDED_FOR"), ";")-1)
    else
        strIPAddr = request.servervariables("HTTP_X_FORWARDED_FOR")
    end if
    getIP = Trim(mid(strIPAddr, 1, 30))
End Function
%>
<!--#include file="function_js.asp"-->
<!--#include file="oauth.asp"-->
<!--#include file="api.asp"-->