<!--#include file="multi.asp"-->
<%
'授权过程
Sub get_oauth_http(oauthUrl)
	Dim oauth_http
	Set oauth_http=Server.CreateObject("MSXML2.XMLHTTP")
	oauth_http.Open "GET",oauthUrl,False,"",""
	oauth_http.Send
	If oauth_http.Status = "200" Then
		rs = Split(Replace(oauth_http.responseText,"=","&"),"&")
		Session("oauth_token") = rs(1)
		Session("oauth_token_Secret") = rs(3)
		Session("name") = rs(5)
	Else
		'Response.Write oauth_http.Status & "<br />" & oauth_http.responseText
	End If
	'Response.Write oauth_http.Status & "<br />" & oauth_http.responseText
	Set oauth_http=nothing
End Sub

'Get函数
Function gethttp(gethttp_url,param)
	Dim http_get
	gethttp_url = get_oauth_url(gethttp_url,"GET",param)
	Set http_get=Server.CreateObject("MSXML2.XMLHTTP")
	http_get.Open "GET",gethttp_url,False,"",""
	http_get.Send
	gethttp = http_get.responseText
	'Response.Write http_get.Status & "<br />" & http_get.responseText& "<br /><br />"
	Set http_get=nothing
End Function

'Post函数
Function posthttp(posthttp_url,param)
	Dim http_post
	postData = get_oauth_url(posthttp_url,"POST",param)
	Set http_post=Server.CreateObject("MSXML2.XMLHTTP")
	http_post.Open "POST",posthttp_url,False
	http_post.Send(postData)
	posthttp = http_post.responseText
	'Response.Write http_post.Status & "<br />" & posthttp& "<br /><br />"
	Set http_post=Nothing
End Function

'multi
Function postmulti(posthttp_url,param)
	Dim http_post
	pic = param.pic
	content = param.content
	JSON = get_oauth_url(posthttp_url,"POST",param)
	Set param = toObject(JSON)
	Set param2 = toObject2(param)
	Dim UploadData
	Set UploadData = New XMLUploadImpl
	UploadData.Charset = "utf-8"
	For Each item In param2
	UploadData.AddForm item.name, item.value
	Next
	UploadData.AddForm "content", content
	UploadData.AddFile "pic", "pic.jpg", "image/jpg", pic
	postmulti = UploadData.Upload(posthttp_url)
	'response.write postmulti
	Set UploadData = Nothing
End Function

'生成签名
Function get_oauth_url(oauth_url,httptype,param)
	If isObj(param,"pic") Then ismulti=1:delObj param,"pic"
	delObj param,"f"
	If isObj(param,"content") Then param.content = strUrlEnCode(param.content)
	addObj param,"oauth_consumer_key",App_Key
	addObj param,"oauth_nonce",makePassword(12)
	addObj param,"oauth_signature_method","HMAC-SHA1"
	addObj param,"oauth_timestamp",DateDiff("s","01/01/1970 08:00:00",Now())
	addObj param,"oauth_version","1.0"
	If (oauth_url<>request_token_url and oauth_url<>access_token_url) Then addObj param,"format",format
	If oauth_url=request_token_url Then addObj param,"oauth_callback",callback
	If oauth_url<>request_token_url Then
		addObj param,"oauth_token",Session("oauth_token")
		If oauth_url=access_token_url Then addObj param,"oauth_verifier",Request.QueryString("oauth_verifier")
		oauth_token_Secret = Session("oauth_token_Secret")
	End If
	Base_Par = toStr(param)
	BASE_STRING = httptype & "&" & strUrlEnCode(oauth_url) & "&" & strUrlEnCode(Base_Par)
	oauth_signature = b64_hmac_sha1(App_Secret&"&"&oauth_token_Secret,BASE_STRING)
	If httptype = "GET" Then
		get_oauth_url = oauth_url&"?"&Base_Par&"&oauth_signature=" & strUrlEnCode(oauth_signature)
	ElseIf ismulti=1 Then
		addObj param,"oauth_signature",oauth_signature
		delObj param,"content"
		get_oauth_url = toJSON(param)
	Else
		get_oauth_url = Base_Par&"&oauth_signature=" & strUrlEnCode(oauth_signature)
	End If
End Function
%>