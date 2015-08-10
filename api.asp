<%
'用户信息 => 有name查询name信息,无name查询登录者信息
Sub userInfo(param)
If isObj(param,"name") Then
	Set obj = toObject(gethttp("http://open.t.qq.com/api/user/other_info",param))
Else
	Set obj = toObject(gethttp("http://open.t.qq.com/api/user/info",param))
End If
End Sub

'时间线 => 有name：查询name时间线；无name有f:(1:登录者时间线;2:发布的广播;3:提及时间线;4:关注者时间线)
Sub timeline(param)
If noObj(param,"pageflag") Then addObj param,"pageflag","0"
If noObj(param,"pagetime") Then addObj param,"pagetime","0"
If noObj(param,"reqnum") Then addObj param,"reqnum","20"
If isObj(param,"name") Then
	Set obj = toObject(gethttp("http://open.t.qq.com/api/statuses/user_timeline",param))
ElseIf isObj(param,"f") Then
	If param.f = 1 Then
		Set obj = toObject(gethttp("http://open.t.qq.com/api/statuses/home_timeline",param))
	ElseIf param.f = 2 Then
		Set obj = toObject(gethttp("http://open.t.qq.com/api/statuses/broadcast_timeline",param))
	ElseIf param.f = 3 Then
		Set obj = toObject(gethttp("http://open.t.qq.com/api/statuses/mentions_timeline",param))
	ElseIf param.f = 4 Then
		Set obj = toObject(gethttp("http://open.t.qq.com/api/statuses/special_timeline",param))
	Else
		Set obj = toObject("{msg:'参数f错误'}")
	End If
End If
End Sub

'广播大厅
Sub public_timeline(param)
If noObj(param,"pos") Then addObj param,"pos","0"
If noObj(param,"reqnum") Then addObj param,"reqnum","20"
Set obj = toObject(gethttp("http://open.t.qq.com/api/statuses/public_timeline",param))
End Sub

'话题查询
Sub ht_timeline(param)
If noObj(param,"pageflag") Then addObj param,"pageflag","4"
If noObj(param,"pageinfo") Then addObj param,"pageinfo",""
If noObj(param,"reqnum") Then addObj param,"reqnum","20"
If noObj(param,"httext") Then
addObj param,"httext",strUrlEnCode("山寨站长")
Else
param.httext = strUrlEnCode(param.httext)
End If
Set obj = toObject(gethttp("http://open.t.qq.com/api/statuses/ht_timeline",param))
End Sub

'发送一条广播
Sub postOne(param)
If noObj(param,"jing") Then addObj param,"jing",""
If noObj(param,"wei") Then addObj param,"wei",""
If noObj(param,"clientip") Then addObj param,"clientip",getIP()
If isObj(param,"pic") Then
	Set obj = toObject(postmulti("http://open.t.qq.com/api/t/add_pic",param))
Else
	Set obj = toObject(posthttp("http://open.t.qq.com/api/t/add",param))
End If
End Sub

'获取一条微博信息
Sub getOne(param)
Set obj = toObject(gethttp("http://open.t.qq.com/api/t/show",param))
End Sub

'删除一条微博信息
Sub delOne(param)
Set obj = toObject(gethttp("http://open.t.qq.com/api/t/del",param))
End Sub

'f:(1:转播,2:回复,3:点评)
Sub reOne(param)
param.content = strUrlEnCode(param.content)
If noObj(param,"jing") Then addObj param,"jing",""
If noObj(param,"wei") Then addObj param,"wei",""
If noObj(param,"clientip") Then addObj param,"clientip",getIP()
If isObj(param,"reid") Then
	If param.f = 1 Then
		Set obj = toObject(posthttp("http://open.t.qq.com/api/t/re_add",param))
	ElseIf param.f = 2 Then
		Set obj = toObject(posthttp("http://open.t.qq.com/api/t/reply",param))
	ElseIf param.f = 3 Then
		Set obj = toObject(posthttp("http://open.t.qq.com/api/t/comment",param))
	Else
		Set obj = toObject("{msg:'参数f错误'}")
	End If
Else
Set obj = toObject("{msg:'缺少reid'}")
End If
End Sub

'听众/偶像 => f:(1:听众;2:偶像;3:特别收听;4:黑名单),有name查询name听众/偶像,无name查询自己
Sub getFans(param)
If noObj(param,"reqnum") Then addObj param,"reqnum","30"
If noObj(param,"startindex") Then addObj param,"startindex","0"
If isObj(param,"f") Then
	If param.f = 1 Then
		If isObj(param,"name") Then
			Set obj = toObject(gethttp("http://open.t.qq.com/api/friends/user_fanslist",param))
		Else
			Set obj = toObject(gethttp("http://open.t.qq.com/api/friends/fanslist",param))
		End If
	ElseIf param.f = 2 Then
		If isObj(param,"name") Then
			Set obj = toObject(gethttp("http://open.t.qq.com/api/friends/user_idollist",param))
		Else
			Set obj = toObject(gethttp("http://open.t.qq.com/api/friends/idollist",param))
		End If
	ElseIf param.f = 3 Then
		If isObj(param,"name") Then
			Set obj = toObject(gethttp("http://open.t.qq.com/api/friends/user_speciallist",param))
		Else
			Set obj = toObject(gethttp("http://open.t.qq.com/api/friends/speciallist",param))
		End If
	ElseIf param.f = 4 Then
		Set obj = toObject(gethttp("http://open.t.qq.com/api/friends/blacklist",param))
	Else
		Set obj = toObject("{msg:'参数f错误'}")
	End If
Else
	Set obj = toObject("{msg:'缺少参数f'}")
End If
End Sub

'收听/取消 => f:(1:收听;2:取消;3:特别收听;4:取消特别收听)
Sub setFans(param)
If isObj(param,"name") Then
	If isobj(param,"f") Then
		If param.f = 1 Then
			Set obj = toObject(posthttp("http://open.t.qq.com/api/friends/add",param))
		ElseIf param.f = 2 Then
			Set obj = toObject(posthttp("http://open.t.qq.com/api/friends/del",param))
		ElseIf param.f = 3 Then
			Set obj = toObject(posthttp("http://open.t.qq.com/api/friends/addspecial",param))
		ElseIf param.f = 4 Then
			Set obj = toObject(posthttp("http://open.t.qq.com/api/friends/delspecial",param))
		Else
			Set obj = toObject("{msg:'参数f错误'}")
		End If
	Else
		Set obj = toObject("{msg:'缺少参数f'}")
	End If
Else
	Set obj = toObject("{msg:'缺少参数name'}")
End If
End Sub

%>