<%
Public Const adTypeBinary = 1
Public Const adTypeText = 2
Public Const adLongVarBinary = 205

'字节数组转指定字符集的字符串
Public Function BytesToString(vtData, ByVal strCharset)
    Dim objFile
    Set objFile = Server.CreateObject("ADODB.Stream")
    objFile.Type = adTypeBinary
    objFile.Open
    If VarType(vtData) = vbString Then
        objFile.Write BinaryToBytes(vtData)
    Else
        objFile.Write vtData
    End If
    objFile.Position = 0
    objFile.Type = adTypeText
    objFile.Charset = strCharset
    BytesToString = objFile.ReadText(-1)
    objFile.Close
    Set objFile = Nothing
End Function

'字节字符串转字节数组，即经过MidB/LeftB/RightB/ChrB等处理过的字符串
Public Function BinaryToBytes(vtData)
    Dim rs
    Dim lSize
    lSize = LenB(vtData)
    Set rs = Server.CreateObject("ADODB.RecordSet")
    rs.Fields.Append "Content", adLongVarBinary, lSize
    rs.Open
    rs.AddNew
    rs("Content").AppendChunk vtData
    rs.Update
    BinaryToBytes = rs("Content").GetChunk(lSize)
    rs.Close
    Set rs = Nothing
End Function

'指定字符集的字符串转字节数组
Public Function StringToBytes(ByVal strData, ByVal strCharset)
    Dim objFile
    Set objFile = Server.CreateObject("ADODB.Stream")
    objFile.Type = adTypeText
    objFile.Charset = strCharset
    objFile.Open
    objFile.WriteText strData
    objFile.Position = 0
    objFile.Type = adTypeBinary
    If UCase(strCharset) = "UNICODE" Then
        objFile.Position = 2 'delete UNICODE BOM
    ElseIf UCase(strCharset) = "UTF-8" Then
        objFile.Position = 3 'delete UTF-8 BOM
    End If
    StringToBytes = objFile.Read(-1)
    objFile.Close
    Set objFile = Nothing
End Function

'获取文件内容的字节数组
Public Function GetFileBinary(ByVal strPath)
    Dim objFile
    Set objFile = Server.CreateObject("ADODB.Stream")
    objFile.Type = adTypeBinary
    objFile.Open
    objFile.LoadFromFile strPath
    GetFileBinary = objFile.Read(-1)
    objFile.Close
    Set objFile = Nothing
End Function

'XML Upload Class
Class XMLUploadImpl
Private xmlHttp
Private objTemp
Private strCharset, strBoundary

Private Sub Class_Initialize()
    Set xmlHttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
    Set objTemp = Server.CreateObject("ADODB.Stream")
    objTemp.Type = adTypeBinary
    objTemp.Open
    strCharset = "utf-8"
    strBoundary = GetBoundary()
End Sub

Private Sub Class_Terminate()
    objTemp.Close
    Set objTemp = Nothing
    Set xmlHttp = Nothing
End Sub

'获取自定义的表单数据分界线
Private Function GetBoundary()
    Dim ret(24)
    Dim table
    Dim i
    table = "ABCDEFGHIJKLMNOPQRSTUVWXZYabcdefghijklmnopqrstuvwxzy0123456789"
    Randomize
    For i = 0 To UBound(ret)
        ret(i) = Mid(table, Int(Rnd() * Len(table) + 1), 1)
    Next
    GetBoundary = "__NextPart__" & Join(ret, Empty)
End Function

'设置上传使用的字符集
Public Property Let Charset(ByVal strValue)
    strCharset = strValue
End Property

'添加文本域的名称和值
Public Sub AddForm(ByVal strName, ByVal strValue)
    Dim tmp
    tmp = "\r\n--$1\r\nContent-Disposition: form-data; name=""$2""\r\n\r\n$3"
    tmp = Replace(tmp, "\r\n", vbCrLf)
    tmp = Replace(tmp, "$1", strBoundary)
    tmp = Replace(tmp, "$2", strName)
    tmp = Replace(tmp, "$3", strValue)
    objTemp.Write StringToBytes(tmp, strCharset)
End Sub

'设置文件域的名称/文件名称/文件MIME类型/文件路径或文件字节数组
Public Sub AddFile(ByVal strName, ByVal strFileName, ByVal strFileType, vtValue)
    Dim tmp
    tmp = "\r\n--$1\r\nContent-Disposition: form-data; name=""$2""; filename=""$3""\r\nContent-Type: $4\r\n\r\n"
    tmp = Replace(tmp, "\r\n", vbCrLf)
    tmp = Replace(tmp, "$1", strBoundary)
    tmp = Replace(tmp, "$2", strName)
    tmp = Replace(tmp, "$3", strFileName)
    tmp = Replace(tmp, "$4", strFileType)
    objTemp.Write StringToBytes(tmp, strCharset)
    If VarType(vtValue) = (vbByte Or vbArray) Then
        objTemp.Write vtValue
    Else
        objTemp.Write GetFileBinary(vtValue)
    End If
End Sub

'设置multipart/form-data结束标记
Private Sub AddEnd()
    Dim tmp
    'tmp = Replace("\r\n--$1--\r\n", "$1", strBoundary)
        tmp = "\r\n--$1--\r\n"
        tmp = Replace(tmp, "\r\n", vbCrLf)
        tmp = Replace(tmp, "$1", strBoundary)
    objTemp.Write StringToBytes(tmp, strCharset)
    objTemp.Position = 2
End Sub

'上传到指定的URL，并返回服务器应答
Public Function Upload(ByVal strURL)
    Call AddEnd
    xmlHttp.Open "POST", strURL, False
    xmlHttp.setRequestHeader "Content-Type", "multipart/form-data,boundary="&strBoundary
    xmlHttp.setRequestHeader "Content-Length", objTemp.size
    xmlHttp.Send objTemp
        If VarType(xmlHttp.responseBody) = (vbByte Or vbArray) Then
            Upload = BytesToString(xmlHttp.responseBody, strCharset)
        End If
End Function
End Class
%>