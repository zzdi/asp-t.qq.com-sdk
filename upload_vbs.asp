<SCRIPT RUNAT=SERVER LANGUAGE=VBSCRIPT>
dim upfile_5xSoft_Stream
Class upload_5xSoft
dim Form,File,Version
Private Sub Class_Initialize
dim iStart,iFileNameStart,iFileNameEnd,iEnd,vbEnter,iFormStart,iFormEnd,theFile
dim strDiv,mFormName,mFormValue,mFileName,mFileSize,mFilePath,iDivLen,mStr,FormName
Version= ""
If Request.TotalBytes <1 Then Exit Sub
Set Form=CreateObject("Scripting.Dictionary")
Set File=CreateObject("Scripting.Dictionary")
Set upfile_5xSoft_Stream=CreateObject( "Adodb.Stream")
upfile_5xSoft_Stream.mode=3
upfile_5xSoft_Stream.type=1
upfile_5xSoft_Stream.open
upfile_5xSoft_Stream.write Request.BinaryRead(Request.TotalBytes)

vbEnter=Chr(13)&Chr(10)
iDivLen=inString(1,vbEnter)+1
strDiv=subString(1,iDivLen)
iFormStart=iDivLen
iFormEnd=inString(iformStart,strDiv)-1
while iFormStart < iFormEnd
  iStart=inString(iFormStart, "name=""")
  iEnd=inString(iStart+6, """")
  mFormName=subString(iStart+6,iEnd-iStart-6)
  iFileNameStart=inString(iEnd+1, "filename=""")
  if iFileNameStart> 0 and iFileNameStart <iFormEnd then
  iFileNameEnd=inString(iFileNameStart+10, """")
  mFileName=subString(iFileNameStart+10,iFileNameEnd-iFileNameStart-10)
  iStart=inString(iFileNameEnd+1,vbEnter&vbEnter)
  iEnd=inString(iStart+4,vbEnter&strDiv)
  if iEnd> iStart then
mFileSize=iEnd-iStart-4
  else
mFileSize=0
  end if
  set theFile=new FileInfo
  theFile.FileName=getFileName(mFileName)
  theFile.FilePath=getFilePath(mFileName)
  theFile.FileSize=mFileSize
  theFile.FileStart=iStart+4
  theFile.FormName=FormName
  file.add Lcase(mFormName),theFile
  else
  iStart=inString(iEnd+1,vbEnter&vbEnter)
  iEnd=inString(iStart+4,vbEnter&strDiv)

  if iEnd> iStart then
mFormValue=subString(iStart+4,iEnd-iStart-4)
  else
mFormValue= ""
  end if
  form.Add Lcase(mFormName),mFormValue
  end if

  iFormStart=iformEnd+iDivLen
  iFormEnd=inString(iformStart,strDiv)-1
wend
End Sub

Private Function subString(theStart,theLen)
  dim i,c,stemp
  upfile_5xSoft_Stream.Position=theStart-1
  if theLen> 0 then subString=BytesToBstr(upfile_5xSoft_Stream.Read(theLen), "utf-8") else subString= ""
End function

Function BytesToBstr(strBody,CodeBase)
      dim objStream
      set objStream = Server.CreateObject("Adodb.Stream")
      objStream.Type = 1
      objStream.Mode =3
      objStream.Open
      objStream.Write strBody
      objStream.Position = 0
      objStream.Type = 2
      objStream.Charset = CodeBase
      BytesToBstr = objStream.ReadText
      objStream.Close
      set objStream = nothing
End Function

Private Function inString(theStart,varStr)
  dim i,j,bt,theLen,str
  InString=0
  Str=toByte(varStr)
  theLen=LenB(Str)
  for i=theStart to upfile_5xSoft_Stream.Size-theLen
  if i> upfile_5xSoft_Stream.size then exit Function
  upfile_5xSoft_Stream.Position=i-1
  if AscB(upfile_5xSoft_Stream.Read(1))=AscB(midB(Str,1)) then
    InString=i
    for j=2 to theLen
    if upfile_5xSoft_Stream.EOS then
      inString=0
      Exit for
    end if
    if AscB(upfile_5xSoft_Stream.Read(1)) <> AscB(MidB(Str,j,1)) then
      InString=0
      Exit For
    end if
    next
    if InString <> 0 then Exit Function
  end if
  next
End Function

Private Sub Class_Terminate
  form.RemoveAll
  file.RemoveAll
  set form=nothing
  set file=nothing
  upfile_5xSoft_Stream.close
  set upfile_5xSoft_Stream=nothing
End Sub


  Private function GetFilePath(FullPath)
  If FullPath <> "" Then
  GetFilePath = left(FullPath,InStrRev(FullPath, "\"))
  Else
  GetFilePath = ""
  End If
  End   function

  Private function GetFileName(FullPath)
  If FullPath <> "" Then
  GetFileName = mid(FullPath,InStrRev(FullPath, "\")+1)
  Else
  GetFileName = ""
  End If
  End   function

  Private function toByte(Str)
  dim i,iCode,c,iLow,iHigh
  toByte= ""
  For i=1 To Len(Str)
  c=mid(Str,i,1)
  iCode =Asc(c)
  If iCode <0 Then iCode = iCode + 65535
  If iCode> 255 Then
    iLow = Left(Hex(Asc(c)),2)
    iHigh =Right(Hex(Asc(c)),2)
    toByte = toByte & chrB( "&H"&iLow) & chrB( "&H"&iHigh)
  Else
    toByte = toByte & chrB(AscB(c))
  End If
  Next
  End function

End Class


Class FileInfo
  dim FormName,FileName,FilePath,FileSize,FileStart
  Private Sub Class_Initialize
    FileName = ""
    FilePath = ""
    FileSize = 0
    FileStart= 0
    FormName = ""
  End Sub

  Public function SaveAs(FullPath)
    dim dr,ErrorChar,i
    SaveAs=1
    if trim(fullpath)= "" or FileSize=0 or FileStart=0 or FileName= "" then exit function
    if FileStart=0 or right(fullpath,1)= "/" then exit function
    set dr=CreateObject("Adodb.Stream")
    dr.Mode=3
    dr.Type=1
    dr.Open
    upfile_5xSoft_Stream.position=FileStart-1
    upfile_5xSoft_Stream.copyto dr,FileSize
    dr.SaveToFile FullPath,2
    dr.Close
    set dr=nothing
    SaveAs=0
  end function
End Class
</SCRIPT>