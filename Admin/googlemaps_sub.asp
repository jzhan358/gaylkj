<%
'配置信息
public WebUrlAddr,GsUrlAddr
WebUrlAddr=cityurl			'网站网址
GsUrlAddr="/sitemap.xml"	'网站地图文件存在位置


if chkNull(WebUrlAddr,5) then alert "网址丢失","",1
if right(WebUrlAddr,1)="/" then WebUrlAddr=left(WebUrlAddr,Len(WebUrlAddr)-1)


Sub ShowSubFolders(objFolder)
Set colFolders = objFolder.SubFolders
For Each objSubFolder In colFolders
if folderpermission(objSubFolder.Path) then
str = str & getfilelink(objSubFolder.Path,objSubFolder.dateLastModified) & vbcrlf
Set colFiles = objSubFolder.Files
For Each objFile In colFiles
str = str & getfilelink(objFile.Path,objFile.dateLastModified) & vbcrlf
Next
ShowSubFolders(objSubFolder)
end if
Next
End Sub


Function getfilelink(file,datafile)
file=replace(file,"\","/")
filedate=NowToDate(datafile)
getfilelink = "<url><loc>"&server.htmlencode(file)&"</loc><lastmod>"&filedate&"</lastmod><changefreq>daily</changefreq><priority>1.0</priority></url>"
'Response.Flush
End Function


Function Folderpermission(pathName)
Dim PathExclusion
'需要过滤的目录(不列在SiteMap里面)
'需要列在SiteMap里面的目录
PathExclusion="|\html|"
if lodo_web_htmlfoldername<>"" Then PathExclusion=PathExclusion&"/"&lodo_web_htmlfoldername&"|"
Folderpermission =False
if instr(LCase(PathExclusion),LCase(pathName))>0 Then Folderpermission = True
End Function

Function GetFileExt(filename)
Dim Tmp_fileExt
Tmp_fileExt=filename
Tmp_fileExt = right(Tmp_fileExt, len(Tmp_fileExt) - instrrev(Tmp_fileExt, "."))
if instr(Tmp_fileExt,"?")>0 Then Tmp_fileExt=left(Tmp_fileExt,instr(Tmp_fileExt,"?")-1)
GetFileExt=Tmp_fileExt
End Function

Function FileExtensionIsBad(sFileName)
Dim sFileExtension, bFileExtensionIsValid, sFileExt,sFile,Fileensions
'modify for your file extension (http://www.googleguide.com/file_type.html)
Extensions=Array("html","htm","asp")'设置列表的文件名,扩展名不在其中的话SiteMap则不会收录该扩展名的文件
Fileensions=Array("")'设置不需要加入的文件名

bFileExtensionIsValid = true 'assume extension is bad
if len(trim(sFileName)) = 0 then Exit Function

sFileExtension =GetFileExt(sFileName)

for each sFile in Fileensions
if ucase(sFile)=ucase(sFileName) then Exit Function
next

for each sFileExt in extensions
if ucase(sFileExt) = ucase(sFileExtension) then
bFileExtensionIsValid = false
exit for
end if
next
FileExtensionIsBad=bFileExtensionIsValid
End Function

Function ShowDatecreated(filespec)
Dim f,ffso
if instr(filespec,":")<=0 Then filespec=Server.MapPath(filespec)
Set ffso = Server.CreateObject("scripting.filesystemobject")
If ffso.fileexists(filespec) Then
Set f = ffso.GetFile(filespec)
ShowDatecreated = f.DateLastModified
Else
ShowDatecreated = -1
End If
Set ffso = Nothing
End Function

Function GetFileSizeStyle(FileName)
Dim Bsize,Filesize
If InStr(FileName, ":\") > 0 Then Bsize = GetFileSize(FileName) Else Bsize = FileName
If Bsize >= 0 Then
Filesize = Bsize / 1024
Filesize = FormatNumber(Filesize, 2)
If Filesize < 1024 And Filesize > 1 Then
GetFileSizeStyle = "<font color=red>" & Filesize & "</font>&nbsp;KB"
ElseIf Filesize >= 1024 Then
GetFileSizeStyle = "<font color=red>" & FormatNumber(Filesize / 1024, 2) & "</font>&nbsp;MB"
Else
GetFileSizeStyle = "<font color=red>" & Bsize & "</font>&nbsp;B"
End If
Else
GetFileSizeStyle = -1
End If
End Function

Function GetFileSize(FileName)
Dim f,fso,Filesize
Set fso = Server.CreateObject("scripting.filesystemobject")
If fso.fileexists(FileName) Then
Set f = fso.GetFile(FileName)
GetFileSize = f.Size
Set f = Nothing
Else
GetFileSize = -1
End If
Set fso = Nothing
End Function

Function NowToDate(TempNow)
If IsDate(TempNow) Then NowToDate = Year(TempNow) & "-" & Right("0" & Month(TempNow), 2) & "-" & Right("0" & Day(TempNow), 2) Else NowToDate = ""
End Function
%>
