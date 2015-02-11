<%
'by selectersky
'2009.9.16修改makefilename函数 2009.11.09增加delContentPic，findContentPic函数
'2009.07.24 加入getallfiles函数，可以获取所以文件列表 2009.07.27新增downloadFile下载文件函数 2009.09.09修改SaveToFileauto函数
'2009.04.02 删除isExists函数，放入safe_info.asp中 2009.07.10加入LoadFromFileauto函数，SaveToFileauto函数。自动识别编码读取保存文件
function makefilename()
dim ranNum
dim dtNow
dtNow=Now()
randomize
ranNum=int(90*rnd)+10
makefilename=year(dtNow) & right("0" & month(dtNow),2) & right("0" & day(dtNow),2) & right("0" & hour(dtNow),2) & right("0" & minute(dtNow),2) & right("0" & second(dtNow),2) & ranNum
end function 

'获取扩展名
function GetExt(byVal FileName)
	GetExt = right(FileName,len(FileName)-instrrev(FileName,"."))	
	FileName=""	
end function

'采用二进制数据流操作,为了解决编码问题生成静态页面
function SaveToFile(ByVal strBody,ByVal sFile)
	On Error Resume Next     
	sfile=server.MapPath(sfile)	
	if err then
		SaveToFile=1
        Err.Clear
		strBody="":sFile=""
		exit function	
	end if	
	Dim objStream
    Set objStream = Server.CreateObject("ADODB"+"."+"Stream")
    If Err.Number=-2147221005 Then 
        SaveToFile=2
        Err.Clear
		set objStream=nothing
		strBody="":sFile=""	
		exit function			
    End If		
    With objStream
        .Type = 2
        .mode = 3
        .Charset = "utf-8"
		.Open
        .Position = objStream.Size
        .WriteText = trim(strBody)
        .SaveToFile sfile,2	
		.flush
        .Close			
    End With
	if err then
		SaveToFile=3
		err.clear
	else
		SaveToFile=0	
	end if		
    Set objStream = Nothing
	strBody="":sFile=""	
End function

function ReadFromUTF (byVal tpath,byVal charset)    'TempString要读取的模板文件路径; Charset是编码
	On Error Resume Next	
	tpath=server.MapPath(tpath)
		
	if err then
		ReadFromUTF=1
		err.clear
		str="":tpath="":charset="" 
		exit function
	end if  
	dim str,stm
    set stm=server.CreateObject("adodb.stream")	
	if err then
		ReadFromUTF=2
		err.clear
		tpath="":charset=""
		exit function
	end if  
    stm.Type=2 
    stm.mode=3 
    stm.charset=charset
    stm.open
    stm.loadfromfile tpath
	str=stm.readtext
    stm.Close
    set stm=nothing
	if err then
		ReadFromUTF=3
		err.clear
	else
		ReadFromUTF=str	
	end if   
	str="":tpath="":charset=""  
end function 

'在指定的位置创建文件夹
function Createfol(byVal folpath)
	On Error Resume Next	
	if right(folpath,1)<>"/" then folpath=folpath&"/"
	folpath=server.MapPath(folpath)	
	if err then
		Createfol=1
		err.clear
		folpath=""
		exit function
	end if 
	dim fso
	set fso=server.CreateObject("Scripting.FileSystemObject")	
	if err then
		Createfol=2	
		err.clear
		folpath=""
		exit function
	end if	
	if not fso.FolderExists(folpath) then		
		fso.CreateFolder(folpath)
		if err then
			Createfol=4		
			err.clear
		else
			Createfol=0
		end if
	else
		Createfol=3
	end if		
	set fso=nothing
	folpath=""
end function

'删除指定的文件夹
function Delfol(byVal folpath)	
	On Error Resume Next
	folpath=server.MapPath(folpath)
	if err then
		Delfol=1	
		err.clear
		folpath=""
		exit function
	end if	
	dim fsobj
	set fsobj=server.CreateObject("Scripting.FileSystemObject")	
	if  fsobj.FolderExists(folpath) then		
		fsobj.DeleteFolder(folpath)
		if err then
			Delfol=3
			err.clear				
		else
			Delfol=0
		end if	
	else
		Delfol=2
	end if	
	set fsobj=nothing
	folpath=""
end function

'删除指定的文件
function DelFile(byVal filename)
	On Error Resume Next	
	filename=server.MapPath(filename)	
	if err then
		DelFile=1
		err.clear
		filename=""
		exit function
	end if	
	dim delobj	
	set delobj=server.CreateObject("Scripting.FileSystemObject")
	if delobj.FileExists(filename) then
		delobj.DeleteFile(filename)
		if err then
			DelFile=3
			err.clear			
		else
			DelFile=0
		end if	
	else
		delFile=2
	end if
	set delobj=nothing	
	filename=""
end function

'判断指定文件夹是否为空
function isNullFolder(byVal path) 	
	On Error Resume Next
	path=server.mapPath(path)
	if err then
		isNullFolder=true
		err.clear
		path=""
		exit function
	end if
	dim fs,f
	set fs=createobject("scripting.filesystemobject") 	
	if fs.FolderExists(path) then
		set f=fs.GetFolder(path)
		if f.size=0 then
			isNullFolder=true
		else
			isNullFolder=false
		end if
		set f=nothing
	else
		isNullFolder=false
	end if 
	set fs=nothing
	path=""
end function

'==================================================
'函数名：GetHttpPage
'作  用：获取网页源码
'参  数：HttpUrl ------网页地址
'==================================================
Function GetHttpPage(HttpUrl)
   If IsNull(HttpUrl) Or Len(HttpUrl)<18 Or HttpUrl="$False$" Then
      GetHttpPage="$False$"
      Exit Function
   End If
   Dim Http
   Set Http=server.createobject("MSXML2.XMLHTTP")
   Http.open "GET",HttpUrl,False
   Http.Send()
   If Http.Readystate<>4 then
      Set Http=Nothing 
      GetHttpPage="$False$"
      Exit function
   End if
   GetHTTPPage=bytesToBSTR(Http.responseBody,"utf-8")
   Set Http=Nothing
   If Err.number<>0 then
      Err.Clear
   End If
End Function

'==================================================
'函数名：BytesToBstr
'作  用：将获取的源码转换为中文
'参  数：Body ------要转换的变量
'参  数：Cset ------要转换的类型
'==================================================
Function BytesToBstr(Body,Cset)
   Dim Objstream
   Set Objstream = Server.CreateObject("adodb.stream")
   objstream.Type = 1
   objstream.Mode =3
   objstream.Open
   objstream.Write body
   objstream.Position = 0
   objstream.Type = 2
   objstream.Charset = Cset
   BytesToBstr = objstream.ReadText 
   objstream.Close
   set objstream = nothing
End Function

'读取文件 自动识别编码
Function LoadFromFileauto(ByVal File)
dim cset
    Dim objStream
    dim a1,b1,c1,a2,b2,c2
    Dim RText,RTexta,RTextb,RTextc
	Dim csettext,tmpcset
    RText = Array(0, "")
    Set objStream = Server.CreateObject("ADODB.Stream")
    With objStream
        .Type = 2
        .Mode = 3
        .Open
        .charset = "unicode"
        .Position = objStream.Size
        .LoadFromFile File
        RTexta = Array(0, .ReadText)
        a2=len(RTexta(1))
        a1=objStream.Size
        .Close
    End With
    With objStream
        .Type = 2
        .Mode = 3
        .Open
        .Position = objStream.Size
        .charset = "utf-8"
        .LoadFromFile File
        RTextb = Array(0, .ReadText)
        b2=len(RTextb(1))
        b1=objStream.Size
        .Close
    End With
    With objStream
        .Type = 2
        .Mode = 3
        .Open
        .Position = objStream.Size
        .charset = "utf-8"
        .LoadFromFile File
        RTextc = Array(0, .ReadText)
        c2=len(RTextc(1))
        c1=objStream.Size
        .Close
    End With
if a1<b1 then 
if a1<c1 then csettext=RTexta:tmpcset="unicode"
if a1<=c1 then 
if a2<c2 then csettext=RTexta:tmpcset="unicode"
end if
end if
if b1<a1 then 
if b1<c1 then csettext=RTextb:tmpcset="utf-8"
if b1<=c1 then 
if b2<c2 then csettext=RTextb:tmpcset="utf-8"
end if
end if
if c1<a1 then 
if c1<b1 then csettext=RTextc:tmpcset="utf-8"
if c1<=b1 then 
if c2<b2 then csettext=RTextc:tmpcset="utf-8"
end if
end if
session("cset")=tmpcset
LoadFromFileauto = csettext(1)
Set objStream = Nothing
erase csettext:erase RTexta:erase RTextb:erase RTextc:erase RText
End Function

'保存到文件 自动识别编码 必须在LoadFromFileauto下使用
Function SaveToFileauto(strBody,sFile,cset)
   On Error Resume Next     
	sfile=server.MapPath(sfile)	
	if cset="" then cset=session("cset")	
	if err then
		SaveToFileauto=array(1,"文件路径不正确")        
		exit function	
	end if	
	Dim objStream
    Set objStream = Server.CreateObject("ADODB"+"."+"Stream")
    If Err.Number=-2147221005 Then 
        SaveToFileauto=array(2,"创建文件组件对象失败，可能您的服务器不支持此组件")         
		exit function			
    End If		
    With objStream
        .Type = 2
        .mode = 3
        .Charset = cset
		.Open
        .Position = objStream.Size
        .WriteText = trim(strBody)
        .SaveToFile sfile,2	
		.flush
        .Close			
    End With
	if err then
		SaveToFileauto=array(3,"写入到文件失败")   
		err.clear
	else
		SaveToFileauto=array(0,"")	
	end if		
    Set objStream = Nothing
	strBody="":sFile=""	
End Function

'返回指定目录的文件 各文件以|分隔
'path 文件夹路径（这里是绝对路径） issub为1是包含子目录 fsoobj FSO的对象
function getAllFiles(byval path,issub,fsoobj)
dim Root1,f1,fileobj
dim fileList
fileList=""
Set Root1 = fsoobj.GetFolder(path)
for each fileobj in Root1.files	
	if fileList="" then fileList=fileobj.path else fileList=fileList&"|"&fileobj.path	
next
'包含子目录
if issub=1 then
	For Each f1 In Root1.subfolders
		if fileList="" then fileList=getAllFiles(f1.path,1,fsoobj) else fileList=fileList&"|"&getAllFiles(f1.path,1,fsoobj)		
	next
	set f1=nothing
end if
getAllFiles=replace(fileList,"||","|")
fileList=""
set Root1=nothing
set fileobj=nothing
end function

'返回指定目录的子目录 各文件以|分隔
'path 文件夹路径（这里是绝对路径） issub为1是包含子目录 fsoobj FSO
function getAllsubfolders(byval path,issub,fsoobj)
dim Root1,f1
dim fileList
Set Root1 = fsoobj.GetFolder(path)
For Each f1 In Root1.subfolders
		if fileList="" then fileList=f1.path else fileList=fileList&"|"&f1.path
		if issub=1 then fileList=fileList&"|"&getAllsubfolders(f1.path,1,fsoobj)		
next
getAllsubfolders=replace(fileList,"||","|")
fileList=""
set f1=nothing
set Root1=nothing
end function

Function downloadFile(strFile)
Const ForReading = 1
Const TristateTrue = -1
Const FILE_TRANSFER_SIZE = 16384
Dim objFileSystem,objFile,objStream,char,sent,path,FileName,send,s_DownFilePath,s_FileExt,TransferFile
send = 0
path = Server.MapPath(strFile)
Set objFileSystem = Server.CreateObject("Scripting.FileSystemObject")
If Not objFileSystem.fileexists(path) Then
Response.Write ("<h1>错误:</h1>" & strFile & "没有发现!<p>")
Response.End
End If
Set objFile = objFileSystem.GetFile(path)
s_DownFilePath = objFile.Name
s_FileExt = Mid(s_DownFilePath, InStrRev(s_DownFilePath, ".") + 1)
If UCase(s_FileExt) <> "LMB" And UCase(s_FileExt) <> "TXT" And UCase(s_FileExt) <> "BAK" And UCase(s_FileExt)<>"MDB" And UCase(s_FileExt)<>"CSV" And UCase(s_FileExt)<>"XLS" Then 
response.Write("只能下载扩展名为lmb或txt或BAK或MDB或CSV或xls的文件")
response.End()
end if
Set objStream = objFile.OpenAsTextStream(ForReading, TristateTrue)
Response.AddHeader "content-type", "application/server"
Response.AddHeader "Content-Disposition", "attachment;filename=" & s_DownFilePath
Response.AddHeader "content-length", objFile.Size
Do While Not objStream.AtEndOfStream
char = objStream.Read(1)
Response.BinaryWrite (char)
sent = sent + 1
If (sent Mod FILE_TRANSFER_SIZE) = 0 Then
Response.Flush
If Not Response.IsClientConnected Then Exit Do
End If
Loop
Response.Flush
If Not Response.IsClientConnected Then TransferFile = False
objStream.Close
Set objStream = Nothing
Set objFileSystem = Nothing
End Function

'查找内容里面的图片 返回一维数组 arr(0) 表示是否有找到 值0为没有找到 值为1有找到
function findContentPic(byval stbody)
redim tmparr(0)
dim matchs,i
dim findnum
tmparr(0)=0
findnum=0
dim regEx
set regEx=new regExp
regEx.ignoreCase=true
regEx.Global=True
regEx.pattern="src=.+?( |/>|>|/&gt;|&gt;)"
set matchs=regEx.execute(stbody)
for each i in matchs
	tmparr(0)=1
	redim preserve tmparr(findnum+1)
	tmparr(findnum+1)=replace(replace(replace(replace(replace(replace(replace(replace(i.value,"src=",""),"&quot;",""),"""","")," ",""),"/>",""),">",""),"/&gt;",""),"&gt;","")
	findnum=findnum+1
next
set regEx=nothing
set matchs=nothing
findContentPic=tmparr
stbody="":erase tmparr
end function

'删除内容里面的图片
function delContentPic(byval stbody)
if chkNull(stbody,1) then 
	delContentPic=array(0,0)
	exit function
end if
dim tmparr,i,delnum
delnum=0
tmparr=findContentPic(stbody)
if tmparr(0)=1 then
	for i=1 to ubound(tmparr)
		if delfile(tmparr(i)) then delnum=delnum+1		
	next
end if
delContentPic=array(i,delnum)
stbody="":erase tmparr
end function
%>