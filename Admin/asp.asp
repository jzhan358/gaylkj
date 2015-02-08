
<%@ Language="VBScript" %>
<%' Option Explicit %>
<%
'####################################
'#									#
'#		 阿江ASP探针 V1.70			#
'#									#
'#  阿江守候 http://www.ajiang.net  #
'#	 电子邮件 info@ajiang.net		#
'#									#
'#    转载本程序时请保留这些信息    #
'#								    #
'####################################


'不使用输出缓冲区，直接将运行结果显示在客户端
Response.Buffer = False

'声明待检测数组
Dim ObjTotest(26,4)

ObjTotest(0,0) = "MSWC.AdRotator"
ObjTotest(1,0) = "MSWC.BrowserType"
ObjTotest(2,0) = "MSWC.NextLink"
ObjTotest(3,0) = "MSWC.Tools"
ObjTotest(4,0) = "MSWC.Status"
ObjTotest(5,0) = "MSWC.Counters"
ObjTotest(6,0) = "IISSample.ContentRotator"
ObjTotest(7,0) = "IISSample.PageCounter"
ObjTotest(8,0) = "MSWC.PermissionChecker"
ObjTotest(9,0) = "Scripting.FileSystemObject"
	ObjTotest(9,1) = "(FSO 文本文件读写)"
ObjTotest(10,0) = "adodb.connection"
	ObjTotest(10,1) = "(ADO 数据对象)"
	
ObjTotest(11,0) = "SoftArtisans.FileUp"
	ObjTotest(11,1) = "(SA-FileUp 文件上传)"
ObjTotest(12,0) = "SoftArtisans.FileManager"
	ObjTotest(12,1) = "(SoftArtisans 文件管理)"
ObjTotest(13,0) = "LyfUpload.UploadFile"
	ObjTotest(13,1) = "(刘云峰的文件上传组件)"
ObjTotest(14,0) = "Persits.Upload.1"
	ObjTotest(14,1) = "(ASPUpload 文件上传)"
ObjTotest(15,0) = "w3.upload"
	ObjTotest(15,1) = "(Dimac 文件上传)"

ObjTotest(16,0) = "JMail.SmtpMail"
	ObjTotest(16,1) = "(Dimac JMail 邮件收发) <a href='http://www.ajiang.net'>中文手册下载</a>"
ObjTotest(17,0) = "CDONTS.NewMail"
	ObjTotest(17,1) = "(虚拟 SMTP 发信)"
ObjTotest(18,0) = "Persits.MailSEnder"
	ObjTotest(18,1) = "(ASPemail 发信)"
ObjTotest(19,0) = "SMTPsvg.Mailer"
	ObjTotest(19,1) = "(ASPmail 发信)"
ObjTotest(20,0) = "DkQmail.Qmail"
	ObjTotest(20,1) = "(dkQmail 发信)"
ObjTotest(21,0) = "Geocel.Mailer"
	ObjTotest(21,1) = "(Geocel 发信)"
ObjTotest(22,0) = "IISmail.Iismail.1"
	ObjTotest(22,1) = "(IISmail 发信)"
ObjTotest(23,0) = "SmtpMail.SmtpMail.1"
	ObjTotest(23,1) = "(SmtpMail 发信)"
	
ObjTotest(24,0) = "SoftArtisans.ImageGen"
	ObjTotest(24,1) = "(SA 的图像读写组件)"
ObjTotest(25,0) = "W3Image.Image"
	ObjTotest(25,1) = "(Dimac 的图像读写组件)"

public IsObj,VerObj,TestObj

'检查预查组件支持情况及版本

dim i
for i=0 to 25
	on error resume next
	IsObj=false
	VerObj=""
	'dim TestObj
	TestObj=""
	set TestObj=server.CreateObject(ObjTotest(i,0))
	if -2147221005 <> Err then		'感谢网友iAmFisher的宝贵建议
		IsObj = True
		VerObj = TestObj.version
		if VerObj="" or isnull(VerObj) then VerObj=TestObj.about
	End if
	ObjTotest(i,2)=IsObj
	ObjTotest(i,3)=VerObj
next

'检查组件是否被支持及组件版本的子程序
sub ObjTest(strObj)
	on error resume next
	IsObj=false
	VerObj=""
	TestObj=""
	set TestObj=server.CreateObject (strObj)
	if -2147221005 <> Err then		'感谢网友iAmFisher的宝贵建议
		IsObj = True
		VerObj = TestObj.version
		if VerObj="" or isnull(VerObj) then VerObj=TestObj.about
	End if	
End sub
%>
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<TITLE>ASP探针</TITLE>
<style>
<!--
BODY
{
	FONT-FAMILY: 宋体;
	FONT-SIZE: 9pt
}
TD
{
	FONT-SIZE: 9pt
}
A
{
	COLOR: #000000;
	TEXT-DECORATION: none
}
A:hover
{
	COLOR: #9FA9B3;
	TEXT-DECORATION: underline
}
.input
{
	BORDER: #9FA9B3 1px solid;
	FONT-SIZE: 9pt;
	BACKGROUND-color: #E3E3E3}
.backs
{
	BACKGROUND-COLOR: #E3E3E3;
	COLOR: #ffffff;

}
.backq
{
	BACKGROUND-COLOR: #E3E3E3
}
.backc
{
	BACKGROUND-COLOR: #E3E3E3;
	BORDER: medium none;
	COLOR: #999999;
	HEIGHT: 18px;
	font-size: 9pt
}
.fonts
{
	COLOR: #FFFFFF
}
.style1 {BACKGROUND-COLOR: #E3E3E3; color: #130700; }
-->
</STYLE>
</HEAD>
<BODY leftmargin="20">
<table width=550 height="30" border=0 cellpadding=0 cellspacing=1 bgcolor="#999999">
<tr>
  <td align="center" valign="middle" bgcolor="#e9e9e9"><strong>服务器测试</strong></td>
</tr></table>


<font class=fonts>服务器的有关参数</font>
<table border=0 width=550 cellspacing=0 cellpadding=0 bgcolor="#999999">
<tr><td>

	<table border=0 width=550 cellspacing=1 cellpadding=0>
	  <tr bgcolor="#E3E3E3" height=18>
		<td width="168" align=left>&nbsp;服务器名</td>
		<td width="379">&nbsp;<%=Request.ServerVariables("SERVER_NAME")%></td>
	  </tr>
	  <tr bgcolor="#E3E3E3" height=18>
		<td align=left>&nbsp;服务器IP</td>
		<td>&nbsp;<%=Request.ServerVariables("LOCAL_ADDR")%></td>
	  </tr>
	  <tr bgcolor="#E3E3E3" height=18>
		<td align=left>&nbsp;服务器端口</td>
		<td>&nbsp;<%=Request.ServerVariables("SERVER_PORT")%></td>
	  </tr>
	  <tr bgcolor="#E3E3E3" height=18>
		<td align=left>&nbsp;服务器时间</td>
		<td>&nbsp;<%=now%></td>
	  </tr>
	  <tr bgcolor="#E3E3E3" height=18>
		<td align=left>&nbsp;IIS版本</td>
		<td>&nbsp;<%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
	  </tr>
	  <tr bgcolor="#E3E3E3" height=18>
		<td align=left>&nbsp;脚本超时时间</td>
		<td>&nbsp;<%=Server.ScriptTimeout%> 秒</td>
	  </tr>
	  <tr bgcolor="#E3E3E3" height=18>
		<td align=left>&nbsp;本文件路径</td>
		<td>&nbsp;<%=Request.ServerVariables("PATH_TRANSLATED")%></td>
	  </tr>
	  <tr bgcolor="#E3E3E3" height=18>
		<td align=left>&nbsp;服务器CPU数量</td>
		<td>&nbsp;<%=Request.ServerVariables("NUMBER_OF_PROCESSORS")%> 个</td>
	  </tr>
	  <tr bgcolor="#E3E3E3" height=18>
		<td align=left>&nbsp;服务器解译引擎</td>
		<td>&nbsp;<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
	  </tr>
	  <tr bgcolor="#E3E3E3" height=18>
		<td align=left>&nbsp;服务器操作系统</td>
		<td>&nbsp;<%=Request.ServerVariables("OS")%></td>
	  </tr>
	</table>

</td></tr>
</table>
<br>
<font class=fonts>组件支持情况</font>
<%
Dim strClass
	strClass = Trim(Request.Form("classname"))
	if "" <> strClass then
	Response.Write "<br>您指定的组件的检查结果："
	Dim Verobj1
	ObjTest(strClass)
	  if Not IsObj then 
		Response.Write "<br><font color=red>很遗憾，该服务器不支持 " & strclass & " 组件！</font>"
	  else
		if VerObj="" or isnull(VerObj) then 
			Verobj1="无法取得该组件版本"
		else
			Verobj1="该组件版本是：" & VerObj
		End if
		Response.Write "<br><font class=fonts>恭喜！该服务器支持 " & strclass & " 组件。" & verobj1 & "</font>"
	  End if
	  Response.Write "<br>"
	End if
	%>


<br>■ IIS自带的ASP组件
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#999999" width="550">
	<tr height=18 class=style1 align=center><td width=320>组 件 名 称</td><td width=130>支持及版本</td></tr>
	<%For i=0 to 10%>
	<tr height="18" class=backq>
		<td align=left>&nbsp;<%=ObjTotest(i,0) & "<font color=#888888>&nbsp;" & ObjTotest(i,1)%></font></td>
		<td align=left>&nbsp;<%
		if Not ObjTotest(i,2) Then 
			Response.Write "<font color=red><b>×</b></font>"
		else
			Response.Write "<font class=fonts><b>√</b></font> <a title='" & ObjTotest(i,3) & "'>" & left(ObjTotest(i,3),11) & "</a>"
		End if%></td>
	</tr>
	<%next%>
</table>

<br>■ 常见的文件上传和管理组件
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#999999" width="550">
	<tr height=18 class=style1 align=center><td width=320>组 件 名 称</td><td width=130>支持及版本</td></tr>
	<%For i=11 to 15%>
	<tr height="18" class=backq>
		<td align=left>&nbsp;<%=ObjTotest(i,0) & "<font color=#888888>&nbsp;" & ObjTotest(i,1)%></font></td>
		<td align=left>&nbsp;<%
		if Not ObjTotest(i,2) Then 
			Response.Write "<font color=red><b>×</b></font>"
		else
			Response.Write "<font class=fonts><b>√</b></font> <a title='" & ObjTotest(i,3) & "'>" & left(ObjTotest(i,3),11) & "</a>"
		End if%></td>
	</tr>
	<%next%>
</table>

<br>■ 常见的收发邮件组件
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#999999" width="550">
	<tr height=18 class=style1 align=center><td width=320>组 件 名 称</td><td width=130>支持及版本</td></tr>
	<%For i=16 to 23%>
	<tr height="18" class=backq>
		<td align=left>&nbsp;<%=ObjTotest(i,0) & "<font color=#888888>&nbsp;" & ObjTotest(i,1)%></font></td>
		<td align=left>&nbsp;<%
		if Not ObjTotest(i,2) Then 
			Response.Write "<font color=red><b>×</b></font>"
		else
			Response.Write "<font class=fonts><b>√</b></font> <a title='" & ObjTotest(i,3) & "'>" & left(ObjTotest(i,3),11) & "</a>"
		End if%></td>
	</tr>
	<%next%>
</table>

<br>■ 图像处理组件
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#999999" width="550">
	<tr height=18 class=style1 align=center><td width=320>组 件 名 称</td><td width=130>支持及版本</td></tr>
	<%For i=24 to 25%>
	<tr height="18" class=backq>
		<td align=left>&nbsp;<%=ObjTotest(i,0) & "<font color=#888888>&nbsp;" & ObjTotest(i,1)%></font></td>
		<td align=left>&nbsp;<%
		if Not ObjTotest(i,2) Then 
			Response.Write "<font color=red><b>×</b></font>"
		else
			Response.Write "<font class=fonts><b>√</b></font> <a title='" & ObjTotest(i,3) & "'>" & left(ObjTotest(i,3),11) & "</a>"
		End if%></td>
	</tr>
	<%next%>
</table>

<br>■ 其他组件支持情况检测<br>
在下面的输入框中输入你要检测的组件的ProgId或ClassId。
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#999999" width="550">
<FORM action=<%=Request.ServerVariables("SCRIPT_NAME")%> method=post id=form1 name=form1>
	<tr height="18" class=backq>
		<td align=center height=30><input class=input type=text value="" name="classname" size=40>
<INPUT type=submit value=" 确 定 " class=backc id=submit1 name=submit1>
<INPUT type=reset value=" 重 填 " class=backc id=reset1 name=reset1> 
</td>
    </tr>
</FORM>
</table>

<%if ObjTest("Scripting.FileSystemObject") then

	set fsoobj=server.CreateObject("Scripting.FileSystemObject")

%>

<br>■ 当前文件夹信息
<%
	dPath = server.MapPath("./")
	set dDir = fsoObj.GetFolder(dPath)
	set dDrive = fsoObj.GetDrive(dDir.Drive)
%>
文件夹: <%=dPath%>
<table class=backq border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#999999" width="550">
  <tr height="18" align="center" class="backs">
	<td width="75" class="style1">已用空间</td>
	<td width="75" class="style1">可用空间</td>
	<td width="75" class="style1">文件夹数</td>
	<td width="75" class="style1">文件数</td>
	<td width="150" class="style1">创建时间</td>
  </tr>
  <tr height="18" align="center">
	<td><%=cSize(dDir.Size)%></td>
	<td><%=cSize(dDrive.AvailableSpace)%></td>
	<td><%=dDir.SubFolders.Count%></td>
	<td><%=dDir.Files.Count%></td>
	<td><%=dDir.DateCreated%></td>
  </tr>
</td></tr>
</table>
<br>
</BODY>
</HTML>

<%
end if
function cdrivetype(tnum)
    Select Case tnum
        Case 0: cdrivetype = "未知"
        Case 1: cdrivetype = "可移动磁盘"
        Case 2: cdrivetype = "本地硬盘"
        Case 3: cdrivetype = "网络磁盘"
        Case 4: cdrivetype = "CD-ROM"
        Case 5: cdrivetype = "RAM 磁盘"
    End Select
End function

function cIsReady(trd)
    Select Case trd
		case true: cIsReady="<font class=fonts><b>√</b></font>"
		case false: cIsReady="<font color='red'><b>×</b></font>"
	End Select
End function

function cSize(tSize)
    if tSize>=1073741824 then
		cSize=int((tSize/1073741824)*1000)/1000 & " GB"
    elseif tSize>=1048576 then
    	cSize=int((tSize/1048576)*1000)/1000 & " MB"
    elseif tSize>=1024 then
		cSize=int((tSize/1024)*1000)/1000 & " KB"
	else
		cSize=tSize & "B"
	End if
End function
%>