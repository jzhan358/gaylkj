<%@ Language="VBScript" %>
<%
' 不使用输出缓冲区，直接将运行结果显示在客户端
Response.Buffer = true

' 网页立即超时，防止缓存导致测速失败。
Response.Expires = -1

' 将检测的组件的列表
Dim OtT(3,15,1)
' 服务器变量
dim okCPUS, okCPU, okOS
' 检测组件变量
dim isobj,VerObj,TestObj

T = Request("T")
if T="" then T="ABGH"
%>

<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>万网组件说明系统</TITLE>
<style>
<!--
h1 {font-size:24px;color:#002E5B;font-family:Arial;margin:15px 0px 5px 0px}
h2 {font-size:16px;color:#0054A8;margin:15px 0px 8px 0px}
h3 {font-size:12px;color:#0054A8;font-family:Arial;margin:7px 0px 3px 12px;font-weight: normal;}
BODY,TD{FONT-FAMILY: 宋体;FONT-SIZE: 12px;word-break:break-all}
tr{BACKGROUND-COLOR: #F4F9FB}
table{BORDER: #336699 1px solid;background-color:#336699;margin-left:12px}
p{margin:5px 12px;color:#000000}
.input{BORDER: #111111 1px solid;FONT-SIZE: 9pt;BACKGROUND-color: #F8FFF0}
.backs{BACKGROUND-COLOR: #0054A8;COLOR: #ffffff;text-align:center}
.backc{BACKGROUND-COLOR: #0054A8;BORDER: medium none;COLOR: #ffffff;HEIGHT: 18px;font-size: 9pt}
.fonts{	COLOR: #3F8805}
-->
</STYLE>
</HEAD>
<body>


<center>
<h1>万网组件说明系统</h>
<%
for qq=1 to len(T)
  call BodyGo(mid(T,qq,1))
next
%>
</center>
<br>
<br>
<%
sub servinfo()
%>
  <h2>服务器概况</h2>
	<table border=0 width=500 cellspacing=1 cellpadding=3>
	  <tr>
		<td width="110">服务器地址</td><td width="390">名称 <%=Request.ServerVariables("SERVER_NAME")%>(IP:<%=Request.ServerVariables("LOCAL_ADDR")%>) 端口:<%=Request.ServerVariables("SERVER_PORT")%></td>
	  </tr>
	  <%
	  tnow = now():oknow = cstr(tnow)
	  if oknow <> year(tnow) & "-" & month(tnow) & "-" & day(tnow) & " " & hour(tnow) & ":" & right(FormatNumber(minute(tnow)/100,2),2) & ":" & right(FormatNumber(second(tnow)/100,2),2) then oknow = oknow & " (日期格式不规范)"
	  %>
	  <tr>
		<td>服务器时间</td><td><%=oknow%></td>
	  </tr>
	  <tr>
		<td>IIS版本</td><td><%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
	  </tr>
	  <tr>
		<td>脚本超时时间</td><td><%=Server.ScriptTimeout%> 秒</td>
	  </tr>
	  <tr>
		<td>服务器脚本引擎</td><td><%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %> , <%="JScript/" & getjver()%></td>
	  </tr>
	  <%getsysinfo()  '获得服务器数据%>
	  <tr>
		<td>服务器操作系统</td><td>
		<%
		strOS = Request.ServerVariables("SERVER_SOFTWARE")
		if instr(strOS,"IIS/6.0") then 
			os = "Microsoft Windows 2003"
		elseif instr(strOS, "IIS/7") then 
			os = "Microsoft Windows 2008"
		elseif instr(strOS, "IIS/5.0") then 
			os = "Microsoft Windows 2000"
		else 
			os = "未知"
		end if 
		response.Write os
		
		%></td>
	  </tr>
	</table>
<%
end sub

%>
<SCRIPT language="JScript" runat="server">
function getJVer(){
  //获取JScript 版本
  return ScriptEngineMajorVersion() +"."+ScriptEngineMinorVersion()+"."+ ScriptEngineBuildVersion();
}
</SCRIPT>
<%

' 获取服务器常用参数
sub getsysinfo()
  on error resume next
  Set WshShell = server.CreateObject("WScript.Shell")
  Set WshSysEnv = WshShell.Environment("SYSTEM")
  okCPUS = cstr(WshSysEnv("NUMBER_OF_PROCESSORS"))
  okCPU = cstr(WshSysEnv("PROCESSOR_IDENTIFIER"))
  if isempty(okCPUS) then
    okCPUS = Request.ServerVariables("NUMBER_OF_PROCESSORS")
  end if
  if okCPUS & "" = "" then
    okCPUS = "(未知)"
  end if
end sub



' *******************************************************************************
' 　　[ G ] 组件检测
' *******************************************************************************
sub comlist()
  on error resume next
  OtT(0,0,0) = "MSWC.BrowserType"
  OtT(0,1,0) = "Microsoft.XMLHTTP"
	OtT(0,1,1) = "(Http 组件, 常在采集系统中用到)"
  OtT(0,2,0) = "Scripting.FileSystemObject"
	OtT(0,2,1) = "(FSO 文件系统管理、文本文件读写)"
  OtT(0,3,0) = "Adodb.Connection"
	OtT(0,3,1) = "(ADO 数据对象)"
  OtT(0,4,0) = "Adodb.Stream"
	OtT(0,4,1) = "(ADO 数据流对象, 常见被用在无组件上传程序中)"
	

  OtT(1,0,0) = "Hichinafso.Upload"
	OtT(1,0,1) = "(万网Hichinafso 上传组件)"
  OtT(1,1,0) = "Persits.Upload.1"
	OtT(1,1,1) = "(ASPUpload 文件上传)"
  OtT(1,2,0) = "SoftArtisans.FileUp"
	OtT(1,2,1) = "(SA-FileUp 文件上传)"	
  OtT(1,3,0) = "SoftArtisans.FileManager"
	OtT(1,3,1) = "(SoftArtisans 文件管理)"
		
  OtT(2,0,0) = "JMail.SmtpMail"
	OtT(2,0,1) = "(Dimac JMail 邮件收发)"
  OtT(2,1,0) = "CDO.Message"
	OtT(2,1,1) = "(CDOSYS)"
	

  OtT(3,0,0) = "WinHttp.WinHttpRequest.5.1"
	OtT(3,0,1) = "(Http 组件, 常在采集系统中用到)"
  OtT(3,1,0) = "Persits.Jpeg"
	OtT(3,1,1) = "(ASPJpeg)"
  OtT(3,2,0) = "PE_Common6.GetVersion"
	OtT(3,2,1) = "(动易2006组件支持)"
  OtT(3,3,0) = "md5_vb.md5class"
	OtT(3,3,1) = "(md5数据加密)"
%>
<h2>其他组件支持情况</h2>

<h3>■ 检查组件是否被支持</h3>
<table border=0 width=500 cellspacing=1 cellpadding=3>
  <FORM action="?T=<%=T%>#G" method="post">
  <tr><td align="center" style="padding:10px 0px">
  在下面的文本框中输入您要检测的组件的 ProgId 或 ClassId
  <input class=input type=text value="" name="classname" size=50>
  <input type=submit value=" 检 查 " class=backc id=submit1 name=submit1>
<%
Dim strClass
strClass = Trim(Request.Form("classname"))
If "" <> strClass then
Response.Write "<p style=""margin:9px 0px 0px 0px"">"
Dim Verobj1
ObjTest(strClass)
  If Not IsObj then 
	Response.Write "<font color=red>很遗憾，该服务器不支持 " & strclass & " 组件！</font>"
  Else
	if VerObj="" or isnull(VerObj) then 
	  Verobj1="无法取得该组件版本"
	Else
	  Verobj1="该组件版本是：" & VerObj
	End If
	Response.Write "<font class=fonts>恭喜！该服务器支持 " & strclass & " 组件。" & verobj1 & "</font>"
  End If
end if
%>
  </td></tr>
  </FORM>
</table>

<h3>■ 操作系统自带的组件</h3>
<table border=0 width=500 cellspacing=1 cellpadding=3>
  <tr class="backs"><td width="380">组件名称及简介</td><td width="120">支持/版本</td></tr>
  <%
  k=0
  for i=0 to 4
    call ObjTest(OtT(k,i,0))
  %>
  <tr><td width="380"><%=OtT(k,i,0) & " <font color='#888888'>" & OtT(k,i,1) & "</font>"%></td><td width="120" title="<%=VerObj%>"><%=cIsReady(isobj) & " " & VerObj %></td></tr>
  <%next%>
</table>

<h3>■ 常见文件上传和管理组件</h3>
<table border=0 width=500 cellspacing=1 cellpadding=3>
  <tr class="backs"><td width="380">组件名称及简介</td><td width="120">支持/版本</td></tr>
  <%
  k=1
  for i=0 to 3
    call ObjTest(OtT(k,i,0))
  %>
  <tr><td width="380"><%=OtT(k,i,0) & " <font color='#888888'>" & OtT(k,i,1) & "</font>"%></td><td width="120" title="<%=VerObj%>"><%=cIsReady(isobj) & " " & VerObj %></td></tr>
  <%next%>
</table>

<h3>■ 常见邮件处理组件</h3>
<table border=0 width=500 cellspacing=1 cellpadding=3>
  <tr class="backs"><td width="380">组件名称及简介</td><td width="120">支持/版本</td></tr>
  <%
  k=2
  for i=0 to 1
    call ObjTest(OtT(k,i,0))
  %>
  <tr><td width="380"><%=OtT(k,i,0) & " <font color='#888888'>" & OtT(k,i,1) & "</font>"%></td><td width="120" title="<%=VerObj%>"><%=cIsReady(isobj) & " " & VerObj %></td></tr>
  <%next%>
</table>

<h3>■ 其它常见组件</h3>
<table border=0 width=500 cellspacing=1 cellpadding=3>
  <tr class="backs"><td width="380">组件名称及简介</td><td width="120">支持/版本</td></tr>
  <%
  k=3
  for i=0 to 3
    call ObjTest(OtT(k,i,0))
  %>
  <tr><td width="380"><%=OtT(k,i,0) & " <font color='#888888'>" & OtT(k,i,1) & "</font>"%></td><td width="120" title="<%=VerObj%>"><%=cIsReady(isobj) & " " & VerObj %></td></tr>
  <%next%>
</table>
<%
	
end sub




' *******************************************************************************
' 　　其他函数和子程序
' *******************************************************************************

' 展示栏目
sub BodyGo(gCon)
  select case gCon
  case "B"
    call servinfo()
  case "E"
    call sevalist()
  case "G"
    call comlist()
  end select
end sub


' 将是否可用转换为对号和错号
function cIsReady(trd)
  Select Case trd
    case true: cIsReady="<font class=fonts><b>√</b></font>"
    case false: cIsReady="<font color='red'><b>×</b></font>"
  End Select
end function

'检查组件是否被支持及组件版本的子程序
sub ObjTest(strObj)
  on error resume next
  IsObj=false
  VerObj=""
  set TestObj=server.CreateObject (strObj)
  If Err.Number = 0 then
    IsObj = True
		if strobj <>  "PE_Common6.GetVersion" then 
    		VerObj = TestObj.version
		else 
			VerObj = TestObj.strVersion
		end if 
   		if VerObj="" or isnull(VerObj) then VerObj=TestObj.about
  end if
  set TestObj=nothing
End sub

%><!-- 
2009年5月7日，刘辉。根据网上最流行的探针进行增减，删除不必要的冗余代码，只为适合于万网主机使用
http://www.anywolfs.com/liuhui
-->