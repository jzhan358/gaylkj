<%@ Language="VBScript" %>
<%
' ��ʹ�������������ֱ�ӽ����н����ʾ�ڿͻ���
Response.Buffer = true

' ��ҳ������ʱ����ֹ���浼�²���ʧ�ܡ�
Response.Expires = -1

' ������������б�
Dim OtT(3,15,1)
' ����������
dim okCPUS, okCPU, okOS
' ����������
dim isobj,VerObj,TestObj

T = Request("T")
if T="" then T="ABGH"
%>

<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>�������˵��ϵͳ</TITLE>
<style>
<!--
h1 {font-size:24px;color:#002E5B;font-family:Arial;margin:15px 0px 5px 0px}
h2 {font-size:16px;color:#0054A8;margin:15px 0px 8px 0px}
h3 {font-size:12px;color:#0054A8;font-family:Arial;margin:7px 0px 3px 12px;font-weight: normal;}
BODY,TD{FONT-FAMILY: ����;FONT-SIZE: 12px;word-break:break-all}
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
<h1>�������˵��ϵͳ</h>
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
  <h2>�������ſ�</h2>
	<table border=0 width=500 cellspacing=1 cellpadding=3>
	  <tr>
		<td width="110">��������ַ</td><td width="390">���� <%=Request.ServerVariables("SERVER_NAME")%>(IP:<%=Request.ServerVariables("LOCAL_ADDR")%>) �˿�:<%=Request.ServerVariables("SERVER_PORT")%></td>
	  </tr>
	  <%
	  tnow = now():oknow = cstr(tnow)
	  if oknow <> year(tnow) & "-" & month(tnow) & "-" & day(tnow) & " " & hour(tnow) & ":" & right(FormatNumber(minute(tnow)/100,2),2) & ":" & right(FormatNumber(second(tnow)/100,2),2) then oknow = oknow & " (���ڸ�ʽ���淶)"
	  %>
	  <tr>
		<td>������ʱ��</td><td><%=oknow%></td>
	  </tr>
	  <tr>
		<td>IIS�汾</td><td><%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
	  </tr>
	  <tr>
		<td>�ű���ʱʱ��</td><td><%=Server.ScriptTimeout%> ��</td>
	  </tr>
	  <tr>
		<td>�������ű�����</td><td><%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %> , <%="JScript/" & getjver()%></td>
	  </tr>
	  <%getsysinfo()  '��÷���������%>
	  <tr>
		<td>����������ϵͳ</td><td>
		<%
		strOS = Request.ServerVariables("SERVER_SOFTWARE")
		if instr(strOS,"IIS/6.0") then 
			os = "Microsoft Windows 2003"
		elseif instr(strOS, "IIS/7") then 
			os = "Microsoft Windows 2008"
		elseif instr(strOS, "IIS/5.0") then 
			os = "Microsoft Windows 2000"
		else 
			os = "δ֪"
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
  //��ȡJScript �汾
  return ScriptEngineMajorVersion() +"."+ScriptEngineMinorVersion()+"."+ ScriptEngineBuildVersion();
}
</SCRIPT>
<%

' ��ȡ���������ò���
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
    okCPUS = "(δ֪)"
  end if
end sub



' *******************************************************************************
' ����[ G ] ������
' *******************************************************************************
sub comlist()
  on error resume next
  OtT(0,0,0) = "MSWC.BrowserType"
  OtT(0,1,0) = "Microsoft.XMLHTTP"
	OtT(0,1,1) = "(Http ���, ���ڲɼ�ϵͳ���õ�)"
  OtT(0,2,0) = "Scripting.FileSystemObject"
	OtT(0,2,1) = "(FSO �ļ�ϵͳ�����ı��ļ���д)"
  OtT(0,3,0) = "Adodb.Connection"
	OtT(0,3,1) = "(ADO ���ݶ���)"
  OtT(0,4,0) = "Adodb.Stream"
	OtT(0,4,1) = "(ADO ����������, ����������������ϴ�������)"
	

  OtT(1,0,0) = "Hichinafso.Upload"
	OtT(1,0,1) = "(����Hichinafso �ϴ����)"
  OtT(1,1,0) = "Persits.Upload.1"
	OtT(1,1,1) = "(ASPUpload �ļ��ϴ�)"
  OtT(1,2,0) = "SoftArtisans.FileUp"
	OtT(1,2,1) = "(SA-FileUp �ļ��ϴ�)"	
  OtT(1,3,0) = "SoftArtisans.FileManager"
	OtT(1,3,1) = "(SoftArtisans �ļ�����)"
		
  OtT(2,0,0) = "JMail.SmtpMail"
	OtT(2,0,1) = "(Dimac JMail �ʼ��շ�)"
  OtT(2,1,0) = "CDO.Message"
	OtT(2,1,1) = "(CDOSYS)"
	

  OtT(3,0,0) = "WinHttp.WinHttpRequest.5.1"
	OtT(3,0,1) = "(Http ���, ���ڲɼ�ϵͳ���õ�)"
  OtT(3,1,0) = "Persits.Jpeg"
	OtT(3,1,1) = "(ASPJpeg)"
  OtT(3,2,0) = "PE_Common6.GetVersion"
	OtT(3,2,1) = "(����2006���֧��)"
  OtT(3,3,0) = "md5_vb.md5class"
	OtT(3,3,1) = "(md5���ݼ���)"
%>
<h2>�������֧�����</h2>

<h3>�� �������Ƿ�֧��</h3>
<table border=0 width=500 cellspacing=1 cellpadding=3>
  <FORM action="?T=<%=T%>#G" method="post">
  <tr><td align="center" style="padding:10px 0px">
  ��������ı�����������Ҫ��������� ProgId �� ClassId
  <input class=input type=text value="" name="classname" size=50>
  <input type=submit value=" �� �� " class=backc id=submit1 name=submit1>
<%
Dim strClass
strClass = Trim(Request.Form("classname"))
If "" <> strClass then
Response.Write "<p style=""margin:9px 0px 0px 0px"">"
Dim Verobj1
ObjTest(strClass)
  If Not IsObj then 
	Response.Write "<font color=red>���ź����÷�������֧�� " & strclass & " �����</font>"
  Else
	if VerObj="" or isnull(VerObj) then 
	  Verobj1="�޷�ȡ�ø�����汾"
	Else
	  Verobj1="������汾�ǣ�" & VerObj
	End If
	Response.Write "<font class=fonts>��ϲ���÷�����֧�� " & strclass & " �����" & verobj1 & "</font>"
  End If
end if
%>
  </td></tr>
  </FORM>
</table>

<h3>�� ����ϵͳ�Դ������</h3>
<table border=0 width=500 cellspacing=1 cellpadding=3>
  <tr class="backs"><td width="380">������Ƽ����</td><td width="120">֧��/�汾</td></tr>
  <%
  k=0
  for i=0 to 4
    call ObjTest(OtT(k,i,0))
  %>
  <tr><td width="380"><%=OtT(k,i,0) & " <font color='#888888'>" & OtT(k,i,1) & "</font>"%></td><td width="120" title="<%=VerObj%>"><%=cIsReady(isobj) & " " & VerObj %></td></tr>
  <%next%>
</table>

<h3>�� �����ļ��ϴ��͹������</h3>
<table border=0 width=500 cellspacing=1 cellpadding=3>
  <tr class="backs"><td width="380">������Ƽ����</td><td width="120">֧��/�汾</td></tr>
  <%
  k=1
  for i=0 to 3
    call ObjTest(OtT(k,i,0))
  %>
  <tr><td width="380"><%=OtT(k,i,0) & " <font color='#888888'>" & OtT(k,i,1) & "</font>"%></td><td width="120" title="<%=VerObj%>"><%=cIsReady(isobj) & " " & VerObj %></td></tr>
  <%next%>
</table>

<h3>�� �����ʼ��������</h3>
<table border=0 width=500 cellspacing=1 cellpadding=3>
  <tr class="backs"><td width="380">������Ƽ����</td><td width="120">֧��/�汾</td></tr>
  <%
  k=2
  for i=0 to 1
    call ObjTest(OtT(k,i,0))
  %>
  <tr><td width="380"><%=OtT(k,i,0) & " <font color='#888888'>" & OtT(k,i,1) & "</font>"%></td><td width="120" title="<%=VerObj%>"><%=cIsReady(isobj) & " " & VerObj %></td></tr>
  <%next%>
</table>

<h3>�� �����������</h3>
<table border=0 width=500 cellspacing=1 cellpadding=3>
  <tr class="backs"><td width="380">������Ƽ����</td><td width="120">֧��/�汾</td></tr>
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
' ���������������ӳ���
' *******************************************************************************

' չʾ��Ŀ
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


' ���Ƿ����ת��Ϊ�Ժźʹ��
function cIsReady(trd)
  Select Case trd
    case true: cIsReady="<font class=fonts><b>��</b></font>"
    case false: cIsReady="<font color='red'><b>��</b></font>"
  End Select
end function

'�������Ƿ�֧�ּ�����汾���ӳ���
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
2009��5��7�գ����ԡ��������������е�̽�����������ɾ������Ҫ��������룬ֻΪ�ʺ�����������ʹ��
http://www.anywolfs.com/liuhui
-->