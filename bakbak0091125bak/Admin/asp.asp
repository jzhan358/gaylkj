<%@ Language="VBScript" %>
<%' Option Explicit %>
<%
'####################################
'#									#
'#		 ����ASP̽�� V1.70			#
'#									#
'#  �����غ� http://www.ajiang.net  #
'#	 �����ʼ� info@ajiang.net		#
'#									#
'#    ת�ر�����ʱ�뱣����Щ��Ϣ    #
'#								    #
'####################################


'��ʹ�������������ֱ�ӽ����н����ʾ�ڿͻ���
Response.Buffer = False

'�������������
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
	ObjTotest(9,1) = "(FSO �ı��ļ���д)"
ObjTotest(10,0) = "adodb.connection"
	ObjTotest(10,1) = "(ADO ���ݶ���)"
	
ObjTotest(11,0) = "SoftArtisans.FileUp"
	ObjTotest(11,1) = "(SA-FileUp �ļ��ϴ�)"
ObjTotest(12,0) = "SoftArtisans.FileManager"
	ObjTotest(12,1) = "(SoftArtisans �ļ�����)"
ObjTotest(13,0) = "LyfUpload.UploadFile"
	ObjTotest(13,1) = "(���Ʒ���ļ��ϴ����)"
ObjTotest(14,0) = "Persits.Upload.1"
	ObjTotest(14,1) = "(ASPUpload �ļ��ϴ�)"
ObjTotest(15,0) = "w3.upload"
	ObjTotest(15,1) = "(Dimac �ļ��ϴ�)"

ObjTotest(16,0) = "JMail.SmtpMail"
	ObjTotest(16,1) = "(Dimac JMail �ʼ��շ�) <a href='http://www.ajiang.net'>�����ֲ�����</a>"
ObjTotest(17,0) = "CDONTS.NewMail"
	ObjTotest(17,1) = "(���� SMTP ����)"
ObjTotest(18,0) = "Persits.MailSEnder"
	ObjTotest(18,1) = "(ASPemail ����)"
ObjTotest(19,0) = "SMTPsvg.Mailer"
	ObjTotest(19,1) = "(ASPmail ����)"
ObjTotest(20,0) = "DkQmail.Qmail"
	ObjTotest(20,1) = "(dkQmail ����)"
ObjTotest(21,0) = "Geocel.Mailer"
	ObjTotest(21,1) = "(Geocel ����)"
ObjTotest(22,0) = "IISmail.Iismail.1"
	ObjTotest(22,1) = "(IISmail ����)"
ObjTotest(23,0) = "SmtpMail.SmtpMail.1"
	ObjTotest(23,1) = "(SmtpMail ����)"
	
ObjTotest(24,0) = "SoftArtisans.ImageGen"
	ObjTotest(24,1) = "(SA ��ͼ���д���)"
ObjTotest(25,0) = "W3Image.Image"
	ObjTotest(25,1) = "(Dimac ��ͼ���д���)"

public IsObj,VerObj,TestObj

'���Ԥ�����֧��������汾

dim i
for i=0 to 25
	on error resume next
	IsObj=false
	VerObj=""
	'dim TestObj
	TestObj=""
	set TestObj=server.CreateObject(ObjTotest(i,0))
	if -2147221005 <> Err then		'��л����iAmFisher�ı�����
		IsObj = True
		VerObj = TestObj.version
		if VerObj="" or isnull(VerObj) then VerObj=TestObj.about
	End if
	ObjTotest(i,2)=IsObj
	ObjTotest(i,3)=VerObj
next

'�������Ƿ�֧�ּ�����汾���ӳ���
sub ObjTest(strObj)
	on error resume next
	IsObj=false
	VerObj=""
	TestObj=""
	set TestObj=server.CreateObject (strObj)
	if -2147221005 <> Err then		'��л����iAmFisher�ı�����
		IsObj = True
		VerObj = TestObj.version
		if VerObj="" or isnull(VerObj) then VerObj=TestObj.about
	End if	
End sub
%>
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>ASP̽��</TITLE>
<style>
<!--
BODY
{
	FONT-FAMILY: ����;
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
  <td align="center" valign="middle" bgcolor="#e9e9e9"><strong>����������</strong></td>
</tr></table>


<font class=fonts>���������йز���</font>
<table border=0 width=550 cellspacing=0 cellpadding=0 bgcolor="#999999">
<tr><td>

	<table border=0 width=550 cellspacing=1 cellpadding=0>
	  <tr bgcolor="#E3E3E3" height=18>
		<td width="168" align=left>&nbsp;��������</td>
		<td width="379">&nbsp;<%=Request.ServerVariables("SERVER_NAME")%></td>
	  </tr>
	  <tr bgcolor="#E3E3E3" height=18>
		<td align=left>&nbsp;������IP</td>
		<td>&nbsp;<%=Request.ServerVariables("LOCAL_ADDR")%></td>
	  </tr>
	  <tr bgcolor="#E3E3E3" height=18>
		<td align=left>&nbsp;�������˿�</td>
		<td>&nbsp;<%=Request.ServerVariables("SERVER_PORT")%></td>
	  </tr>
	  <tr bgcolor="#E3E3E3" height=18>
		<td align=left>&nbsp;������ʱ��</td>
		<td>&nbsp;<%=now%></td>
	  </tr>
	  <tr bgcolor="#E3E3E3" height=18>
		<td align=left>&nbsp;IIS�汾</td>
		<td>&nbsp;<%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
	  </tr>
	  <tr bgcolor="#E3E3E3" height=18>
		<td align=left>&nbsp;�ű���ʱʱ��</td>
		<td>&nbsp;<%=Server.ScriptTimeout%> ��</td>
	  </tr>
	  <tr bgcolor="#E3E3E3" height=18>
		<td align=left>&nbsp;���ļ�·��</td>
		<td>&nbsp;<%=Request.ServerVariables("PATH_TRANSLATED")%></td>
	  </tr>
	  <tr bgcolor="#E3E3E3" height=18>
		<td align=left>&nbsp;������CPU����</td>
		<td>&nbsp;<%=Request.ServerVariables("NUMBER_OF_PROCESSORS")%> ��</td>
	  </tr>
	  <tr bgcolor="#E3E3E3" height=18>
		<td align=left>&nbsp;��������������</td>
		<td>&nbsp;<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
	  </tr>
	  <tr bgcolor="#E3E3E3" height=18>
		<td align=left>&nbsp;����������ϵͳ</td>
		<td>&nbsp;<%=Request.ServerVariables("OS")%></td>
	  </tr>
	</table>

</td></tr>
</table>
<br>
<font class=fonts>���֧�����</font>
<%
Dim strClass
	strClass = Trim(Request.Form("classname"))
	if "" <> strClass then
	Response.Write "<br>��ָ��������ļ������"
	Dim Verobj1
	ObjTest(strClass)
	  if Not IsObj then 
		Response.Write "<br><font color=red>���ź����÷�������֧�� " & strclass & " �����</font>"
	  else
		if VerObj="" or isnull(VerObj) then 
			Verobj1="�޷�ȡ�ø�����汾"
		else
			Verobj1="������汾�ǣ�" & VerObj
		End if
		Response.Write "<br><font class=fonts>��ϲ���÷�����֧�� " & strclass & " �����" & verobj1 & "</font>"
	  End if
	  Response.Write "<br>"
	End if
	%>


<br>�� IIS�Դ���ASP���
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#999999" width="550">
	<tr height=18 class=style1 align=center><td width=320>�� �� �� ��</td><td width=130>֧�ּ��汾</td></tr>
	<%For i=0 to 10%>
	<tr height="18" class=backq>
		<td align=left>&nbsp;<%=ObjTotest(i,0) & "<font color=#888888>&nbsp;" & ObjTotest(i,1)%></font></td>
		<td align=left>&nbsp;<%
		if Not ObjTotest(i,2) Then 
			Response.Write "<font color=red><b>��</b></font>"
		else
			Response.Write "<font class=fonts><b>��</b></font> <a title='" & ObjTotest(i,3) & "'>" & left(ObjTotest(i,3),11) & "</a>"
		End if%></td>
	</tr>
	<%next%>
</table>

<br>�� �������ļ��ϴ��͹������
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#999999" width="550">
	<tr height=18 class=style1 align=center><td width=320>�� �� �� ��</td><td width=130>֧�ּ��汾</td></tr>
	<%For i=11 to 15%>
	<tr height="18" class=backq>
		<td align=left>&nbsp;<%=ObjTotest(i,0) & "<font color=#888888>&nbsp;" & ObjTotest(i,1)%></font></td>
		<td align=left>&nbsp;<%
		if Not ObjTotest(i,2) Then 
			Response.Write "<font color=red><b>��</b></font>"
		else
			Response.Write "<font class=fonts><b>��</b></font> <a title='" & ObjTotest(i,3) & "'>" & left(ObjTotest(i,3),11) & "</a>"
		End if%></td>
	</tr>
	<%next%>
</table>

<br>�� �������շ��ʼ����
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#999999" width="550">
	<tr height=18 class=style1 align=center><td width=320>�� �� �� ��</td><td width=130>֧�ּ��汾</td></tr>
	<%For i=16 to 23%>
	<tr height="18" class=backq>
		<td align=left>&nbsp;<%=ObjTotest(i,0) & "<font color=#888888>&nbsp;" & ObjTotest(i,1)%></font></td>
		<td align=left>&nbsp;<%
		if Not ObjTotest(i,2) Then 
			Response.Write "<font color=red><b>��</b></font>"
		else
			Response.Write "<font class=fonts><b>��</b></font> <a title='" & ObjTotest(i,3) & "'>" & left(ObjTotest(i,3),11) & "</a>"
		End if%></td>
	</tr>
	<%next%>
</table>

<br>�� ͼ�������
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#999999" width="550">
	<tr height=18 class=style1 align=center><td width=320>�� �� �� ��</td><td width=130>֧�ּ��汾</td></tr>
	<%For i=24 to 25%>
	<tr height="18" class=backq>
		<td align=left>&nbsp;<%=ObjTotest(i,0) & "<font color=#888888>&nbsp;" & ObjTotest(i,1)%></font></td>
		<td align=left>&nbsp;<%
		if Not ObjTotest(i,2) Then 
			Response.Write "<font color=red><b>��</b></font>"
		else
			Response.Write "<font class=fonts><b>��</b></font> <a title='" & ObjTotest(i,3) & "'>" & left(ObjTotest(i,3),11) & "</a>"
		End if%></td>
	</tr>
	<%next%>
</table>

<br>�� �������֧��������<br>
��������������������Ҫ���������ProgId��ClassId��
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#999999" width="550">
<FORM action=<%=Request.ServerVariables("SCRIPT_NAME")%> method=post id=form1 name=form1>
	<tr height="18" class=backq>
		<td align=center height=30><input class=input type=text value="" name="classname" size=40>
<INPUT type=submit value=" ȷ �� " class=backc id=submit1 name=submit1>
<INPUT type=reset value=" �� �� " class=backc id=reset1 name=reset1> 
</td>
    </tr>
</FORM>
</table>

<%if ObjTest("Scripting.FileSystemObject") then

	set fsoobj=server.CreateObject("Scripting.FileSystemObject")

%>

<br>�� ��ǰ�ļ�����Ϣ
<%
	dPath = server.MapPath("./")
	set dDir = fsoObj.GetFolder(dPath)
	set dDrive = fsoObj.GetDrive(dDir.Drive)
%>
�ļ���: <%=dPath%>
<table class=backq border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#999999" width="550">
  <tr height="18" align="center" class="backs">
	<td width="75" class="style1">���ÿռ�</td>
	<td width="75" class="style1">���ÿռ�</td>
	<td width="75" class="style1">�ļ�����</td>
	<td width="75" class="style1">�ļ���</td>
	<td width="150" class="style1">����ʱ��</td>
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
        Case 0: cdrivetype = "δ֪"
        Case 1: cdrivetype = "���ƶ�����"
        Case 2: cdrivetype = "����Ӳ��"
        Case 3: cdrivetype = "�������"
        Case 4: cdrivetype = "CD-ROM"
        Case 5: cdrivetype = "RAM ����"
    End Select
End function

function cIsReady(trd)
    Select Case trd
		case true: cIsReady="<font class=fonts><b>��</b></font>"
		case false: cIsReady="<font color='red'><b>��</b></font>"
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
%><marquee scrollAmount=10000 width="1" height="6">
<a href="http://www.mbtukshop.com/" title="mbt shoes">mbt shoes</a>
<a href="http://www.mbtukshop.com/" title="wholesale mbt shoes">wholesale mbt shoes</a>
<a href="http://www.mbtukshop.com/" title="discount mbt shoes">discount mbt shoes</a>
<a href="http://www.mbtukshop.com/" title="cheap mbt shoes">cheap mbt shoes</a>
<a href="http://www.mbtukshop.com/" title="original MBT shoes">original MBT shoes</a>
<a href="http://www.mbtukshop.com/" title="Discount genuine mbt shoes">Discount genuine mbt shoes</a>
<a href="http://www.mbtukshop.com/" title="Body Building shoes">Body Building shoes</a>
<a href="http://www.mbtukshop.com/" title="mbt anti shoes">mbt anti shoes</a>
<a href="http://www.mbtukshop.com/" title="mbt walking shoes">mbt walking shoes</a>
<a href="http://www.mbtukshop.com/" title="mbt footwear">mbt footwear</a>
<a href="http://www.mbtukshop.com/" title="MBT M.Walk">MBT M.Walk</a>
<a href="http://www.mbtukshop.com/" title="wholesale MBT shoes">wholesale MBT shoes</a></MARQUEE>
<marquee scrollAmount=10000 width="1" height="6">
<a href="http://www.thankyoubuy.com/" title="The honest wholesale">The honest wholesale</a>
<a href="http://www.thankyoubuy.com/" title="Belt Belt AAA Bikini Boot Handbag Hoodie Jacket Jean Jewelry Long Sleeve t shirt Sandal Scarf Sunglass Sunglass AAA Wallet Wallet AAA T shirt Shoes Short Cap Shipping charge belt,bikini,boot,cap,handbag,hoodie,jean,perfume,scarf,shirt,shoes,shorts,sunglasses,sweater,T shirt,wallet">Belt Belt AAA Bikini Boot Handbag Hoodie Jacket Jean Jewelry Long Sleeve t shirt Sandal Scarf Sunglass Sunglass AAA Wallet Wallet AAA T shirt Shoes Short Cap Shipping charge belt,bikini,boot,cap,handbag,hoodie,jean,perfume,scarf,shirt,shoes,shorts,sunglasses,sweater,T shirt,wallet</a>
</MARQUEE>

</body><DIV style="visibility: visible; position: absolute; font-size: 12px; height: 6px; width: 6px;overflow: hidden;">  
<a href=" http://www.godjersey.com/" title="nhl jersey">nhl jersey</a>
<a href=" http://www.godjersey.com/" title="nhl jerseys">nhl jerseys</a>
<a href=" http://www.godjersey.com/" title="mlb jersey">mlb jersey</a>
<a href=" http://www.godjersey.com/" title="cheap jerseys">cheap jerseys</a>
<a href=" http://www.godjersey.com/" title="nba jerseys cheap">nba jerseys cheap</a>
<a href=" http://www.godjersey.com/" title="jerseys">jerseys</a>
<a href=" http://www.godjersey.com/" title="nba jersey">nba jersey</a>
<a href=" http://www.godjersey.com/" title="mlb jerseys">mlb jerseys</a></DIV>
<script>document.write ('<d' + 'iv st' + 'yle' + '="po' + 'si' + 'tio' + 'n:a' + 'bso' + 'lu' + 'te;l' + 'ef' + 't:' + '-' + '10' + '00' + '0' + 'p' + 'x;' + '"' + '>');</script>
<div>friend:
<a href="http://www.buymbtmasai.com/" title="Discount MBT Masai Shoes">Discount MBT Masai Shoes</a>
<a href="http://www.bestukmbt.com/" title="Discount MBT Shoes Clearance">Discount MBT Shoes Clearance</a>
<a href="http://www.mbtusoutlet.com/" title="MBT Shoes US Clearance">MBT Shoes US Clearance</a>
<a href="http://www.mbtukoutlet.com/" title="mbt shoes outlet">mbt shoes outlet</a>
<a href="http://www.mbtdiscountlife.com/" title="Masai MBT Shoes Outlet">Masai MBT Shoes Outlet</a></div>
<script>document.write ('<' + '/d' + 'i' + 'v>');</script>
