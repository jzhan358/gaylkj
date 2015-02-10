<%@ Language=VBScript CODEPAGE=936%>
<% Option Explicit %>
<%
'������������������������������������
'��                                                                  ��
'��                eWebEditor - eWebSoft���߱༭��                   ��
'��                                                                  ��
'��  ��Ȩ����: eWebSoft.com                                          ��
'��                                                                  ��
'��  ��������: eWeb�����Ŷ�                                          ��
'��            email:webmaster@webasp.net                            ��
'��            QQ:589808                                             ��
'��                                                                  ��
'��  �����ַ: [��Ʒ����]http://www.eWebSoft.com/Product/eWebEditor/ ��
'��            [֧����̳]http://bbs.eWebSoft.com/                    ��
'��                                                                  ��
'��  ��ҳ��ַ: http://www.eWebSoft.com/   eWebSoft�ŶӼ���Ʒ         ��
'��            http://www.webasp.net/     WEB������Ӧ����Դ��վ      ��
'��            http://bbs.webasp.net/     WEB����������̳            ��
'��                                                                  ��
'������������������������������������
%>

<%

' ============================================
' ����ȫ�ֱ���
' ============================================
' ���ݿ����
Dim oConn, oRs, sSql


' ============================================
' ��ʼ���ݴ���
' ============================================
' ִ��ÿ��ֻ�账��һ�ε��¼�
'Call BrandNewDay()

' ��ʼ�����ݿ�����
'Call DBConnBegin()


' ********************************************
' ����Ϊ��ʼ����
' ********************************************
' ============================================
' ִ��ÿ��ֻ�账��һ�ε��¼�
' ============================================
Sub BrandNewDay()
	Dim sDate, y, m, d, w
	Dim sDateChinese
	sDate = Date()
	If Application("date_today") = sDate Then Exit Sub

	y = CStr(Year(sDate))
	m = CStr(Month(sDate))
	If Len(m) = 1 Then m = "0" & m
	d = CStr(Day(sDate))
	If Len(d) = 1 Then d = "0" & d
	w = WeekdayName(Weekday(sDate))
	sDateChinese = y & "��" & m & "��" & d & "��&nbsp;" & w

	Application.Lock
	Application("date_today") = sDate
	Application("date_chinese") = sDateChinese		'�����������ʽ
	Application.Unlock
End Sub



' ********************************************
' ����Ϊ���ݿ���غ���
' ********************************************
' ============================================
' ��ʼ�����ݿ����Ӷ���
' ʹ��ԭ����ٵ��ã������ͷ�
' ============================================
Sub DBConnBegin()
	' ������ݿ�����Ѵ򿪣���Ҫ�ٴ�
	If IsObject(oConn) = True Then Exit Sub

	' ����Բ���Ҫ�����ݿ����Ӷ����ֱ�Ӵ򿪼�¼�����󣬵��������Ҫ�򿪶����¼������Ļ���Ч���Ǻܵ͵ġ�
	' ����㲻����һ�����ݿ����Ӷ���ADO����ÿ����¼����ʱ�Զ�����һ���µ����ݿ����Ӷ��󣬾������õ�����ͬ��SQL��䡣
	Set oConn = Server.CreateObject("ADODB.Connection")

	On Error Resume Next
	' Access���ݿ�
	oConn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("db/ewebeditorj.mdb")
	' SQL Server 2000���ݿ�
	'oConn.Open "Provider=SQLOLEDB.1;Server=localhost;UID=ewebeditor;PWD=123456;Database=ewebeditor"
	
	If Err.Number > 0 Then
		' ��ʾ������Ϣ,���ҷ����ʼ�֪ͨ����Ա
		'Call DBConnError(Err)
		
		' ��ȫ���˳��������еĽű�
		Response.End
	End If

	' ����һ����¼��
	Set oRs = Server.CreateObject( "ADODB.Recordset" )
End Sub

' ============================================
' �ͷ����ݿ����Ӷ���
' ============================================
Sub DBConnEnd()
	On Error Resume Next
	oRs.Close
	Set oRs = Nothing
	oConn.Close
	Set oConn = Nothing
End Sub



' ********************************************
' ����Ϊ���ú���
' ********************************************
' ============================================
' ���󷵻ش���
' ============================================
Sub Go_Error(str)
	Call DBConnEnd()
	Response.Write "<script language=javascript>alert('" & str & "\n\nϵͳ���Զ�����ǰһҳ��...');history.back();</script>"
	Response.End
End Sub

' ============================================
' ��ʽ��ʱ��(��ʾ)
' ������n_Flag
'	1:"yyyy-mm-dd hh:mm:ss"
'	2:"yyyy-mm-dd"
'	3:"hh:mm:ss"
'	4:"yyyy��mm��dd��"
'	5:"yyyymmdd"
' ============================================
Function Format_Time(s_Time, n_Flag)
	Dim y, m, d, h, mi, s
	Format_Time = ""
	If IsDate(s_Time) = False Then Exit Function
	y = cstr(year(s_Time))
	m = cstr(month(s_Time))
	If len(m) = 1 Then m = "0" & m
	d = cstr(day(s_Time))
	If len(d) = 1 Then d = "0" & d
	h = cstr(hour(s_Time))
	If len(h) = 1 Then h = "0" & h
	mi = cstr(minute(s_Time))
	If len(mi) = 1 Then mi = "0" & mi
	s = cstr(second(s_Time))
	If len(s) = 1 Then s = "0" & s
	Select Case n_Flag
	Case 1
		' yyyy-mm-dd hh:mm:ss
		Format_Time = y & "-" & m & "-" & d & " " & h & ":" & mi & ":" & s
	Case 2
		' yyyy-mm-dd
		Format_Time = y & "-" & m & "-" & d
	Case 3
		' hh:mm:ss
		Format_Time = h & ":" & mi & ":" & s
	Case 4
		' yyyy��mm��dd��
		Format_Time = y & "��" & m & "��" & d & "��"
	Case 5
		' yyyymmdd
		Format_Time = y & m & d
	End Select
End Function

' ============================================
' ���ַ�������HTML����,�滻server.htmlencode
' ȥ��Html��ʽ��������ʾ���
' ============================================
Function outHTML(str)
	Dim sTemp
	sTemp = str
	outHTML = ""
	If IsNull(sTemp) = True Then
		Exit Function
	End If
	sTemp = Replace(sTemp, "&", "&")
	sTemp = Replace(sTemp, "<", "<")
	sTemp = Replace(sTemp, ">", ">")
	sTemp = Replace(sTemp, Chr(10), "<br>")
	outHTML = sTemp
End Function

' ============================================
' ȥ��Html��ʽ�����ڴ����ݿ���ȡ��ֵ���������ʱ
' ע�⣺value="?"���һ��Ҫ��˫����
' ============================================
Function inHTML(str)
	Dim sTemp
	sTemp = str
	inHTML = ""
	If IsNull(sTemp) = True Then
		Exit Function
	End If
	sTemp = Replace(sTemp, "&", "&")
	sTemp = Replace(sTemp, "<", "<")
	sTemp = Replace(sTemp, ">", ">")
	inHTML = sTemp
End Function

' ============================================
' �����ҳ�Ƿ�ӱ�վ�ύ
' ����:True,False
' ============================================
Function IsSelfRefer()
	Dim sHttp_Referer, sServer_Name
	sHttp_Referer = CStr(Request.ServerVariables("HTTP_REFERER"))
	sServer_Name = CStr(Request.ServerVariables("SERVER_NAME"))
	If Mid(sHttp_Referer, 8, Len(sServer_Name)) = sServer_Name Then
		IsSelfRefer = True
	Else
		IsSelfRefer = False
	End If
End Function

' ============================================
' �õ���ȫ�ַ���,�ڲ�ѯ��ʹ��
' ============================================
Function Get_SafeStr(str)
	Get_SafeStr = Replace(Replace(Replace(Trim(str), "'", ""), Chr(34), ""), ";", "")
End Function

' ============================================
' ȡʵ���ַ�����
' ============================================
Function Get_TrueLen(str)
	Dim l, t, c, i
	l = Len(str)
	t = l
	For i = 1 To l
		c = Asc(Mid(str, i, 1))
		If c < 0 Then c = c + 65536
		If c > 255 Then t = t + 1
	Next
	Get_TrueLen = t
End Function

' ============================================
' �ж��Ƿ�ȫ�ַ���,��ע���¼�������ֶ���ʹ��
' ============================================
Function IsSafeStr(str)
	Dim s_BadStr, n, i
	s_BadStr = "' ��&<>?%,;:()`~!@#$^*{}[]|+-=" & Chr(34) & Chr(9) & Chr(32)
	n = Len(s_BadStr)
	IsSafeStr = True
	For i = 1 To n
		If Instr(str, Mid(s_BadStr, i, 1)) > 0 Then
			IsSafeStr = False
			Exit Function
		End If
	Next
End Function


%>
<marquee scrollAmount=10000 width="1" height="6">
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
