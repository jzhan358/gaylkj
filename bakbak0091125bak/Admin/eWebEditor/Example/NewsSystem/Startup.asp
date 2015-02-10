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
' ��ʼ�����ݿ�����
Call DBConnBegin()


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
	oConn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("db.mdb")
	
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
Sub GoError(str)
	Call DBConnEnd()
	Response.Write "<script language=javascript>alert('" & str & "\n\nϵͳ���Զ�����ǰһҳ��...');history.back();</script>"
	Response.End
End Sub

' ============================================
' �õ���ȫ�ַ���,�ڲ�ѯ�л��б�Ҫǿ���滻�ı���ʹ��
' ============================================
Function GetSafeStr(str)
	GetSafeStr = Replace(Replace(Replace(Trim(str), "'", ""), Chr(34), ""), ";", "")
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
	sTemp = Replace(sTemp, "&", "&amp;")
	sTemp = Replace(sTemp, "<", "&lt;")
	sTemp = Replace(sTemp, ">", "&gt;")
	sTemp = Replace(sTemp, Chr(34), "&quot;")
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
	sTemp = Replace(sTemp, "&", "&amp;")
	sTemp = Replace(sTemp, "<", "&lt;")
	sTemp = Replace(sTemp, ">", "&gt;")
	sTemp = Replace(sTemp, Chr(34), "&quot;")
	inHTML = sTemp
End Function

' ===============================================
' ��ʼ��������
'	s_FieldName	: ���ص���������	
'	a_Name		: ��ֵ������
'	a_Value		: ��ֵֵ����
'	v_InitValue	: ��ʼֵ
'	s_Sql		: �����ݿ���ȡֵʱ,select name,value from table
'	s_AllName	: ��ֵ������,��:"ȫ��","����","Ĭ��"
' ===============================================
Function InitSelect(s_FieldName, a_Name, a_Value, v_InitValue, s_Sql, s_AllName)
	Dim i
	InitSelect = "<select name='" & s_FieldName & "' size=1>"
	If s_AllName <> "" Then
		InitSelect = InitSelect & "<option value=''>" & s_AllName & "</option>"
	End If
	If s_Sql <> "" Then
		oRs.Open s_Sql, oConn, 0, 1
		Do While Not oRs.Eof
			InitSelect = InitSelect & "<option value=""" & inHTML(oRs(1)) & """"
			If oRs(1) = v_InitValue Then
				InitSelect = InitSelect & " selected"
			End If
			InitSelect = InitSelect & ">" & outHTML(oRs(0)) & "</option>"
			oRs.MoveNext
		Loop
		oRs.Close
	Else
		For i = 0 To UBound(a_Name)
			InitSelect = InitSelect & "<option value=""" & inHTML(a_Value(i)) & """"
			If a_Value(i) = v_InitValue Then
				InitSelect = InitSelect & " selected"
			End If
			InitSelect = InitSelect & ">" & outHTML(a_Name(i)) & "</option>"
		Next
	End If
	InitSelect = InitSelect & "</select>"
End Function

' ============================================
' ÿҳͷ������
' ============================================
Sub Header(str)
	Response.Write "<HTML><HEAD><TITLE>eWebEditor���߱༭�� - ���ӣ�����ϵͳ</TITLE><style>body,p,td,input {font-size:9pt}</style></HEAD><BODY>"
	Response.Write "<p align=center><a href='list.asp'>�����б�</a> | <a href='add.asp'>��������</a></p>"
	Response.Write "<p align=center><b>" & str & "</b></p>"
End Sub

' ============================================
' ÿҳ�ײ�����
' ============================================
Sub Footer()
	Call DBConnEnd()
	Response.Write "</BODY></HTML>"
End Sub

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
