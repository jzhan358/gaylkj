<!--#include file = "Startup.asp"-->
<!--#include file = "../../Include/DeCode.asp"-->

<%

' ======================
' ���ܣ���ʾ����
' ��������ʾ�༭������ҳ����ҳע��һ��DeCode�ӿں����ĵ��á�
' ======================

Call Header("��ʾ��������")
Call Content()
Call Footer()


' ��ҳ������
Sub Content()

	' �������������ID
	Dim sNewsID
	sNewsID = Trim(Request("id"))

	' ����ID��Ч����֤����ֹ��Щ�˶�����ƻ�����ʾ����
	If IsNumeric(sNewsID) = False Then
		GoError "��ͨ��ҳ���ϵ����ӽ��в�������Ҫ��ͼ�ƻ�����ʾϵͳ��"
	End If

	' �����ݿ���ȡ��ʼֵ
	Dim sTitle, sContent, sPicture, sOriginalFileName, sSaveFileName, sSavePathFileName
	sSql = "SELECT * FROM NewsData WHERE D_ID=" & sNewsID
	oRs.Open sSql, oConn, 0, 1
	If Not oRs.Eof Then
		sTitle = oRs("D_Title")
		sContent = oRs("D_Content")
		sPicture = oRs("D_Picture")
		sOriginalFileName = oRs("D_OriginalFileName")
		sSaveFileName = oRs("D_SaveFileName")
		sSavePathFileName = oRs("D_SavePathFileName")
	Else
		GoError "��Ч������ID�����ҳ���ϵ����ӽ��в�����"
	End If
	oRs.Close

	' ����ĳЩ��ǩ������ڰ�ȫ���ǵ�Script��ǩ����
	' Ҫʹ�ô˹�����Ҫ�Ȱ���"Include/DeCode.asp"�ļ���
	' ����ֻ����SCRIPT��ǩ������ζ�������еĿͻ��˽ű�������Ч�����ɸ���ʵ�ʵ���Ҫ����������ǩ��
	' ��ǰ֧�ֹ��˵ı�ǩ�����Բ鿴DeCode.asp�ļ��е�˵����
	sContent = eWebEditor_DeCode(sContent, "SCRIPT")


	' �������
	Response.Write "<table border=0 cellpadding=5 width='90%' align=center>" & _
		"<tr><td align=center><b>" & sTitle & "</b></td></tr>" & _
		"<tr><td>" & sContent & "</td></tr>" & _
		"</table>"

	' �������ļ���Ϣ
	Response.Write "<p><b>�����ŵ�����ϴ��ļ���Ϣ���£�</b></p>"

	' �Ѵ�"|"���ַ���תΪ���飬�����г���ʾ
	Dim aOriginalFileName, aSaveFileName, aSavePathFileName
	aOriginalFileName = Split(sOriginalFileName, "|")
	aSaveFileName = Split(sSaveFileName, "|")
	aSavePathFileName = Split(sSavePathFileName, "|")

	Response.Write "<table border=1 cellpadding=3 cellspacing=0>" & _
		"<tr>" & _
			"<td>���</td>" & _
			"<td>ԭ�ļ���(�ӿڣ�d_originalfilename)</td>" & _
			"<td>�����ļ���(�ӿڣ�d_savefilename)</td>" & _
			"<td>����·���ļ���(�ӿڣ�d_savepathfilename)</td>" & _
		"</tr>"
	Dim i
	For i = 0 To UBound(aOriginalFileName)
		Response.Write "<tr>" & _
				"<td>" & CStr(i + 1) & "</td>" & _
				"<td>" & aOriginalFileName(i) & "</td>" & _
				"<td>" & aSaveFileName(i) & "</td>" & _
				"<td>" & aSavePathFileName(i) & "</td>" & _
			"</tr>"
	Next
	Response.Write "</table>"

End Sub

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
