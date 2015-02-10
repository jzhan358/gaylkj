<!--#include file = "Startup.asp"-->
<!--#include file = "../../Include/DeCode.asp"-->

<%

' ======================
' 功能：显示新闻
' 描述：显示编辑的内容页，此页注意一下DeCode接口函数的调用。
' ======================

Call Header("显示新闻内容")
Call Content()
Call Footer()


' 本页内容区
Sub Content()

	' 传入参数：新闻ID
	Dim sNewsID
	sNewsID = Trim(Request("id"))

	' 新闻ID有效性验证，防止有些人恶意的破坏此演示程序
	If IsNumeric(sNewsID) = False Then
		GoError "请通过页面上的链接进行操作，不要试图破坏此演示系统。"
	End If

	' 从数据库中取初始值
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
		GoError "无效的新闻ID，请点页面上的链接进行操作！"
	End If
	oRs.Close

	' 禁用某些标签，如出于安全考虑的Script标签，等
	' 要使用此功能需要先包含"Include/DeCode.asp"文件。
	' 此例只过滤SCRIPT标签，即意味着内容中的客户端脚本不会生效，您可根据实际的需要加入其它标签。
	' 当前支持过滤的标签，可以查看DeCode.asp文件中的说明。
	sContent = eWebEditor_DeCode(sContent, "SCRIPT")


	' 输出新闻
	Response.Write "<table border=0 cellpadding=5 width='90%' align=center>" & _
		"<tr><td align=center><b>" & sTitle & "</b></td></tr>" & _
		"<tr><td>" & sContent & "</td></tr>" & _
		"</table>"

	' 输出相关文件信息
	Response.Write "<p><b>此新闻的相关上传文件信息如下：</b></p>"

	' 把带"|"的字符串转为数组，用于列出显示
	Dim aOriginalFileName, aSaveFileName, aSavePathFileName
	aOriginalFileName = Split(sOriginalFileName, "|")
	aSaveFileName = Split(sSaveFileName, "|")
	aSavePathFileName = Split(sSavePathFileName, "|")

	Response.Write "<table border=1 cellpadding=3 cellspacing=0>" & _
		"<tr>" & _
			"<td>序号</td>" & _
			"<td>原文件名(接口：d_originalfilename)</td>" & _
			"<td>保存文件名(接口：d_savefilename)</td>" & _
			"<td>保存路径文件名(接口：d_savepathfilename)</td>" & _
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
