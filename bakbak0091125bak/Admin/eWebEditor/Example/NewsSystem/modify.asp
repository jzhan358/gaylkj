<!--#include file = "Startup.asp"-->

<%

' ======================
' ���ܣ��޸�����
' �������ṩһ���޸ı����������ű�����������ݣ���������ʹ��eWebEditor���б༭��
'       ͬʱ�����ϴ��ļ����Ա�ɾ������ʱ��ͬʱɾ���ϴ��ļ���
'       ���ɱ༭�����ϴ����ļ����ṩ�������ŵ�ͼƬѡ��
'		������������ͬ����ͬ����ֻ���޸ı����г�ʼֵ��
' ======================

Call Header("�޸�����")
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

	' �Ѵ�"|"���ַ���תΪ���飬���ڳ�ʼ�������
	Dim aSavePathFileName
	aSavePathFileName = Split(sSavePathFileName, "|")

	' ������������������ϴ��ļ�Ϊ���ݣ���ʼ�������
	' ����InitSelect����������ֵ����ʼֵ��������������ִ����������startup.asp�ļ��к�����˵������
	Dim sOptionSavePath
	sOptionSavePath = InitSelect("d_picture", aSavePathFileName, aSavePathFileName, sPicture, "", "��")

	%>

	<Script Language=JavaScript>
	// ���ϴ�ͼƬ���ļ�ʱ����������������ͼƬ·�����ɸ���ʵ����Ҫ���Ĵ˺���
	function doChange(objText, objDrop){
		if (!objDrop) return;
		var str = objText.value;
		var arr = str.split("|");
		var nIndex = objDrop.selectedIndex;
		objDrop.length=1;
		for (var i=0; i<arr.length; i++){
			objDrop.options[objDrop.length] = new Option(arr[i], arr[i]);
		}
		objDrop.selectedIndex = nIndex;
	}

	// ���ύ�ͻ��˼��
	function doSubmit(){
		if (document.myform.d_title.value==""){
			alert("���ű��ⲻ��Ϊ�գ�");
			return false;
		}
		// getHTML()ΪeWebEditor�Դ��Ľӿں���������Ϊȡ�༭��������
		if (eWebEditor1.getHTML()==""){
			alert("�������ݲ���Ϊ�գ�");
			return false;
		}
		document.myform.submit();
	}
	</Script>
	
	<form action="modifysave.asp?id=<%=sNewsID%>" method="post" name="myform">
	<% 'ȡԴ�ļ��� %>
	<input type=hidden name=d_originalfilename value="<%=sOriginalFileName%>">
	<% 'ȡ����ķ���������·�����������Ҫ��·������������򣬿���������ı������onchange�¼� %>
	<input type=hidden name=d_savefilename value="<%=sSaveFileName%>">
	<% 'ȡ������ļ�����ʹ�ô�·������������� %>
	<input type=hidden name=d_savepathfilename onchange="doChange(this,document.myform.d_picture)" value="<%=sSavePathFileName%>">

	<table cellspacing=3 align=center>
	<tr>
		<td>���ű��⣺</td>
		<% ' ����inHTML��ȥ��Html��ʽ�����ڴ����ݿ���ȡ��ֵ���������ʱ���������startup.asp�ļ��к�����˵�� %>
		<td><input type="text" name="d_title" value="<%=inHTML(sTitle)%>" size="90"></td>
	</tr>
	<tr>
		<td>����ͼƬ��</td>
		<td><%=sOptionSavePath%> ���༭���в���ͼƬʱ�����Զ�����������</td>
	</tr>
	<tr>
		<td>�������ݣ�</td>
		<td>
			<%
			' ewebeditor.asp�ļ����õĲ�����
			'	id���������textarea�����ƣ��ڴ˱�����d_content��ע���Сд
			'	style���༭������ʽ���ƣ�����eWebEditor�ĺ�̨����
			'	originalfilename�����ڻ�ȡԴ�ļ����ı��������ڴ˱�����d_originalfilename
			'	savefilename�����ڻ�ȡ�����ļ����ı��������ڴ˱�����d_savefilename
			'	savepathfilename�����ڻ�ȡ�����·���ļ����ı��������ڴ˱�����d_savepathfilename

			' ��һ��Ҫ��Server.HtmlEncode()���������ݽ���ת���������п��Գ���ĳЩ���ݳ��˱༭�������
			%>
			<textarea name="d_content" style="display:none"><%=Server.HtmlEncode(sContent)%></textarea>
			<iframe ID="eWebEditor1" src="../../ewebeditor.asp?id=d_content&style=s_newssystem&originalfilename=d_originalfilename&savefilename=d_savefilename&savepathfilename=d_savepathfilename" frameborder="0" scrolling="no" width="550" HEIGHT="350"></iframe>
		</td>
	</tr>
	</table>
	<p align=center><input type=button name=btnSubmit value=" �� �� " onclick="doSubmit()"> <input type=reset name=btnReset value=" �� �� "></p>
	</form>

	<%
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
