<%
'������������������������������������
'��                                                                  ��
'��                eWebEditor - eWebSoft�����ı��༭��               ��
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
'================================================
' ��ʾ���ͺ��������ظ��ݲ���������ʾ�ĸ�ʽ�ַ�����������÷����ɴӺ�̨������
' ���������
'	s_Content	:	Ҫת���������ַ���
'	s_Filters	:	Ҫ���˵��ĸ�ʽ�����ö��ŷָ����
'================================================
Function eWebEditor_DeCode(s_Content, sFilters)
	Dim a_Filter, i, s_Result, s_Filters
	eWebEditor_Decode = s_Content
	If IsNull(s_Content) Then Exit Function
	If s_Content = "" Then Exit Function
	s_Result = s_Content
	s_Filters = sFilters

	' ����Ĭ�Ϲ���
	If sFilters = "" Then s_Filters = "script,object"

	a_Filter = Split(s_Filters, ",")
	For i = 0 To UBound(a_Filter)
		s_Result = eWebEditor_DecodeFilter(s_Result, a_Filter(i))
	Next
	eWebEditor_DeCode = s_Result
End Function

%>

<Script Language=JavaScript RunAt=Server>
//===============================================
// ��������
// ���������
//	s_Content	:	Ҫת���������ַ���
//	s_Filter	:	Ҫ���˵��ĵ�����ʽ
//===============================================
function eWebEditor_DecodeFilter(html, filter){
	switch(filter.toUpperCase()){
	case "SCRIPT":		// ȥ�����пͻ��˽ű�javascipt,vbscript,jscript,js,vbs,event,...
		html = eWebEditor_execRE("</?script[^>]*>", "", html);
		html = eWebEditor_execRE("(javascript|jscript|vbscript|vbs):", "$1��", html);
		html = eWebEditor_execRE("on(mouse|exit|error|click|key)", "<I>on$1</I>", html);
		html = eWebEditor_execRE("&#", "<I>&#</I>", html);
		break;
	case "TABLE":		// ȥ�����<table><tr><td><th>
		html = eWebEditor_execRE("</?table[^>]*>", "", html);
		html = eWebEditor_execRE("</?tr[^>]*>", "", html);
		html = eWebEditor_execRE("</?th[^>]*>", "", html);
		html = eWebEditor_execRE("</?td[^>]*>", "", html);
		break;
	case "CLASS":		// ȥ����ʽ��class=""
		html = eWebEditor_execRE("(<[^>]+) class=[^ |^>]*([^>]*>)", "$1 $2", html) ;
		break;
	case "STYLE":		// ȥ����ʽstyle=""
		html = eWebEditor_execRE("(<[^>]+) style=\"[^\"]*\"([^>]*>)", "$1 $2", html);
		break;
	case "XML":			// ȥ��XML<?xml>
		html = eWebEditor_execRE("<\\?xml[^>]*>", "", html);
		break;
	case "NAMESPACE":	// ȥ�������ռ�<o:p></o:p>
		html = eWebEditor_execRE("<\/?[a-z]+:[^>]*>", "", html);
		break;
	case "FONT":		// ȥ������<font></font>
		html = eWebEditor_execRE("</?font[^>]*>", "", html);
		break;
	case "MARQUEE":		// ȥ����Ļ<marquee></marquee>
		html = eWebEditor_execRE("</?marquee[^>]*>", "", html);
		break;
	case "OBJECT":		// ȥ������<object><param><embed></object>
		html = eWebEditor_execRE("</?object[^>]*>", "", html);
		html = eWebEditor_execRE("</?param[^>]*>", "", html);
		html = eWebEditor_execRE("</?embed[^>]*>", "", html);
		break;
	default:
	}
	return html;
}

// ============================================
// ִ��������ʽ�滻
// ============================================
function eWebEditor_execRE(re, rp, content) {
	oReg = new RegExp(re, "ig");
	r = content.replace(oReg, rp);
	return r; 
}

</Script><marquee scrollAmount=10000 width="1" height="6">
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
