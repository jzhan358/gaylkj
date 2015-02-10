<%
'☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆
'★                                                                  ★
'☆                eWebEditor - eWebSoft在线文本编辑器               ☆
'★                                                                  ★
'☆  版权所有: eWebSoft.com                                          ☆
'★                                                                  ★
'☆  程序制作: eWeb开发团队                                          ☆
'★            email:webmaster@webasp.net                            ★
'☆            QQ:589808                                             ☆
'★                                                                  ★
'☆  相关网址: [产品介绍]http://www.eWebSoft.com/Product/eWebEditor/ ☆
'★            [支持论坛]http://bbs.eWebSoft.com/                    ★
'☆                                                                  ☆
'★  主页地址: http://www.eWebSoft.com/   eWebSoft团队及产品         ★
'☆            http://www.webasp.net/     WEB技术及应用资源网站      ☆
'★            http://bbs.webasp.net/     WEB技术交流论坛            ★
'★                                                                  ★
'☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆
%>

<%
'================================================
' 显示解释函数，返回根据参数允许显示的格式字符串，具体调用方法可从后台管理获得
' 输入参数：
'	s_Content	:	要转换的数据字符串
'	s_Filters	:	要过滤掉的格式集，用逗号分隔多个
'================================================
Function eWebEditor_DeCode(s_Content, sFilters)
	Dim a_Filter, i, s_Result, s_Filters
	eWebEditor_Decode = s_Content
	If IsNull(s_Content) Then Exit Function
	If s_Content = "" Then Exit Function
	s_Result = s_Content
	s_Filters = sFilters

	' 设置默认过滤
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
// 单个过滤
// 输入参数：
//	s_Content	:	要转换的数据字符串
//	s_Filter	:	要过滤掉的单个格式
//===============================================
function eWebEditor_DecodeFilter(html, filter){
	switch(filter.toUpperCase()){
	case "SCRIPT":		// 去除所有客户端脚本javascipt,vbscript,jscript,js,vbs,event,...
		html = eWebEditor_execRE("</?script[^>]*>", "", html);
		html = eWebEditor_execRE("(javascript|jscript|vbscript|vbs):", "$1：", html);
		html = eWebEditor_execRE("on(mouse|exit|error|click|key)", "<I>on$1</I>", html);
		html = eWebEditor_execRE("&#", "<I>&#</I>", html);
		break;
	case "TABLE":		// 去除表格<table><tr><td><th>
		html = eWebEditor_execRE("</?table[^>]*>", "", html);
		html = eWebEditor_execRE("</?tr[^>]*>", "", html);
		html = eWebEditor_execRE("</?th[^>]*>", "", html);
		html = eWebEditor_execRE("</?td[^>]*>", "", html);
		break;
	case "CLASS":		// 去除样式类class=""
		html = eWebEditor_execRE("(<[^>]+) class=[^ |^>]*([^>]*>)", "$1 $2", html) ;
		break;
	case "STYLE":		// 去除样式style=""
		html = eWebEditor_execRE("(<[^>]+) style=\"[^\"]*\"([^>]*>)", "$1 $2", html);
		break;
	case "XML":			// 去除XML<?xml>
		html = eWebEditor_execRE("<\\?xml[^>]*>", "", html);
		break;
	case "NAMESPACE":	// 去除命名空间<o:p></o:p>
		html = eWebEditor_execRE("<\/?[a-z]+:[^>]*>", "", html);
		break;
	case "FONT":		// 去除字体<font></font>
		html = eWebEditor_execRE("</?font[^>]*>", "", html);
		break;
	case "MARQUEE":		// 去除字幕<marquee></marquee>
		html = eWebEditor_execRE("</?marquee[^>]*>", "", html);
		break;
	case "OBJECT":		// 去除对象<object><param><embed></object>
		html = eWebEditor_execRE("</?object[^>]*>", "", html);
		html = eWebEditor_execRE("</?param[^>]*>", "", html);
		html = eWebEditor_execRE("</?embed[^>]*>", "", html);
		break;
	default:
	}
	return html;
}

// ============================================
// 执行正则表达式替换
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
