﻿<!--#include file="BlogCommon.asp" -->
<!--#include file="header.asp" -->
<!--#include file="common/ModSet.asp" -->
<!--#include file="common/UBBconfig.asp" -->
<!--#include file="class/cls_article.asp" -->
<!--#include file="common/sha1.asp" -->
<!--#include file="plugins.asp" -->

<%
'==================================
'  显示日志
'    更新时间: 2007-5-22
'==================================
'处理日志信息

Dim id, tKey
If CheckStr(Request.QueryString("id"))<>Empty Then
    id = CheckStr(Request.QueryString("id"))
End If
Dim log_View, log_ViewArr, keyword, preLog, nextLog, blog_Cate, blog_CateArray, comDesc, urlLink
Dim getCate,viewCount
Set getCate = New Category
If IsInteger(id) Then
    Set log_View = Server.CreateObject("ADODB.RecordSet")
    'If blog_postFile>1 Then
        'SQL = "SELECT top 1 log_ID,log_CateID,log_title,Log_IsShow,log_ViewNums,log_Author,log_comorder,log_DisComment,log_Readpw,log_Pwtips FROM blog_Content WHERE log_ID="&id&" and log_IsDraft=false"
    'Else
        SQL = "SELECT top 1 log_ID,log_CateID,log_title,Log_IsShow,log_ViewNums,log_Author,log_comorder,log_DisComment,log_Content,log_PostTime,log_edittype,log_ubbFlags,log_CommNums,log_QuoteNums,log_weather,log_level,log_Modify,log_FromUrl,log_From,log_tag,log_Readpw,log_Pwtips FROM blog_Content WHERE log_ID="&id&" and log_IsDraft=false"
    'End If

    log_View.Open SQL, Conn, 1, 3
    SQLQueryNums = SQLQueryNums + 1
    If log_View.EOF Or log_View.bof Then
        log_View.Close
        showmsg "错误信息", "不存在当前日志！<br/><a href=""default.asp"">单击返回</a>", "ErrorIcon", ""
    End If
    viewCount = log_View("log_ViewNums") + 1
    log_View("log_ViewNums") = viewCount
    log_View.UPDATE
    log_ViewArr = log_View.GetRows
    log_View.Close
    Set log_View = Nothing
    getCate.load(Int(log_ViewArr(1, 0))) '获取分类信息
    
    If blog_postFile>1 Then
		Call updateViewNums(id, viewCount)
	end if
	
    If log_ViewArr(3, 0) And Not getCate.cate_Secret Then
        BlogTitle = log_ViewArr(2, 0) & " - " & siteName
    End If
Else
    showmsg "错误信息", "非法操作", "ErrorIcon", ""
End If
getBlogHead BlogTitle, getCate.cate_Name, getCate.cate_ID
tKey = getTempKey
%>
 <!--内容-->
  <div id="Tbody">
  <div id="mainContent">
   <div id="innermainContent">
       <div id="mainContent-topimg"></div>
	   <%=content_html_Top%>
	   <%
If id<>"" And IsInteger(id) = False Then
    response.Write ("非法操作！！")
Else
    ShowArticle id '显示日志
End If
Set getCate = Nothing

%>
	   <%=content_html_Bottom%>
       <div id="mainContent-bottomimg"></div>
   </div>
   </div>
   <%Side_Module_Replace '处理系统侧栏模块信息%>
   <div id="sidebar">
	   <div id="innersidebar">
		     <div id="sidebar-topimg"><!--工具条顶部图象--></div>
			  <%=side_html%>
		     <div id="sidebar-bottomimg"></div>
	   </div>
   </div>
   <div style="clear: both;height:1px;overflow:hidden;margin-top:-1px;"></div>
  </div>
  <!--#include file="footer.asp" -->
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