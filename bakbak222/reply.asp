<!--#include file="BlogCommon.asp" -->
﻿<!--#include file="common/function.asp" -->
<!--#include file="common/library.asp" -->
<!--#include file="common/cache.asp" -->
<!--#include file="common/checkUser.asp" -->
<!--#include file="common/ModSet.asp" -->
<!--#include file="class/cls_article.asp" -->
<!--#include file="class/cls_logAction.asp" -->
<!--#include file="common/ubbcode.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="UTF-8">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />

<%
Sub replyComm
	dim cID,replay,quest,result,a1,a2,a3,logId
	logId = CheckStr(Request.Form("logId"))
	cID = CheckStr(Request.Form("id"))
	a1 = Int(Request.Form("a1"))
	a2 = Int(Request.Form("a2"))
	a3 = Int(Request.Form("a3"))
	
	replay = CheckStr(Request.Form("replay"))
	
	if isEmpty(memName) Or stat_Admin <> True then
		response.write ("<script>parent.removeReplyMsg("&cID&");alert('您没有权限回复评论。')</script>")
		exit Sub
	end if 
	

	
	if isEmpty(logId) or isEmpty(cID) or isEmpty(replay) then
		response.write ("<script>parent.removeReplyMsg("&cID&");alert('您没有权限回复评论。')</script>")
		exit sub
	end if
	
	if isEmpty(replay) then
		response.write ("<script>parent.removeReplyMsg("&cID&");alert('您没有输入评论内容')</script>")
		exit sub
	end if
	
	set quest = Conn.Execute("select top 1 comm_Content from blog_Comment where comm_ID=" & cID)
	
 	If not quest.EOF Then
		result = quest(0) & vbcrlf & "[reply=" + memName + "," & DateToStr(now(),"Y-m-d H:I A") & "]" & replay & "[/reply]"
  		conn.Execute("UPDATE blog_Comment SET comm_Content='"&result&"' WHERE comm_ID="&cID)
  		
  		dim ubbResult
  		 ubbResult  = UBBCode(HtmlEncode(result),cBool(a1),blog_commUBB,blog_commIMG,cBool(a2),cBool(a3))
		response.write ("<script>parent.$(""commcontent_"&cID&""").innerHTML = '"&output(ubbResult)&"';</script>")
		
		 PostArticle logId, False
    End If

	quest.close
	set quest = nothing
End Sub

call replyComm
%>
<script runat="server" Language="javascript">
	function output(html){
		html = html.replace(/[\n\r]/g,"");
		html = html.replace(/\\/g,"\\\\");
		html = html.replace(/\'/g,"\\'");
		return html
	}
</script>
</head>
<body></body></html>
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
