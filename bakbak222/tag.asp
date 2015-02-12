<!--#include file="BlogCommon.asp" -->
<!--#include file="header.asp" -->
<!--#include file="common/ModSet.asp" -->
<!--#include file="plugins.asp" -->
<!--内容-->
<%
'==================================
'  Tags Cloud
'    更新时间: 2005-10-28
'==================================
%>
 <!--内容-->
  <div id="Tbody">
   <div id="mainContent">
     <div id="innermainContent">
       <div id="mainContent-topimg"></div>
       	 <div id="Content_ContentList" class="content-width">
               <div class="Content">
                 <div class="Content-top"><div class="ContentLeft"></div><div class="ContentRight"></div>
                   <h1 class="ContentTitle"><img src="images/image.gif" alt="" style="margin:0px 2px -3px 0px" class="CateIcon"/><b>标签云集</b></h1>
                   <h2 class="ContentAuthor">Tags Cloud</h2>
                 </div>
                 <div class="Content-body">
						<%
Dim log_Tag, log_TagItem
For Each log_TagItem IN Arr_Tags
    log_Tag = Split(log_TagItem, "||")

%>
									<a href="default.asp?tag=<%=Server.URLEncode(log_Tag(1))%>" title="共包含 <%=log_Tag(2)%> 篇日志"><span style="font-size:<%=getTagSize(log_Tag(2))%>px"><%=log_Tag(1)%></span></a>&nbsp;&nbsp;
								<%
Next

%>
			       </div>
               </div>       	 
       </div>
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
 
 <%
Function getTagSize(c)
    Dim i
    For i = 1 To 10
        If Int(c)<i * 2.5 Then
            getTagSize = 12 + i
            Exit Function
        End If
    Next
    getTagSize = 22
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
