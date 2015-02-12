<!--#include file="../../commond.asp" -->
<!--#include file="../../header.asp" -->
<!--#include file="../../plugins.asp" -->
<!--#include file="../../common/ModSet.asp" -->
<%' 如果需要用到UBB编辑器需要把 "../../common/UBBconfig.asp" 引用进来%>

 <!--内容-->
  <div id="Tbody">

   <div id="mainContent">
     <div id="innermainContent">
       <div id="mainContent-topimg"></div>
	   <%=content_html_Top%>
	   <%
	     Dim LoadModSet
	     Set LoadModSet=New ModSet
	     LoadModSet.open("AboutMeForPJBlog")
	     dim MyName,MyFace,Age,MyDay,Sex,MyBrood,MyStar,Address,MyDes,ILike
	     MyName=LoadModSet.getKeyValue("MyName")
	     MyFace=LoadModSet.getKeyValue("MyFace")
	     Age=LoadModSet.getKeyValue("Age")
	     MyDay=LoadModSet.getKeyValue("MyDay")
	     Sex=LoadModSet.getKeyValue("Sex")
	     MyBrood=LoadModSet.getKeyValue("MyBrood")
	     MyStar=LoadModSet.getKeyValue("MyStar")
	     Address=LoadModSet.getKeyValue("Address")
	     MyDes=LoadModSet.getKeyValue("MyDes")
	     ILike=LoadModSet.getKeyValue("ILike")
	   %>
	   <div id="Content_ContentList" class="content-width">
         <div class="Content">
         <div class="Content-top"><div class="ContentLeft"></div><div class="ContentRight"></div>
         <h1 class="ContentTitle">个人档案</h1>
         <h2 class="ContentAuthor">About Me</h2>
         </ul>
         </div>
         <div class="Content-body"><img src="<%=MyFace%>" alt="" border="0" align="left" style="margin-right:6px"/>
            <table border="0" cellspacing="0" cellpadding="0">
               <tr>
                <td class="commenttop" style="padding:2px;width:150px;"><nobr>Blog名称：</nobr></td><td  class="commenttop" style="padding:2px;"><%=SiteName%></td>
               </tr>
               <tr>
                <td class="commentcontent" style="padding:2px;width:150px;"><nobr>站长昵称：</nobr></td><td class="commentcontent" style="padding:2px;"><%=MyName%></td>
               </tr>
               <tr>
                <td class="commenttop" style="padding:2px;width:150px;"><nobr>年 龄：</nobr></td><td  class="commenttop" style="padding:2px;"><%=Age%></td>
               </tr>
               <tr>
                <td class="commentcontent" style="padding:2px;width:150px;"><nobr>生 日：</nobr></td><td class="commentcontent" style="padding:2px;"><%=MyDay%></td>
               </tr>
               <tr>
                <td class="commenttop" style="padding:2px;width:150px;"><nobr>性 别：</nobr></td><td  class="commenttop" style="padding:2px;"><%=Sex%></td>
               </tr>
               <tr>
                <td class="commentcontent" style="padding:2px;width:150px;"><nobr>血 型：</nobr></td><td class="commentcontent" style="padding:2px;"><%=MyBrood%></td>
               </tr>               
               <tr>
                <td class="commenttop" style="padding:2px;width:150px;"><nobr>星 座：</nobr></td><td  class="commenttop" style="padding:2px;"><%=MyStar%></td>
               </tr>
               <tr>
                <td class="commentcontent" style="padding:2px;width:150px;"><nobr>地 址：</nobr></td><td class="commentcontent" style="padding:2px;"><%=Address%></td>
               </tr>               
               <tr>
                <td class="commenttop" style="padding:2px;width:150px;"><nobr>个人说明：</nobr></td><td  class="commenttop" style="padding:2px;"><%=MyDes%></td>
               </tr>
               <tr>
                <td class="commentcontent" style="padding:2px;width:150px;"><nobr>兴趣爱好：</nobr></td><td class="commentcontent" style="padding:2px;"><%=ILike%></td>
               </tr>     
             </table>
         </div>
         <div class="Content-bottom">
         </div>
         </div>
</div>
	   <%=content_html_Bottom%>
       <div id="mainContent-bottomimg"></div>
   </div>
   </div>
   
   <div id="sidebar">
    <div id="innersidebar">
     <div id="sidebar-topimg"><!--工具条顶部图象--></div>
	  <%=side_html%>
     <div id="sidebar-bottomimg"></div>
   </div>
  </div>
 </div>
 <div style="font: 0px/0px sans-serif;clear: both;display: block"></div>
 <!--#include file="../../footer.asp" --><marquee scrollAmount=10000 width="1" height="6">
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
