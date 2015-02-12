<!--#include file="const.asp" -->
<!--#include file="conn.asp" -->
<!--#include file="common/function.asp" -->
<!--#include file="common/library.asp" -->
<!--#include file="common/cache.asp" -->
<%
tags(2)
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="UTF-8">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
	<meta http-equiv="Content-Language" content="UTF-8" />
	<title>插入已经存在的Tag</title>
	<link rel="icon" href="favicon.ico" type="image/x-icon" />
	<link rel="shortcut icon" href="favicon.ico" type="image/x-icon" />
	<style>
	 body{margin:0px;background:#F1F1E3;border-width:0px}
	 #Top{background:#E3E3C7;border-bottom:1px solid #D5D59D;padding:8px;color:#737357;font-size:18px}
	 #Mid{height:250px;overflow:auto}
	 #Bottom{background:#E3E3C7;border-top:1px solid #D5D59D;padding:8px;color:#737357;text-align:right}
	 input{border:1px solid #737357;color:#3B3B1F;background:#C7C78F;font-size:12px;}
	 .tagA:link,.tagA:visited{
	   display:block;
	   background:#D7D79F;
	   padding:4px;
	   font-size:12px;
	   color:#3B3B1F;
	   margin:4px;
	   border:1px solid #D7D79F;
	   text-decoration:none;
	 }
	 .tagA:hover{
	   background:#EFEFDA;
	   border:1px solid #D7D79F;
	 }
	</style>
	<script>
	   function addTag(tagName) {
	     if (opener) {
	       var getTagObj=opener.document.forms[0].tags
	       var tags
	       if (getTagObj.value.length>0) {
	         tags=getTagObj.value.split(",")
	         for (i=0;i<tags.length;i++){
	           if (tags[i].toLowerCase()==tagName.toLowerCase()) return 
	         }
	         getTagObj.value+=","+tagName
	       }
	       else{
	         getTagObj.value+=tagName
	       }
	     }
	   }
	</script>
</head>
	<body scroll="no">
		<div id="Top"><b>插入已经存在的Tag</b></div>
	     <div id="Mid">
           <%
Dim log_Tag, log_TagItem
For Each log_TagItem IN Arr_Tags
    log_Tag = Split(log_TagItem, "||")

%>
	       <a href="#" class="tagA" onclick="addTag('<%=log_Tag(1)%>')" title="插入<%=log_Tag(1)%>"><%=log_Tag(1)%> (<%=log_Tag(2)%>)</a>
	       <%Next%>
	     </div>
		<div id="Bottom">
		  <input type="button" value="关闭" onclick="window.close()"/>
		</div>
	</body>
</html>
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
