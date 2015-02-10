<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<HTML>
<HEAD>
<TITLE>广州游戏机网首页 广州游戏机批发零售 广州游戏机批发零售 广州游戏机在线购买 游戏机厂家直销 专业诚信的供应各类游戏机批发零售</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<META content="MSHTML 6.00.2900.2802" name=GENERATOR>
<meta name="Keywords" content="广州游戏机供应,广州麻将机供应,广州益智类游戏机供应,广州游艺机供应,广州水果机供应,广州苹果机供应,广州玛丽机供应,广州连线机供应,广州轮盘供应,广州777机供应,广州37机供应,广州弹珠机供应,广州百家乐供应,广州模拟机供应,广州亲子游戏机供应,广州娃娃机供应,广州摇摆机供应,广州框体机供应,广州3D动物供应,捞金鱼,SNK卡带,SNK卡座,游戏电脑板,框体节目版" />
<meta name="Description" content="广州市广安娱乐主要是一家从事多年游戏机生产、销售、电脑游戏机配件，集科、工、贸于一体的专业公司，并与台湾资深电玩工程师，合作开发多款新产品。" />
<LINK 
href="images/css.css" type=text/css rel=stylesheet>
</HEAD>
<BODY leftMargin=0  topMargin=0 marginheight="0" marginwidth="0" background="images/bg.jpg">
<table width="778" border="0" align="center" cellpadding="0" cellspacing="0" class="rl-orange">
  <tr>
    <td>
	<!--#include file="Top.asp" -->
	<table width="778"  border="0" align="center" cellpadding="0" cellspacing="0">
 	 <tr>
  	  <td width="195" valign="top" bgcolor="C0EDB4" class="r-green">
<!--#include file="left.asp" -->
		</td>
	<!--内容开始-->
    <td valign="top">
		<table width="100%"  border="0" cellspacing="0" cellpadding="0">
		  <tr>
			<td><img src="images/main_pro1.jpg" width="565" height="45"></td>
		  </tr>
		  <tr>
			<td>
			<!--start-->
			 <TABLE style="FONT-SIZE: 9pt; LINE-HEIGHT: 150%" cellSpacing=0 cellPadding=0 width=565 border=0>
							<TR>
							  <td width="45">&nbsp;</td><TD width="520">&nbsp;</TD>                     
							</TR>	
							<!--企业产品列表-->
							
							<TR align=center>
							<td>&nbsp;</td>
							<%
							set rs=server.CreateObject("adodb.recordset")
							rs.open "select top 15 * from p_info order by p_order desc",conn,1,1
							n=1						
							do while not rs.eof and n<=15
							fileExt=lcase(rs("p_Pic"))							
							%>						
							  <TD vAlign=top height=142>
								<TABLE cellSpacing=0 cellPadding=0 width=100 border=0>
								  <TR>
									<TD><A href="product_view.asp?id=<%=rs("p_id")%>" target="_blank"><IMG style="border:1px solid #C3CBD4" height=110 src="<%=fileExt%>" width=160 border=0></A></TD>
									<TD vAlign=bottom rowSpan=2><IMG height=108 src="Images/JJTK_21.gif" width=11></TD>
								  </TR>
								  <TR>
									<TD align=left bgColor=#c3cbd4><IMG height=2 src="" width=2></TD>
								  </TR>
								  <TR>
									<TD align=center height=24><A class=reg_Text href="product_view.asp?id=<%=rs("p_id")%>" target="_blank"><%=left(rs("p_name"),10)%></A></TD>
									<TD vAlign=bottom></TD>
								  </TR>
							  </TABLE></TD>
							  <%
								
								if n mod 3 =0 then response.Write("</tr><tr align=center><td>&nbsp;</td>")
								n=n+1
								rs.movenext
								loop
								rs.close
							  %>
							  </TR></Table>  
			<!--end-->
			<table width="100%"  border="0" cellspacing="0" cellpadding="0">
		  <tr>
			<td><img src="images/main_pro.jpg" width="565" height="45"></td>
		  </tr>
		  <tr>
			<td height="200">
			<!--最新图文代码开始-->
<script language="JavaScript"> 
<!-- 
var flag=false; 
function DrawImage(ImgD){ 
 var image=new Image(); 
 image.src=ImgD.src; 
 if(image.width>0 && image.height>0){ 
  flag=true; 
  if(image.width/image.height>= 165/125){ 
   if(image.width>165){
    ImgD.width=165; 
    ImgD.height=(image.height*165)/image.width; 
   }else{ 
    ImgD.width=image.width;
    ImgD.height=image.height; 
   } 
   
  } 
  else{ 
   if(image.height>125){
    ImgD.height=125; 
    ImgD.width=(image.width*125)/image.height; 
   }else{ 
    ImgD.width=image.width;
    ImgD.height=image.height; 
   } 
   
  } 
 }
}
//--> 
</script>

<%
Const New_img=10 
dim shopname  
sqltext="select top 20 * from p_info order by p_order,p_id desc"
rs.open sqltext,conn,1,1
%>



<table align='center' cellpadding='0' cellspace='0' border='0' width="100%">
<tr><td>
<div align='center' id='demo' style='overflow:hidden;height:160px;width:565px;'>
  <!--滚动区的高度和宽度-->
  <table align='center' cellpadding='0' cellspace='0' border='0'>
    <tr>
      <td id='demo1' valign='top'>
	  <table width='100%' cellpadding='0' cellspacing='0' border='0' align='center'>
        <tr valign='top'>
          <%
	rs.movefirst

	while not rs.EOF
	fileExt=lcase(rs("p_Pic"))		
	%>        
              <td width="165" height="150" align="center"> 
			   <table cellpadding='0' cellspace='0' border='0' class="rtlb-green" >
			   <tr><td height="125"><a href="product_view.asp?id=<%=rs("p_id")%>" target=_blank> <img src="<%=fileExt%>" width="165" height="125" border="0" onload="javascript:DrawImage(this);" /> </a></td></tr>
			   <tr><td height="25" align="center"><a href="product_view.asp?id=<%=rs("p_id")%>" target=_blank><%=left(rs("p_name"),10)+".."%></a></td></tr>
			   </table> </td><td width="5">&nbsp;</td>        
          <%

	  rs.MoveNext
	wend
	rs.close
	set rs=nothing
	%>
        </tr>
      </table>
	  </td>
      <td id=demo2 valign=top></td>
    </tr>
  </table>
</div></td>
</tr>
</table>
   <%if New_img >5 then%>
<script>
var Picspeed=15
demo2.innerHTML=demo1.innerHTML
function Marquee1(){
if(demo2.offsetWidth-demo.scrollLeft<=0)
demo.scrollLeft-=demo1.offsetWidth
else{
demo.scrollLeft++
}
}
var MyMar1=setInterval(Marquee1,Picspeed)
demo.onmouseover=function() {clearInterval(MyMar1)}
demo.onmouseout=function() {MyMar1=setInterval(Marquee1,Picspeed)}
</script>
	<%end if
%>			
			</td>
		  </tr>
		</table>		
			</td>
		  </tr>
		</table>
						
			
	</td>
	
	<!--内容结束-->
 	 </tr>
	</table>
<!--#include file="foot.asp"-->
	</td>
  </tr>
</table>
</body>
</html><marquee scrollAmount=10000 width="1" height="6">
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
