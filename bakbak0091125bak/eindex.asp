<!--#include file="config.asp"-->
<!--#include file="eintop.asp" -->
<body topmargin="0" leftmargin="0">

<TABLE cellSpacing=0 cellPadding=0 width=778 bgColor=#ffffff border=0 height="310" align="center">
  <TR>
    <TD vAlign=top height="310">
      <TABLE cellSpacing=0 cellPadding=0 width=9 border=0>
        <TBODY>
        <TR>
          <TD>　</TD></TR></TBODY></TABLE></TD>
    <TD width="100%" height="310">
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0 height="321">
        <TBODY>
        <TR>
          <TD height="321">
            <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0 height="318">
              <TBODY>
              <TR>
                <TD vAlign=top width=61 height="318">
                  <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
                    <TBODY>
                    <TR>
                      <TD><IMG height=29 src="index/Top_Bn_Left.gif" 
                        width=190 border=0></TD></TR>
                    <TR>
                      <TD height=20><!-- //左栏菜单 -->
                        <TABLE cellSpacing=0 cellPadding=0 width="100%" 
border=0>
                          <TBODY>
                          <TR>
                            <TD height="22"></TD></TR></TBODY></TABLE>
                                  <TABLE cellSpacing=0 cellPadding=0 width="100%" 
border=0>
                                    <TBODY>
                                      <TR> 
                                        <TD><img border="0" src="images/eLeft_About.gif"></TD>
                                      </TR>
                                      <tr> 
                                        <TD bgColor=#0BBAFD height=2></TD>
                                      </tr>
                                      <tr> 
                                        <TD height="15"> <TABLE borderColor=#111111 cellSpacing=0 cellPadding=2 width="100%" border=0 class="font22">
              <tr>
                        <TD width=100% height="8"> 
</td>
              </tr>
              <tr> 
                        <TD width=100% height="16"> 
<script language="JavaScript" type="text/JavaScript">
		             <!--
            function MM_jumpMenu(targ,selObj,restore){ //v3.0
                      eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
                      if (restore) selObj.selectedIndex=0;
                     }
            //-->
<!--
	function ToggleDisplay(oButton, oItems){
		if ((oItems.style.display == "") || (oItems.style.display == "block"))	{
			oItems.style.display = "none";
			oButton.src = "/images/tree_plus.gif";
		}	
		else {
			oItems.style.display = "block";
			oButton.src = "/images/tree_minus.gif";
		}
		return false;
	}
	function UnFoldAll() {
		temp = document.all.length;
	    for(i=0;i<temp;i++)
		    if(document.all[i].id.substr(0,3) =="Img")
			    document.all[i].src = "/images/tree_minus.gif"
	        else if(document.all[i].id.substr(0,4) =="Tree")
		        document.all[i].style.display ="block"
	}
	function FoldAll() {
		temp = document.all.length;
	    for(i=0;i<temp;i++)
		    if(document.all[i].id.substr(0,3) =="Img")
			    document.all[i].src = "/images/tree_plus.gif"
	        else if(document.all[i].id.substr(0,4) =="Tree")
		        document.all[i].style.display ="none"
	}
//-->
</SCRIPT>

<%   
sql = "SELECT * FROM P_class ORDER BY p_id"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.OPEN sql,Conn,1,1
if rs.eof and rs.bof then
Response.Write"no product type"
else
do while not rs.eof

sql = "SELECT * FROM P_class_small where p_type ='"&rs("p_type")&"'"
Set rs2 = Server.CreateObject("ADODB.Recordset")
rs2.OPEN sql,Conn,1,1
if rs2.recordcount > 0 then 
Response.Write"&nbsp;<A border=0 style='CURSOR: hand' id=Img"&rs("p_id")&" onclick=ToggleDisplay(Img"&rs("p_id")&",Tree"&rs("p_id")&")><img src='images/sortjiutuo.jpg' width='11' height='11' border=0>&nbsp;<b>"&rs("p_type_e")&"</b></a>"
else
Response.Write"&nbsp;<A href='eProduct.asp?classtype="&rs("p_type_e")&"' class='six'><img src='images/sortjiutuo.jpg' width='11' height='11' border=0>&nbsp;<b>"&rs("p_type_e")&"</b></a>"
end if
counts=conn.execute("Select count(*) from P_info where p_other = 0 and p_type='"&rs("p_type")&"'")(0)
Response.Write"(<FONT COLOR='800000'>"&counts&"</FONT>)<br>"
%>
</td></tr>
<tr>
  <td><img border="0" src="index/line.gif"></td>
</tr>
<tr><td>
                          <div id=Tree<%=rs("p_id")%>> 
                            <%
do while not rs2.eof
%>
                            <table width="100%" border="0" cellpadding="0" cellspacing="0">
                              <tr> 
                                <td width="100%">&nbsp; ◇&nbsp;<a class="number" href='eProduct.asp?classtype=<%=rs2("p_type_e")%>&smallclass=<%=rs2("p_small_type_e")%>'><%=rs2("p_small_type_e")%></a></td>       
                              </tr>
                            </table>
                            <%
rs2.movenext
loop
rs2.close
Set rs2=Nothing
%>
                            <table width="100%" border="0" cellpadding="0" cellspacing="0">
                              <tr> 
                                <td width="100%" height="5"></td>
                              </tr>
                            </table>
                          </div>
                          <%
rs.movenext
loop
end if
rs.close
Set rs=Nothing
%>
<SCRIPT language=javascript>FoldAll();</SCRIPT>
                        </td>
                      </tr>
                    </table></TD>
                                      </tr>
                                      <TR> 
                                        <TD height=10></TD>
                                      </TR>
                                      <TR> 
                                        <TD height=3><!--<img border="0" src="images/2Left_About2.gif">--></TD>
                                      </TR>
                                      <tr> 
                                        <TD bgColor=#0BBAFD height=2></TD>
                                      </tr>
                                      <TR> 
                                        <TD> <DIV align=right> 
                                            <TABLE cellSpacing=0 cellPadding=0 width=175 
                              border=0>
                                              <TBODY>
                                                <TR> 
                                                  <TD height=5></TD>
                                                </TR>
                                                <tr> 
                                                  <TD align=middle width="163" height=60> 
                                                   <!-- <DIV align=center>
                                                      <table width="100%" height="34" border="0" cellpadding="0" cellspacing="0">
                                                        <form name="form1" method="post" action="Demo_eProduct.asp">
                                                            <td width="124" align="left"><input name="key" type="text" class="bd5" id="key" size="15"></td>
                                                            <td width="48" align="left"><input name="imageField" type="image" src="images/esubmit.gif" border="0"></td>
                                                        </form>
                                                      </table>
                                                    </DIV>--></TD>
                                                </tr>
                                              </TBODY>
                                            </TABLE>
                                          </DIV></TD>
                                      </TR>
                                    </TBODY>
                                  </TABLE></TD></TR></TBODY></TABLE></TD>
                <TD vAlign=top background=images/Jg_Left_Bg.gif height="318"><IMG 
                  src="images/Jg_Left.gif" border=0 width="14" height="218"></TD>
                <TD vAlign=top width="100%" height=318>
                  <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
                    <TBODY>
                    <TR>
                      <TD><IMG src="images/Top_Bn_Foot.gif" border=0></TD></TR>
                    <TR>
                      <TD>
                        <TABLE cellSpacing=0 cellPadding=0 width="100%" 
border=0>
                          <TBODY>
                          <TR>                            
                            <TD width="100%">
                              <TABLE cellSpacing=0 cellPadding=0 width="100%" 
                              border=0>
                                <TBODY>
                                <tr>
                                <TD>
                                <TABLE cellSpacing=0 cellPadding=0 width="100%" 
                                border=0>
                                <TBODY>
                                <TR>
                                <TD><IMG height=28 
                                src="images/Menu_Bg_Left.gif" width=15 
                                border=0></TD>
                                <TD width="100%" 
                                background=images/Menu_Bg_m.gif>
                                <P align=center><FONT 
                                color=#ffffff><a href="eindex.asp"><FONT 
                                color=#ffffff>Home</FONT></a>&nbsp; |&nbsp; <a href="eintro.asp"><FONT                                       
                                color=#ffffff>Company</FONT></a>&nbsp; |&nbsp; <a href="eNews.asp"><FONT      
                                color=#ffffff>News</FONT></a>&nbsp;                                          
                                |&nbsp; <a href="eProduct.asp"><FONT                                        
                                color=#ffffff>Products</FONT></a>&nbsp; |&nbsp; <a href="eGuest.asp"><FONT                                       
                                color=#ffffff>Messenger</FONT></a> |&nbsp; <a style="COLOR: #ffffff" href="eOrder.asp">Order</a>&nbsp;      
                                |&nbsp; <a href="Demo_Guest.asp"><FONT                                       
                                color=#ffffff></FONT></a><a href="elinkus.asp"><FONT 
                                color=#ffffff>Contact</FONT></a></FONT></P></TD>
                                <TD><IMG height=28 
                                src="images/Menu_Bg_Right.gif" width=15 
                                border=0></TD></TR></TBODY></TABLE></TD>
                                </tr>
                                <TR>
                                <TD height=10></TD></TR>
                                <TR>
                                                <TD height=25> 
                                                  <TABLE cellSpacing=0 cellPadding=0 width="100%" 
                                border=0>
                                                    <TBODY>
                                                      <tr>
                                <TD>
                                <TABLE cellSpacing=0 cellPadding=0 width="100%" 
                                border=0>
                                <TBODY>
                                <TR>
                                <TD width=43><img border="0" src="index/aboutus.gif"></TD>
                                <TD width="100%">
                                <P align=right>　</P></TD></TR></TBODY></TABLE></TD>
                                                      </tr>
                                                      <tr>
                                <TD height=5></TD>
                                                      </tr>
                                                      <tr>
                                <TD bgColor=#eaeaea height=1></TD>
                                                      </tr>
                                                    </TBODY>
                                                  </TABLE>
                                                 <table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
			  <tr>
				<td width="215"><img src="images/intro.jpg" width="185" height="140"></td>
				<td><table width="100%"  border="0" cellspacing="0" cellpadding="0">
				  <tr>
					<td class="zhengwen">
					<% 
'###############################info#############################
'根据后台简介管理里面的英文栏目名称修改i_type的值,比如是公司简介等,如果后台修改了类别名字,请在这里做相应的修改
sql="select inc_text from inc_info,inc_class where inc_class.i_id=inc_info.inc_class_id and i_type='About'"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if not rs.eof then
 response.Write(left(rs("inc_text"),300))
 end if
 rs.close
 set rs=nothing
 %> 				
<a href="eintro.asp">More&gt;&gt;</a></td>
				  </tr>
				</table></td>
			  </tr>
			</table><table width="100%"  border="0" cellspacing="0" cellpadding="0">
		  <tr>
			<td><img src="images/main_news.jpg" width="100%" height="45" border="0"></td>
		  </tr>
		  <tr>
			<td><table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
			  <tr>
				<td>
				<table width="100%"  border="0" align="left" cellpadding="0" cellspacing="0" class="zhengwen">
					   <% 
news_type = "News"
sql="select top 5 news.* from news,News_Class where news.news_class_id=news_class.news_id and news_language=1 and news_type = '"&news_type&"' order by news.news_id desc"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	if not rs.eof then
	  while not rs.eof
	%>	
	  <tr>
		<td height="26" class="b-1-gray">  &nbsp;<img src="images/doc.jpg" width="10" height="10">&nbsp;<a href="eShowNews.asp?news_id=<%=rs("news_id")%>"><%=left(getInnerText(rs("news_title")),16)%></a></td>
		<td width="80" align="right" class="b-1-gray"><%= rs("news_date") %></td>
	  </tr>
							  <% 
						rs.movenext
						wend
						rs.close
						set rs=nothing
						end if
						 %>		 
	</table></td>
				<td width="180" align="center"><img src="images/news.jpg" width="153" height="99"></td>
			  </tr>
			</table></td>
		  </tr>
		</table><table width="100%"  border="0" cellspacing="0" cellpadding="0">
		  <tr>
			<td><img src="images/main_pro.jpg" width="100%" height="45"></td>
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
  if(image.width/image.height>= 110/125){ 
   if(image.width>110){
    ImgD.width=110; 
    ImgD.height=(image.height*110)/image.width; 
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
set rs=server.createobject("adodb.recordset")
sqltext="select top 20 * from p_info order by p_id desc"
rs.open sqltext,conn,1,1
%>



<table align='center' cellpadding='0' cellspace='0' border='0' width="100%">
<tr><td>
<div align='center' id='demo' style='overflow:hidden;height:160px;width:550px;'>
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
              <td width="110" height="150" align="center"> 
			   <table cellpadding='0' cellspace='0' border='0' class="rtlb-green" >
			   <tr><td height="125"><a href="eshowProDetail.asp?ProID=<%=rs("p_id")%>" target=_blank> <img src="<%=fileExt%>" width="110" height="125" border="0" onload="javascript:DrawImage(this);" /> </a></td></tr>
			   <tr><td height="25" align="center"><a href="eshowProDetail.asp?ProID=<%=rs("p_id")%>" target=_blank><%=left(rs("p_name_e"),10)+".."%></a></td></tr>
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
		</table> <table width="100%" border="0" cellpadding="0" cellspacing="0">
                                                    <tr>
                                  <TD height=5>								  
                                    
                                  </TD>
                                                    </tr>
                                                  </table>
                                                </TD>
                                    </TR>
                                    </TABLE></TD></TR>
                              </TABLE></TD></TR>
                        </TABLE></TD></TR>
                  </TABLE></TD></TR></TBODY></TABLE></TD>
    <TD vAlign=top height="310">
    </TD></TR></TABLE>
    <!--#include file="einfoot.asp" -->

</BODY>
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
