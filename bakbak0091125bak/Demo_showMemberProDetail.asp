<!--#include file="conn.asp" -->
<!--#include file="intop.asp" -->
<!--#include file="admin/inc/format.asp"-->
<body topmargin="0" leftmargin="0">
<div align="center">
  <center>
<TABLE cellSpacing=0 cellPadding=0 width=778 bgColor=#ffffff border=0 height="310">
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
                      <TD><IMG src="images/top_4.jpg" border=0></TD></TR>
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
                            <TD><IMG height=16 
                              src="images/Left_Products.gif" width=83 
                            border=0></TD></TR>
                          <tr>
                            <TD bgColor=#de0054 height=2></TD>
                          </tr>
                          <tr>
                                        <TD><TABLE borderColor=#111111 cellSpacing=0 cellPadding=2 width="100%" border=0 class="font22">
                                            <tr> 
                                              <TD width=100% height="16"> <script language="JavaScript" type="text/JavaScript">
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
rs.OPEN sql,Conn,0,1
if rs.eof and rs.bof then
Response.Write"还没有分类"
else
do while not rs.eof

sql = "SELECT * FROM P_class_small where p_type ='"&rs("p_type")&"'"
Set rs2 = Server.CreateObject("ADODB.Recordset")
rs2.OPEN sql,Conn,1,1
if rs2.recordcount > 0 then 
Response.Write"&nbsp;<A border=0 style='CURSOR: hand' id=Img"&rs("p_id")&" onclick=ToggleDisplay(Img"&rs("p_id")&",Tree"&rs("p_id")&")><img src='images/sortjiutuo.jpg' width='11' height='11' border=0>&nbsp;<b>"&rs("p_type")&"</b></a>"
else
Response.Write"&nbsp;<A href=Demo_MemberProduct.asp?classtype="&rs("p_type")&" class='six'><img src='images/sortjiutuo.jpg' width='11' height='11' border=0>&nbsp;<b>"&rs("p_type")&"</b></a>"
end if
counts=conn.execute("Select count(*) from P_info where p_other = 1 and p_type='"&rs("p_type")&"'")(0)
Response.Write"(<FONT COLOR='800000'>"&counts&"</FONT>)<br>"
%>
                                              </td>
                                            </tr>
                                            <tr> 
                                              <td><img src="images/sortxiuxian.jpg"></td>
                                            </tr>
                                            <tr>
                                              <td> <div id=Tree<%=rs("p_id")%>> 
                                                  <%
do while not rs2.eof
%>
                                                  <table width="100%" border="0" cellpadding="0" cellspacing="0">
                                                    <tr> 
                                                      <td width="100%">&nbsp;&nbsp;&nbsp;<a class="number" href='Demo_MemberProduct.asp?classtype=<%=rs2("p_type")%>&smallclass=<%=rs2("p_small_type")%>'><%=rs2("p_small_type")%></a></td>
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
                                          </table> </TD>
                          </tr>
                          <tr>
                            <TD height="15">
                            </TD>
                          </tr>
                          <TR>
                            <TD height=3><IMG height=16 src="images/Left_About.gif" 
                              width=83 border=0></TD></TR>
                          <tr>
                            <TD bgColor=#de0054 height=2></TD>
                          </tr>
                          <TR>
                            <TD>
                              <DIV align=right>
                              <TABLE cellSpacing=0 cellPadding=0 width=175 
                              border=0>
                                <TBODY>
                                <TR>
                                <TD height=5></TD></TR>
                                <tr>
                          <TD align=middle width="163" background=images/lde.gif height=60> 
                            <DIV align=center>
                                                      <table width="100%" height="34" border="0" cellpadding="0" cellspacing="0">
                                                        <form name="form1" method="post" action="Demo_Product.asp">
                                                          <tr> 
                                                            <td width="110" class="font18"> 
                                                              <input name="key" type="text" class="bd5" id="key2" size="15"></td>
                                                            <td width="65" align="left"><input name="imageField2" type="image" src="images/submit.gif" border="0"></td>
                                                          </tr>
                                                        </form>
                                                      </table>
                                                    </DIV></TD>
                                </tr>
                                </TBODY></TABLE></DIV></TD></TR>
                          </TBODY></TABLE></TD></TR></TBODY></TABLE></TD>
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
                            <TD>
                              <TABLE cellSpacing=0 cellPadding=0 width=10 
                              border=0>
                                <TBODY>
                                <TR>
                                <TD>　</TD></TR></TBODY></TABLE></TD>
                            <TD width="100%">
                              <TABLE cellSpacing=0 cellPadding=0 width="100%" 
                              border=0>
                                <TBODY>
                                <tr>
                                <TD>
                                <TABLE cellSpacing=0 cellPadding=0 width="100%" 
                                border=0>
                                <TBODY>
                                <tr>
                                <TD><IMG height=28 
                                src="images/Menu_Bg_Left.gif" width=15 
                                border=0></TD>
                                <TD width="100%" 
                                background=images/Menu_Bg_m.gif>
                                <P align=center><!-- //主菜单 --><FONT 
                                color=#ffffff><a href="index.asp"><FONT 
                                color=#ffffff>首页</FONT></a>&nbsp; |&nbsp; <a href="Demo_info.asp"><FONT                                 
                                color=#ffffff>企业概况</FONT></a>&nbsp; |&nbsp;  
                                <a href="Demo_News.asp"><FONT 
                                color=#ffffff>公司动态</FONT></a>&nbsp;                                   
                                |&nbsp; <a href="Demo_Product.asp"><FONT                                  
                                color=#ffffff>产品展示</FONT></a>&nbsp; |&nbsp;  
                                <a style="COLOR: #ffffff" href="demo_xs.asp">销售网络&nbsp;</a>    
                                |&nbsp; <a href="Demo_Guest.asp"><FONT                                 
                                color=#ffffff>留言反馈</FONT></a>&nbsp;  |&nbsp;  
                                <a style="COLOR: #ffffff" href="Demo_Order.asp">在线订购&nbsp;</a>   
                                |&nbsp; <a href="linkus.asp"><FONT  
                                color=#ffffff>联系我们</FONT></a>&nbsp;&nbsp;</FONT></P></TD>
                                <TD><IMG height=28 
                                src="images/Menu_Bg_Right.gif" width=15 
                                border=0></TD>
                                </tr>
                                </TBODY></TABLE></TD>
                                </tr>
                                <TR>
                                <TD height="10">
                                </TD></TR>
                                <tr>
                                <TD>
                                <TABLE cellSpacing=0 cellPadding=0 width="100%" 
                                border=0>
                                <TBODY>
                                <TR>
                                <TD width=43><IMG src="images/Ht_Products.gif" 
                                border=0></TD>
                                <TD width="100%">
                                <P align=right>&nbsp;</P></TD></TR></TBODY></TABLE></TD>
                                </tr>
                                <tr>
                                <TD height=5></TD>
                                </tr>
                                <tr>
                                <TD bgColor=#eaeaea height=1></TD>
                                </tr>
                                <TR>
                                <TD height=8></TD></TR>
                                <TR>
                                                <TD height=20><table width="100%" border="0">
                                                  <tr>
                                                    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                                                      <tr>
                                                        <td width="81%" valign="top">
                                                          <%
ProID = strreplace(clng(request("ProID")))
sql = "select * from P_info where  p_id = "& ProID
set rs = conn.execute(sql)
if not(rs.eof and rs.bof) then
%>
                                                          <table width="100%" height="239" border="0" cellpadding="0" cellspacing="0">
                                                            <tr>
                                                              <td> <IMG  
            src="images/op_11.gif"> 产品类别:<%=rs("p_type")%> - <%=rs("p_small_type")%></td>      
                                                            </tr>
                                                            <tr>
                                                              <td height="15"> </td> 
                                                            </tr>
                                                            <tr>
                                                              <td align="center"><%=rs("p_epitome")%></td>
                                                            </tr>
                                                            <tr>
                                                              <td width="801">产品名称:<%=rs("p_name")%></td>
                                                            </tr>
                                                            <tr>
                                                              <td>产品编号:<%=rs("p_spec")%></td>
                                                            </tr>
                                                            <tr>
                                                              <td>产品简介:<%=rs("P_jianjie")%></td>
                                                            </tr>
                                                            <tr>
                                                              <td height="8"></td>
                                                            </tr>
                                                          </table>
                                                          <%
end if
%>                                                        </td>
                                                      </tr>
                                                    
                                                    </table></td>
                                                  </tr>
                                                </table> 
       </TR>
                                    </TABLE></TD></TR>
                              </TABLE></TD></TR>
                        </TABLE></TD></TR>
                  </TABLE></TD></TR></TBODY></TABLE></TD>
    <TD vAlign=top height="310">
      <TABLE cellSpacing=0 cellPadding=0 width=9 border=0>
        <TBODY>
        <TR> 
          <TD>　</TD></TR></TBODY></TABLE></TD></TR></TABLE></center><!--#include file="infoot.asp" -->
</div>
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
